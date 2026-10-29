import os
import json
import time
import base64
import socket
import asyncio
from urllib.parse import urlparse
from typing import Optional, Dict, List, Any, Tuple
import httpx
from dotenv import load_dotenv

from overdue_calculator import (
    calculate_overdue_periods,
    is_suspect_case,
    parse_registration_timeline,
    filter_overdue_by_timeline
)

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_M33_DIR = os.path.abspath(os.path.join(_THIS_DIR, "..", "SSO_M33_Overdue_System"))

# Load environment variables (local .env first, fallback to SSO_M33_Overdue_System/.env)
load_dotenv(os.path.join(_THIS_DIR, ".env"))
if os.path.exists(os.path.join(_M33_DIR, ".env")):
    load_dotenv(os.path.join(_M33_DIR, ".env"), override=False)

SSO_USERNAME = os.getenv("SSO_USERNAME", "pchalalai")
SSO_PASSWORD = os.getenv("SSO_PASSWORD", "iwi@4284")
BASE_URL = "https://appssocore.sso.go.th"

CACHE_FILE = os.path.join(_THIS_DIR, ".sso_token_cache.json")
FALLBACK_CACHE_FILE = os.path.join(_M33_DIR, ".sso_token_cache.json")

INTRANET_UNREACHABLE_MSG = (
    "ไม่สามารถเชื่อมต่อระบบ SSO CORE ได้ เนื่องจากเซิร์ฟเวอร์ไม่ได้อยู่ในเครือข่ายอินทราเน็ต สปส. "
    "(appssocore.sso.go.th) ระบบจะยึดข้อมูลงวดค้างชำระตามไฟล์ Excel เป็นหลัก"
)


def check_intranet_reachable(timeout: float = 3.0) -> bool:
    """ตรวจสอบว่าเครื่องปัจจุบันสามารถเชื่อมต่อโครงข่ายอินทราเน็ต สปส. (appssocore.sso.go.th) ได้หรือไม่"""
    parsed = urlparse(BASE_URL)
    host = parsed.hostname or "appssocore.sso.go.th"
    port = parsed.port or (443 if parsed.scheme == "https" else 80)
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def _fetch_token_sync(username: str, password: str) -> str:
    if not check_intranet_reachable():
        raise ConnectionError(INTRANET_UNREACHABLE_MSG)

    from playwright.sync_api import sync_playwright
    captured_token = None
    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-web-security"]
            )
        except Exception as e:
            err_str = str(e)
            if "Executable doesn't exist" in err_str:
                raise RuntimeError(
                    "ไม่พบเบราว์เซอร์ Chromium สำหรับยืนยันตัวตน SSO CORE บนเซิร์ฟเวอร์ (กรุณารัน playwright install chromium)"
                ) from e
            raise
        context = browser.new_context(
            viewport={"width": 1440, "height": 900},
            ignore_https_errors=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        def handle_request(req):
            nonlocal captured_token
            auth = req.headers.get("authorization", "")
            if auth.startswith("Bearer ") and not captured_token:
                token_str = auth.replace("Bearer ", "").strip()
                if len(token_str) > 50:
                    captured_token = token_str

        page.on("request", handle_request)

        # 1) Open SSO login
        page.goto(f"{BASE_URL}/", timeout=60000)
        page.wait_for_selector("#username", timeout=45000)
        page.fill("#username", username)
        page.fill("#password", password)
        page.click("#kc-login")

        # 2) Wait for post-login redirect
        try:
            page.wait_for_url("**/appssocore.sso.go.th/**", timeout=30000)
        except Exception:
            pass
        page.wait_for_timeout(3000)

        # If token not intercepted yet, visit contribute-web
        if not captured_token:
            try:
                page.goto(f"{BASE_URL}/contribute-web/m33/contribution", timeout=40000)
                page.wait_for_timeout(4000)
            except Exception:
                pass

        # If still not captured, inspect storage
        if not captured_token:
            storage_tokens = page.evaluate("""() => {
                const tokens = [];
                for (let i = 0; i < sessionStorage.length; i++) {
                    const k = sessionStorage.key(i);
                    const v = sessionStorage.getItem(k);
                    if (v && v.includes('access_token')) {
                        try {
                            const parsed = JSON.parse(v);
                            if (parsed.access_token) tokens.push(parsed.access_token);
                        } catch(e) {}
                    }
                }
                for (let i = 0; i < localStorage.length; i++) {
                    const k = localStorage.key(i);
                    const v = localStorage.getItem(k);
                    if (v && v.includes('access_token')) {
                        try {
                            const parsed = JSON.parse(v);
                            if (parsed.access_token) tokens.push(parsed.access_token);
                        } catch(e) {}
                    }
                }
                return tokens;
            }""")
            if storage_tokens:
                captured_token = storage_tokens[0]

        browser.close()

    if not captured_token:
        raise RuntimeError("Failed to acquire SSO Bearer token during login.")
    return captured_token



def tls_verify_setting():
    """TLS verification for SSO API calls. Always on.
    Optional: set SSO_CA_BUNDLE to a CA bundle path if SSO uses a private CA."""
    ca = os.environ.get("SSO_CA_BUNDLE", "").strip()
    if ca:
        if not os.path.isfile(ca):
            raise RuntimeError(f"SSO_CA_BUNDLE not found: {ca}")
        return ca
    return True

class SSOHybridClient:
    def __init__(self, username: Optional[str] = None, password: Optional[str] = None):
        self.username = username or SSO_USERNAME
        self.password = password or SSO_PASSWORD
        self.token: Optional[str] = None
        self.token_expiry: float = 0
        self._lock = asyncio.Lock()
        self._load_cached_token()

    def _load_cached_token(self):
        for candidate in (CACHE_FILE, FALLBACK_CACHE_FILE):
            try:
                if os.path.exists(candidate):
                    with open(candidate, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        tok = data.get("token")
                        exp = data.get("expiry", 0)
                        if tok and time.time() < (exp - 60):
                            self.token = tok
                            self.token_expiry = exp
                            return
            except Exception as e:
                print(f"[SSO Auth] Cache load warning ({candidate}): {e}")

    def _save_cached_token(self, token: str, expiry: float):
        payload = {"token": token, "expiry": expiry}
        for target in (CACHE_FILE, FALLBACK_CACHE_FILE):
            try:
                if os.path.isdir(os.path.dirname(target)):
                    with open(target, "w", encoding="utf-8") as f:
                        json.dump(payload, f)
            except Exception as e:
                print(f"[SSO Auth] Cache save warning ({target}): {e}")

    def _is_token_valid(self) -> bool:
        if not self.token:
            return False
        return time.time() < (self.token_expiry - 60)

    def _parse_jwt_expiry(self, token: str) -> float:
        try:
            payload_part = token.split(".")[1]
            payload_part += "=" * ((4 - len(payload_part) % 4) % 4)
            decoded = base64.urlsafe_b64decode(payload_part)
            payload = json.loads(decoded)
            return float(payload.get("exp", time.time() + 300))
        except Exception:
            return time.time() + 300

    async def get_token(self, force_refresh: bool = False) -> str:
        async with self._lock:
            if not await asyncio.to_thread(check_intranet_reachable):
                raise ConnectionError(INTRANET_UNREACHABLE_MSG)

            if not force_refresh and self._is_token_valid():
                return self.token

            if not force_refresh:
                self._load_cached_token()
                if self._is_token_valid():
                    return self.token

            print("[SSO Auth] Authenticating with SSO Core to acquire JWT Bearer token...")
            captured_token = await asyncio.to_thread(_fetch_token_sync, self.username, self.password)

            self.token = captured_token
            self.token_expiry = self._parse_jwt_expiry(captured_token)
            self._save_cached_token(self.token, self.token_expiry)
            return self.token

    async def _make_api_request(self, method: str, path: str, json_data: Optional[Dict] = None, params: Optional[Dict] = None) -> Any:
        token = await self.get_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
        url = f"{BASE_URL}{path}"

        async with httpx.AsyncClient(verify=tls_verify_setting(), timeout=30.0) as client:
            resp = await client.request(method, url, json=json_data, params=params, headers=headers)

            if resp.status_code == 401:
                token = await self.get_token(force_refresh=True)
                headers["Authorization"] = f"Bearer {token}"
                resp = await client.request(method, url, json=json_data, params=params, headers=headers)

            if resp.status_code != 200:
                raise RuntimeError(f"SSO API Error [{resp.status_code}]: {resp.text}")

            return resp.json()

    async def search_company(self, company_no: str) -> Optional[Dict]:
        clean_acc = str(company_no).strip().zfill(10)
        payload = {
            "company": clean_acc,
            "paginable": {"rowPerPage": 20, "currentPage": 1},
            "ssoBranchCode": None
        }
        try:
            data = await self._make_api_request("POST", "/con-master-inq-manager/search-company", json_data=payload)
            res = data.get("result") or {}
            comp_list = res.get("companyList") or []
            if comp_list:
                return comp_list[0]
            if res.get("companyId"):
                return res
        except Exception as e:
            print(f"[SSO API] search_company warning: {e}")
        return None

    async def search_branch(self, company_no: str, branch_code: str = "000000") -> Optional[Dict]:
        clean_acc = str(company_no).strip().zfill(10)
        clean_br = str(branch_code).strip().zfill(6)
        payload = {
            "companyNo": clean_acc,
            "companyBranch": clean_br,
            "paginable": {"rowPerPage": 20, "currentPage": 1}
        }
        try:
            data = await self._make_api_request("POST", "/con-master-inq-manager/search-branch", json_data=payload)
            res = data.get("result") or {}
            branch_list = res.get("branchList") or []
            if branch_list:
                for b in branch_list:
                    if str(b.get("companyBranchCode")).strip().zfill(6) == clean_br:
                        return b
                return branch_list[0]
        except Exception as e:
            print(f"[SSO API] search_branch warning: {e}")
        return None

    async def get_contribution_transactions(
        self,
        company_id: int,
        start_year: int = 2500,
        end_year: int = 2569,
        row_per_page: int = 100
    ) -> Tuple[Dict, List[Dict]]:
        payload = {
            "companyId": company_id,
            "startPeriodMonth": "01",
            "startPeriodYear": str(start_year),
            "endPeriodMonth": "12",
            "endPeriodYear": str(end_year),
            "paymentStatus": None,
            "paginable": {
                "rowPerPage": row_per_page,
                "currentPage": 1,
                "orderBy": "period",
                "orderDirection": "DESC",
                "offset": 0
            }
        }
        data = await self._make_api_request("POST", "/con-m33-inq-manager/inquiry/search-transaction-v3", json_data=payload)
        res = data.get("result") or {}
        return res.get("company") or {}, res.get("contributionList") or []

    async def get_company_registration_timeline(self, company_no: str, branch_code: str = "000000") -> Tuple[Optional[Dict], List[Dict]]:
        clean_acc = str(company_no).strip().zfill(10)
        clean_br = str(branch_code).strip().zfill(6)
        payload = {
            "no": clean_acc,
            "branchno": clean_br,
            "paging": 1,
            "rowperpage": 20,
            "orderBy": "OPENNED_DATE",
            "orderType": "desc"
        }
        try:
            res = await self._make_api_request("POST", "/reg-company-manager-service/api/register/getregistertblist", json_data=payload)
            items = res.get("data") or []
            target_comp_id = None
            target_item = None
            for it in items:
                sso_no = str(it.get("ssoCompanyNo") or "").strip().zfill(10)
                sso_br = str(it.get("ssoBranchNo") or "").strip().zfill(6)
                if sso_no == clean_acc and sso_br == clean_br:
                    target_comp_id = it.get("companyId")
                    target_item = it
                    break

            if not target_comp_id and items:
                target_comp_id = items[0].get("companyId")
                target_item = items[0]

            if target_comp_id:
                timeline_res = await self._make_api_request("POST", "/reg-company-manager-service/api/register/getcompanytimeline", json_data={"companyId": target_comp_id})
                timeline_data = timeline_res.get("data") or []
                return target_item, timeline_data

        except Exception as e:
            print(f"[SSO Reg API] get_company_registration_timeline warning for {clean_acc}/{clean_br}: {e}")

        return None, []

    async def check_account_overdue(
        self,
        company_no: str,
        branch_code: str = "000000",
        start_year: int = 2568,
        as_of_date: Optional[Any] = None
    ) -> Dict[str, Any]:
        clean_acc = str(company_no).strip().zfill(10)
        clean_br = str(branch_code).strip().zfill(6)

        try:
            branch_info = await self.search_branch(clean_acc, clean_br)
            company_id = branch_info.get("companyId") if branch_info else None
            comp_search = None

            if not company_id:
                comp_search = await self.search_company(clean_acc)
                if comp_search:
                    company_id = comp_search.get("companyId")

            if not company_id:
                return {
                    "clean_account_no": clean_acc,
                    "clean_branch_code": clean_br,
                    "company_name": "-",
                    "company_status": "ไม่พบบัญชี",
                    "employee_count": "-",
                    "sso_branch": "-",
                    "latest_paid_period": "-",
                    "overdue_count": 0,
                    "overdue_periods": [],
                    "overdue_periods_str": "-",
                    "current_period_str": "-",
                    "is_suspect": False,
                    "suspect_reason": "-",
                    "reg_checked": False,
                    "reg_first_open_date": "-",
                    "reg_close_date": "-",
                    "reg_reopen_date": "-",
                    "reg_timeline_summary": "-",
                    "process_status": "ไม่พบข้อมูลสถานประกอบการในระบบ SSO"
                }

            comp_full, tx_list = await self.get_contribution_transactions(company_id, start_year=start_year - 2)

            raw_comp_name = comp_full.get("companyName") or (comp_search.get("companyName") if comp_search else "-")
            branch_name = comp_full.get("companyBranchName") or (branch_info.get("companyBranchName") if branch_info else "")

            if clean_br == "000000":
                company_name = raw_comp_name or branch_name
            else:
                if branch_name and branch_name != raw_comp_name:
                    company_name = f"{branch_name} ({raw_comp_name})"
                else:
                    company_name = branch_name or raw_comp_name

            status_code = comp_full.get("companyStatus") or (branch_info.get("companyBranchStatus") if branch_info else "-")
            status_map = {
                "A": "A - เป็นสมาชิก",
                "C": "C - ยกเลิกกิจการ",
                "H": "H - นายจ้างขึ้นทะเบียนรอตรวจสอบ",
                "I": "I - ขึ้นทะเบียนกองทุนประกันสังคม",
                "L": "L - ยกเลิกตามกฎหมาย",
                "M": "M - นายจ้างแจ้งยกเลิกสาขา",
                "S": "S - หยุดกิจการชั่วคราว",
                "U": "U - ไม่มีการจ้าง",
                "W": "W - เป็นสมาชิกเฉพาะกองทุนเงินทดแทน"
            }
            company_status = status_map.get(status_code, status_code)
            employee_count = comp_full.get("numOfEmployee", "-")
            sso_branch = comp_full.get("fullBranch") or comp_full.get("ssoBranchName", "-")

            paid_periods = []
            for tx in tx_list:
                p = tx.get("period")
                if p and (tx.get("payStatus") == "0" or tx.get("receiptNo")):
                    paid_periods.append(p)

            calc_res = calculate_overdue_periods(
                paid_periods=paid_periods,
                start_year=start_year,
                start_month=1,
                as_of_date=as_of_date
            )

            is_suspect, suspect_reason = is_suspect_case(calc_res["overdue_periods"])
            reg_checked = False
            reg_first_open_date = "-"
            reg_close_date = "-"
            reg_reopen_date = "-"
            reg_timeline_summary = "-"

            if is_suspect:
                reg_item, timeline_data = await self.get_company_registration_timeline(clean_acc, clean_br)
                if timeline_data:
                    reg_parsed = parse_registration_timeline(timeline_data)
                    reg_checked = True
                    reg_first_open_date = reg_parsed.get("first_open_date", "-")
                    reg_close_date = reg_parsed.get("close_date", "-")
                    reg_reopen_date = reg_parsed.get("reopen_date", "-")
                    reg_timeline_summary = reg_parsed.get("timeline_summary", "-")

                    if reg_parsed.get("active_intervals"):
                        filtered = filter_overdue_by_timeline(calc_res["overdue_periods"], reg_parsed["active_intervals"])
                        calc_res["overdue_periods"] = filtered
                        calc_res["overdue_count"] = len(filtered)
                        calc_res["overdue_periods_str"] = ", ".join(filtered) if filtered else "-"
                        calc_res["status"] = "ค้างชำระ" if len(filtered) > 0 else "ปกติ"

            return {
                "clean_account_no": clean_acc,
                "clean_branch_code": clean_br,
                "company_name": company_name,
                "company_status": company_status,
                "employee_count": employee_count,
                "sso_branch": sso_branch,
                "latest_paid_period": calc_res["latest_paid_period"],
                "overdue_count": calc_res["overdue_count"],
                "overdue_periods": calc_res["overdue_periods"],
                "overdue_periods_str": calc_res["overdue_periods_str"],
                "current_period_str": calc_res["current_period_str"],
                "is_suspect": is_suspect,
                "suspect_reason": suspect_reason if is_suspect else "-",
                "reg_checked": reg_checked,
                "reg_first_open_date": reg_first_open_date,
                "reg_close_date": reg_close_date,
                "reg_reopen_date": reg_reopen_date,
                "reg_timeline_summary": reg_timeline_summary,
                "process_status": "ตรวจสอบสำเร็จ"
            }

        except Exception as e:
            return {
                "clean_account_no": clean_acc,
                "clean_branch_code": clean_br,
                "company_name": "-",
                "company_status": "เกิดข้อผิดพลาด",
                "employee_count": "-",
                "sso_branch": "-",
                "latest_paid_period": "-",
                "overdue_count": 0,
                "overdue_periods": [],
                "overdue_periods_str": "-",
                "current_period_str": "-",
                "is_suspect": False,
                "suspect_reason": "-",
                "reg_checked": False,
                "reg_first_open_date": "-",
                "reg_close_date": "-",
                "reg_reopen_date": "-",
                "reg_timeline_summary": "-",
                "process_status": f"เกิดข้อผิดพลาด: {repr(e)}"
            }

