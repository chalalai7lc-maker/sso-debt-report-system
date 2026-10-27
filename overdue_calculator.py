import re
from datetime import date
from typing import List, Dict, Optional, Tuple, Any

def normalize_period(period_str: str) -> str:
    """Normalizes period strings like '7/2568', ' 07/2568 ' to '07/2568'"""
    if not period_str:
        return ""
    clean = str(period_str).strip()
    match = re.match(r"^(\d{1,2})/(\d{4})$", clean)
    if match:
        m, y = match.groups()
        return f"{int(m):02d}/{int(y)}"
    return clean

def period_to_sort_key(period_str: str) -> Tuple[int, int]:
    """Converts 'MM/YYYY' into (YYYY, MM) for chronological sorting"""
    norm = normalize_period(period_str)
    try:
        parts = norm.split("/")
        return (int(parts[1]), int(parts[0]))
    except Exception:
        return (0, 0)

def parse_iso_date_to_be(date_str: Optional[str]) -> Optional[Tuple[int, int, int]]:
    """
    Parses 'YYYY-MM-DD' or 'YYYY-MM-DDTHH:MM:SS' into (year_be, month, day).
    Converts CE year to Buddhist Era (BE = CE + 543).
    """
    if not date_str:
        return None
    try:
        clean = str(date_str).strip()[:10]
        parts = clean.split("-")
        y = int(parts[0]) + 543
        m = int(parts[1])
        d = int(parts[2])
        return (y, m, d)
    except Exception:
        return None

def format_be_date_str(date_str: Optional[str]) -> str:
    """Formats ISO date string to 'DD/MM/YYYY' in Thai Buddhist Era"""
    be_tuple = parse_iso_date_to_be(date_str)
    if not be_tuple:
        return "-"
    y, m, d = be_tuple
    return f"{d:02d}/{m:02d}/{y}"

def is_suspect_case(overdue_periods: List[str]) -> Tuple[bool, str]:
    """
    Checks if an account meets the criteria to query the registration database:
    1. Overdue count >= 8 months (ค้างชำระ >= 8 งวด)
    2. Has any overdue period before year 2567 (< 2567: 2566, 2565, 2564, ...)
    """
    if not overdue_periods:
        return False, ""

    pre_2567 = []
    for p in overdue_periods:
        norm = normalize_period(p)
        if "/" in norm:
            try:
                y = int(norm.split("/")[1])
                if y < 2567:
                    pre_2567.append(norm)
            except Exception:
                pass

    is_ge_8 = len(overdue_periods) >= 8
    has_pre_2567 = len(pre_2567) > 0

    if is_ge_8 and has_pre_2567:
        sample_pre = ", ".join(pre_2567[:3]) + ("..." if len(pre_2567) > 3 else "")
        return True, f"ค้างชำระ {len(overdue_periods)} งวด (>= 8 งวด) และมีงวดก่อนปี 2567 ({sample_pre})"
    elif is_ge_8:
        return True, f"ค้างชำระ {len(overdue_periods)} งวด (>= 8 งวด)"
    elif has_pre_2567:
        sample_pre = ", ".join(pre_2567[:3]) + ("..." if len(pre_2567) > 3 else "")
        return True, f"มีงวดค้างก่อนปี 2567 ({sample_pre})"

    return False, ""

def parse_registration_timeline(timeline_data: Optional[List[Dict]]) -> Dict:
    """
    Parses timeline records from /getcompanytimeline into structured registration metadata
    and active obligation intervals.
    """
    if not timeline_data:
        return {
            "first_open_date": "-",
            "close_date": "-",
            "reopen_date": "-",
            "timeline_summary": "-",
            "has_timeline": False,
            "active_intervals": []
        }

    all_open_dates = []
    all_close_dates = []
    intervals = []

    for item in timeline_data:
        if not isinstance(item, dict):
            continue
        op = item.get("openDate")
        cl = item.get("closeDate")
        if op:
            all_open_dates.append(str(op)[:10])
        if cl:
            all_close_dates.append(str(cl)[:10])

        op_tuple = parse_iso_date_to_be(op) if op else None
        cl_tuple = parse_iso_date_to_be(cl) if cl else None

        if op_tuple and cl_tuple:
            if str(op)[:10] <= str(cl)[:10]:
                intervals.append({"open": op_tuple, "close": cl_tuple})
            else:
                intervals.append({"open": op_tuple, "close": None})
                intervals.append({"open": None, "close": cl_tuple})
        elif op_tuple and not cl_tuple:
            intervals.append({"open": op_tuple, "close": None})
        elif cl_tuple and not op_tuple:
            intervals.append({"open": None, "close": cl_tuple})

    sorted_open = sorted(list(set(all_open_dates)))
    sorted_close = sorted(list(set(all_close_dates)))

    first_open_str = format_be_date_str(sorted_open[0]) if sorted_open else "-"
    latest_close_str = format_be_date_str(sorted_close[-1]) if sorted_close else "-"

    reopen_str = "-"
    if sorted_close and sorted_open:
        latest_open = sorted_open[-1]
        latest_close = sorted_close[-1]
        if latest_open > latest_close:
            reopen_str = format_be_date_str(latest_open)
    elif len(sorted_open) > 1:
        reopen_str = format_be_date_str(sorted_open[-1])

    parts = []
    if first_open_str != "-":
        parts.append(f"เปิดครั้งแรก: {first_open_str}")
    if latest_close_str != "-":
        parts.append(f"วันปิด: {latest_close_str}")
    if reopen_str != "-":
        parts.append(f"เปิดใหม่: {reopen_str}")
    if not sorted_close and first_open_str != "-":
        parts.append("(ยังไม่เคยปิด)")

    summary = " | ".join(parts) if parts else "มีข้อมูลในฐานทะเบียน"

    return {
        "first_open_date": first_open_str,
        "close_date": latest_close_str,
        "reopen_date": reopen_str,
        "timeline_summary": summary,
        "has_timeline": True,
        "active_intervals": intervals
    }

def is_month_active(year_be: int, month: int, active_intervals: List[Dict]) -> bool:
    """
    Checks whether a specific (year_be, month) falls within an active obligation interval.
    """
    if not active_intervals:
        return True

    for interval in active_intervals:
        op = interval.get("open")
        cl = interval.get("close")

        if op:
            op_y, op_m, _ = op
            if year_be < op_y or (year_be == op_y and month < op_m):
                continue

        if cl:
            cl_y, cl_m, cl_d = cl
            if cl_d == 1:
                if year_be > cl_y or (year_be == cl_y and month >= cl_m):
                    continue
            else:
                if year_be > cl_y or (year_be == cl_y and month > cl_m):
                    continue

        return True

    return False

def filter_overdue_by_timeline(overdue_periods: List[str], active_intervals: List[Dict]) -> List[str]:
    """
    Filters overdue periods against active registration intervals.
    """
    if not active_intervals:
        return overdue_periods

    filtered = []
    for p in overdue_periods:
        norm = normalize_period(p)
        if "/" in norm:
            try:
                parts = norm.split("/")
                m = int(parts[0])
                y = int(parts[1])
                if is_month_active(y, m, active_intervals):
                    filtered.append(norm)
            except Exception:
                filtered.append(norm)
        else:
            filtered.append(norm)
    return filtered

def get_latest_due_period(as_of_date: date) -> Tuple[int, int]:
    """
    Returns (month, year_be) of the latest period whose due date has already passed.
    Due date rule: Period M/Y is due on the 15th of the following month (M+1)/Y.
    """
    curr_year_be = as_of_date.year + 543
    curr_month = as_of_date.month
    curr_day = as_of_date.day

    if curr_day >= 15:
        m = curr_month - 1
        y = curr_year_be
        if m < 1:
            m = 12
            y -= 1
        return (m, y)
    else:
        m = curr_month - 2
        y = curr_year_be
        if m < 1:
            m += 12
            y -= 1
        return (m, y)

def generate_period_sequence(start_year: int, start_month: int, end_year: int, end_month: int) -> List[str]:
    """Generates consecutive 'MM/YYYY' periods from start to end (inclusive)"""
    periods = []
    y = start_year
    m = start_month
    while (y < end_year) or (y == end_year and m <= end_month):
        periods.append(f"{m:02d}/{y}")
        m += 1
        if m > 12:
            m = 1
            y += 1
    return periods

def calculate_overdue_periods(
    paid_periods: List[str],
    start_year: int = 2568,
    start_month: int = 1,
    as_of_date: Optional[date] = None,
    timeline_data: Optional[List[Dict]] = None
) -> Dict[str, Any]:
    """
    Evaluates paid periods against due periods starting from start_year/start_month.
    If timeline_data is provided, filters overdue periods using active registration intervals.
    """
    if as_of_date is None:
        as_of_date = date.today()

    curr_year_be = as_of_date.year + 543
    curr_month = as_of_date.month

    current_period = f"{curr_month:02d}/{curr_year_be}"
    next_m = curr_month + 1
    next_y = curr_year_be
    if next_m > 12:
        next_m = 1
        next_y += 1
    current_period_due_date = f"15/{next_m:02d}/{next_y}"

    normalized_paid = set()
    latest_paid_period = "-"
    if paid_periods:
        sorted_paid = sorted([p for p in paid_periods if p], key=period_to_sort_key, reverse=True)
        if sorted_paid:
            latest_paid_period = normalize_period(sorted_paid[0])
        for p in paid_periods:
            norm = normalize_period(p)
            if norm:
                normalized_paid.add(norm)

    if current_period in normalized_paid:
        current_period_str = f"{current_period} (ชำระแล้ว)"
    else:
        current_period_str = f"{current_period} กำหนดชำระ {current_period_due_date}"

    effective_start_year = start_year
    effective_start_month = start_month

    end_m, end_y = get_latest_due_period(as_of_date)

    if (effective_start_year > end_y) or (effective_start_year == end_y and effective_start_month > end_m):
        expected_periods = []
    else:
        expected_periods = generate_period_sequence(effective_start_year, effective_start_month, end_y, end_m)

    overdue_periods = [p for p in expected_periods if p not in normalized_paid]

    parsed_tl = None
    if timeline_data is not None:
        parsed_tl = parse_registration_timeline(timeline_data)
        if parsed_tl.get("has_timeline") and parsed_tl.get("active_intervals"):
            overdue_periods = filter_overdue_by_timeline(overdue_periods, parsed_tl["active_intervals"])

    overdue_count = len(overdue_periods)
    overdue_periods_str = ", ".join(overdue_periods) if overdue_periods else "-"
    status = "ค้างชำระ" if overdue_count > 0 else "ปกติ"

    return {
        "status": status,
        "latest_paid_period": latest_paid_period,
        "overdue_count": overdue_count,
        "overdue_periods": overdue_periods,
        "overdue_periods_str": overdue_periods_str,
        "current_period": current_period,
        "current_period_due_date": current_period_due_date,
        "current_period_str": current_period_str,
        "evaluated_range": f"{start_month:02d}/{start_year} - {end_m:02d}/{end_y}",
        "parsed_timeline": parsed_tl
    }

