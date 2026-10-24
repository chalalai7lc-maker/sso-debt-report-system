/**
 * หน้าสรุป + พิมพ์หนังสือเตือน/เชิญ
 * เลขหนังสือต้องตรงกับ letter_service.assign_numbers
 */
document.addEventListener("DOMContentLoaded", () => {
    const section = document.getElementById("letter-section");
    const wrap = document.getElementById("letter-table-wrap");
    const summary = document.getElementById("letter-summary-bar");
    const statusMsg = document.getElementById("letter-status-msg");
    const issueInput = document.getElementById("letter-issue-date");
    const apptInput = document.getElementById("letter-appt-date");
    const warnInput = document.getElementById("letter-warn-start");
    const inviteInput = document.getElementById("letter-invite-start");
    let records = [];
    let selected = new Set();
    let apptTouched = false;

    const THAI_MONTHS = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];

    function formatPeriodsLetter(list) {
        const parsed = [];
        (list || []).forEach((raw) => {
            const parts = String(raw).trim().split("/");
            if (parts.length !== 2) return;
            const month = Number(parts[0]);
            const year = Number(parts[1]);
            if (month >= 1 && month <= 12 && Number.isInteger(year)) parsed.push([year, month]);
        });
        parsed.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
        const unique = [];
        parsed.forEach((item) => {
            const prev = unique[unique.length - 1];
            if (!prev || prev[0] !== item[0] || prev[1] !== item[1]) unique.push(item);
        });
        if (!unique.length) return "—";
        const groups = [[unique[0]]];
        unique.slice(1).forEach((item) => {
            const prev = groups[groups.length - 1].at(-1);
            if ((item[0] - prev[0]) * 12 + (item[1] - prev[1]) === 1) groups[groups.length - 1].push(item);
            else groups.push([item]);
        });
        return groups.map((group) => {
            const [startYear, startMonth] = group[0];
            const [endYear, endMonth] = group[group.length - 1];
            if (group.length === 1) return `${THAI_MONTHS[startMonth]} ${startYear}`;
            if (startYear === endYear) return `${THAI_MONTHS[startMonth]}-${THAI_MONTHS[endMonth]} ${endYear}`;
            return `${THAI_MONTHS[startMonth]} ${startYear}-${THAI_MONTHS[endMonth]} ${endYear}`;
        }).join(", ");
    }

    function escapeHtml(str) {
        return String(str ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    function digits(value) {
        return String(value ?? "").trim().replace(/\.0$/, "").replace(/\D/g, "");
    }

    function rowKey(row) {
        return `${digits(row["เลขที่บัญชี"])}|${digits(row["สาขา"]) || "0"}`;
    }

    function displayName(row) {
        return row["ชื่อสถานประกอบการ (จริง)"] || row["ชื่อในรายงานหนี้"] || "สถานประกอบการ";
    }

    function assignNumbers(count, warnStart, inviteStart) {
        const rows = [];
        for (let i = 0; i < count; i++) {
            if (inviteStart == null) {
                rows.push({ warn: warnStart + 2 * i, invite: warnStart + 2 * i + 1 });
            } else {
                rows.push({ warn: warnStart + i, invite: inviteStart + i });
            }
        }
        return rows;
    }

    function formatDocNo(n) {
        return String(n);
    }

    function isoToBe(iso) {
        if (!iso) return "";
        const [y, m, d] = iso.split("-").map(Number);
        if (!y || !m || !d) return "";
        return `${d} ${THAI_MONTHS[m]} ${y + 543}`;
    }

    function addDays(iso, days) {
        if (!iso) return "";
        const date = new Date(`${iso}T00:00:00`);
        date.setDate(date.getDate() + days);
        const m = String(date.getMonth() + 1).padStart(2, "0");
        const d = String(date.getDate()).padStart(2, "0");
        return `${date.getFullYear()}-${m}-${d}`;
    }

    function todayIso() {
        const date = new Date();
        const m = String(date.getMonth() + 1).padStart(2, "0");
        const d = String(date.getDate()).padStart(2, "0");
        return `${date.getFullYear()}-${m}-${d}`;
    }

    function readInviteStart() {
        const raw = inviteInput.value.trim();
        if (!raw) return null;
        const n = Number(raw);
        return Number.isInteger(n) && n >= 1 ? n : null;
    }

    function readWarnStart() {
        const n = Number(warnInput.value);
        return Number.isInteger(n) && n >= 1 ? n : null;
    }

    function periodList(row) {
        return Array.isArray(row["งวดดิบ"]) ? row["งวดดิบ"] : [];
    }

    function matchesPeriod(row) {
        const month = document.getElementById("letter-month").value;
        const year = document.getElementById("letter-year").value.trim();
        if (!month && !year) return true;
        return periodList(row).some((p) => {
            const [mm, yy] = String(p).trim().split("/");
            if (month && Number(mm) !== Number(month)) return false;
            if (year && String(yy ?? "").trim() !== year) return false;
            return true;
        });
    }

    function digitAt(account, pos) {
        const s = digits(account);
        const i = Number(pos) - 1;
        return i >= 0 && i < s.length ? s[i] : "";
    }

    function filteredRows() {
        const q = document.getElementById("letter-q").value.trim().toLowerCase();
        const status = document.getElementById("letter-status").value;
        const minP = document.getElementById("letter-min-periods").value;
        const maxP = document.getElementById("letter-max-periods").value;
        const digitPos = document.getElementById("letter-digit-pos").value;
        const digitVal = document.getElementById("letter-digit-val").value;
        const digit3 = document.getElementById("letter-digit3-9").checked;
        return records.filter((row) => {
            const acc = String(row["เลขที่บัญชี"] || "");
            const branch = String(row["สาขา"] || "");
            const name = displayName(row).toLowerCase();
            if (q && !(acc.toLowerCase().includes(q) || branch.toLowerCase().includes(q) || name.includes(q) || String(row["เลขอ้างอิง"] || "").toLowerCase().includes(q))) {
                return false;
            }
            if (status && (row["สถานะ สปก."] || "") !== status) return false;
            const count = Number(row["จำนวนงวดค้าง"] || 0);
            if (minP !== "" && count < Number(minP)) return false;
            if (maxP !== "" && count > Number(maxP)) return false;
            if (!matchesPeriod(row)) return false;
            if (digit3 && digitAt(acc, 3) !== "9") return false;
            if (digitPos && digitVal !== "" && digitAt(acc, digitPos) !== digitVal) return false;
            return true;
        });
    }

    function scopeValue() {
        return document.querySelector("input[name='letter-scope']:checked")?.value || "all";
    }

    function sortRows(rows) {
        return [...rows].sort((a, b) => rowKey(a).localeCompare(rowKey(b), "en"));
    }

    function printSet(visible) {
        const scope = scopeValue();
        if (scope === "selected") {
            return sortRows(records.filter((row) => selected.has(rowKey(row))));
        }
        if (scope === "ge2") {
            return sortRows(visible.filter((row) => Number(row["จำนวนงวดค้าง"] || 0) >= 2));
        }
        return sortRows(visible);
    }

    function fillStatusOptions() {
        const sel = document.getElementById("letter-status");
        const prev = sel.value;
        const statuses = [...new Set(records.map((r) => r["สถานะ สปก."] || "").filter(Boolean))].sort();
        sel.innerHTML = `<option value="">ทั้งหมด</option>` + statuses.map((s) => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join("");
        if ([...sel.options].some((o) => o.value === prev)) sel.value = prev;
    }

    function render() {
        const visible = sortRows(filteredRows());
        const printing = printSet(visible);
        const warnStart = readWarnStart();
        const inviteInvalid = inviteInput.value.trim() && readInviteStart() == null;
        const inviteStart = warnStart && !inviteInvalid ? readInviteStart() : null;
        const numbers = warnStart && !inviteInvalid ? assignNumbers(printing.length, warnStart, inviteStart) : [];
        const numberByKey = new Map();
        printing.forEach((row, i) => numberByKey.set(rowKey(row), numbers[i]));

        const first = numbers[0];
        const last = numbers[numbers.length - 1];
        const range = !warnStart
            ? "ยังไม่ระบุเลขเริ่มต้นหนังสือเตือน"
            : inviteInvalid
                ? "เลขเริ่มต้นหนังสือเชิญต้องเป็นจำนวนเต็มตั้งแต่ 1 หรือเว้นว่าง"
                : !printing.length
                    ? "ยังไม่มีรายการที่จะพิมพ์"
                    : `เตือน ${formatDocNo(first.warn)}–${formatDocNo(last.warn)} · เชิญ ${formatDocNo(first.invite)}–${formatDocNo(last.invite)}`;
        const hiddenSelected = printing.filter((row) => !visible.some((item) => rowKey(item) === rowKey(row))).length;
        const hiddenNote = hiddenSelected ? ` · ในจำนวนนี้ซ่อนจากตัวกรอง ${hiddenSelected} ราย แต่จะพิมพ์ด้วย` : "";
        summary.textContent = `แสดง ${visible.length} / ${records.length} ราย · จะพิมพ์ ${printing.length} ราย · ${range}${hiddenNote}`;

        const body = visible.map((row) => {
            const key = rowKey(row);
            const nums = numberByKey.get(key);
            const checked = selected.has(key) ? "checked" : "";
            return `<tr>
                <td class="text-center"><input type="checkbox" data-key="${escapeHtml(key)}" ${checked}></td>
                <td class="cell-mono">${escapeHtml(row["เลขที่บัญชี"] || "")}<br>${escapeHtml(row["สาขา"] || "")}</td>
                <td>${escapeHtml(displayName(row))}</td>
                <td class="text-center">${escapeHtml(row["สถานะ สปก."] || "-")}</td>
                <td class="text-center cell-mono">${escapeHtml(row["จำนวนงวดค้าง"] ?? "")}</td>
                <td>${escapeHtml(formatPeriodsLetter(periodList(row)))}</td>
                <td class="text-center cell-mono">${nums ? formatDocNo(nums.warn) : "—"}</td>
                <td class="text-center cell-mono">${nums ? formatDocNo(nums.invite) : "—"}</td>
            </tr>`;
        }).join("");

        wrap.innerHTML = `<table class="data-table"><thead><tr>
            <th><input type="checkbox" id="letter-check-all"></th>
            <th>บัญชี / สาขา</th><th>ชื่อสถานประกอบการ</th><th>สถานะ</th>
            <th>งวด</th><th>งวดเดือน</th><th>เลขเตือน</th><th>เลขเชิญ</th>
        </tr></thead><tbody>${body || `<tr><td colspan="8" class="text-center">ไม่พบรายการตามตัวกรอง</td></tr>`}</tbody></table>`;

        wrap.querySelectorAll("input[data-key]").forEach((box) => {
            box.addEventListener("change", () => {
                if (box.checked) selected.add(box.dataset.key);
                else selected.delete(box.dataset.key);
                render();
            });
        });
        const allBox = document.getElementById("letter-check-all");
        if (allBox) {
            allBox.checked = visible.length > 0 && visible.every((row) => selected.has(rowKey(row)));
            allBox.addEventListener("change", () => {
                visible.forEach((row) => {
                    if (allBox.checked) selected.add(rowKey(row));
                    else selected.delete(rowKey(row));
                });
                render();
            });
        }
    }

    function toCase(row) {
        let street = row["ที่อยู่"] || "";
        const province = String(row["จังหวัด"] || "").trim();
        const isBkk = province.includes("กรุงเทพ") || province === "กทม" || province === "กทม.";
        const tambon = String(row["ตำบล"] || "").trim().replace(/^(ตำบล|ต\.|แขวง)\s*/, "");
        if (tambon && !/(?:ตำบล|ต\.|แขวง)/.test(street)) {
            street = `${street} ${isBkk ? `แขวง${tambon}` : `ต.${tambon}`}`.trim();
        }
        return {
            name: displayName(row),
            account: row["เลขที่บัญชี"] || "",
            branch: row["สาขา"] || "",
            periods: periodList(row),
            street,
            district: row["อำเภอ"] || "",
            province,
            zipcode: String(row["รหัสไปรษณีย์"] || ""),
        };
    }

    async function download(fmt) {
        const warnStart = readWarnStart();
        if (!warnStart) {
            statusMsg.textContent = "กรุณากรอกเลขเริ่มต้นหนังสือเตือน";
            return;
        }
        if (inviteInput.value.trim() && readInviteStart() == null) {
            statusMsg.textContent = "เลขเริ่มต้นหนังสือเชิญต้องเป็นจำนวนเต็มตั้งแต่ 1 หรือเว้นว่าง";
            return;
        }
        if (!issueInput.value || !apptInput.value) {
            statusMsg.textContent = "กรุณาระบุวันที่ออกหนังสือและวันที่นัดพบ";
            return;
        }
        const printing = printSet(filteredRows());
        if (!printing.length) {
            statusMsg.textContent = "ไม่มีรายการที่จะพิมพ์ตามเงื่อนไขนี้";
            return;
        }
        statusMsg.textContent = `กำลังสร้างไฟล์ ${printing.length} ราย...`;
        document.getElementById("letter-docx-btn").disabled = true;
        document.getElementById("letter-pdf-btn").disabled = true;
        try {
            const res = await fetch("/api/letters/generate", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    format: fmt,
                    issue_date: issueInput.value,
                    appoint_date: apptInput.value,
                    warn_start: warnStart,
                    invite_start: readInviteStart(),
                    cases: printing.map(toCase),
                }),
            });
            if (!res.ok) {
                let detail = "สร้างหนังสือไม่สำเร็จ";
                try { detail = (await res.json()).detail || detail; } catch (e) { /* keep */ }
                throw new Error(detail);
            }
            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            const stamp = issueInput.value.replace(/-/g, "");
            a.href = url;
            a.download = `หนังสือเตือนเชิญ_${stamp}.${fmt}`;
            document.body.appendChild(a);
            a.click();
            a.remove();
            URL.revokeObjectURL(url);
            statusMsg.textContent = `ดาวน์โหลดแล้ว ${printing.length} ราย`;
        } catch (err) {
            statusMsg.textContent = err.message;
        } finally {
            document.getElementById("letter-docx-btn").disabled = false;
            document.getElementById("letter-pdf-btn").disabled = false;
        }
    }

    function refreshDates() {
        document.getElementById("letter-issue-preview").textContent = isoToBe(issueInput.value);
        document.getElementById("letter-appt-preview").textContent = isoToBe(apptInput.value);
    }

    issueInput.value = todayIso();
    apptInput.value = addDays(issueInput.value, 30);
    refreshDates();
    issueInput.addEventListener("change", () => {
        if (!apptTouched) apptInput.value = addDays(issueInput.value, 30);
        refreshDates();
    });
    apptInput.addEventListener("change", () => {
        apptTouched = true;
        refreshDates();
    });

    ["letter-q", "letter-status", "letter-min-periods", "letter-max-periods", "letter-month", "letter-year", "letter-digit-pos", "letter-digit-val", "letter-warn-start", "letter-invite-start"].forEach((id) => {
        document.getElementById(id).addEventListener("input", render);
        document.getElementById(id).addEventListener("change", render);
    });
    document.getElementById("letter-digit3-9").addEventListener("change", render);
    document.getElementById("letter-scope").addEventListener("change", render);
    document.getElementById("letter-docx-btn").addEventListener("click", () => download("docx"));
    document.getElementById("letter-pdf-btn").addEventListener("click", () => download("pdf"));
    document.getElementById("goto-letter-btn").addEventListener("click", () => {
        if (!records.length) {
            statusMsg.textContent = "";
            alert("ประมวลผลรายงานหนี้ก่อน แล้วค่อยพิมพ์หนังสือ");
            return;
        }
        showDesk("letter");
    });
    document.getElementById("desk-ledger-btn").addEventListener("click", () => showDesk("ledger"));

    function showDesk(name) {
        const letter = name === "letter";
        document.body.classList.toggle("desk-letter", letter);
        document.body.classList.toggle("desk-ledger", !letter);
        document.getElementById("desk-ledger-btn").classList.toggle("is-active", !letter);
        document.getElementById("goto-letter-btn").classList.toggle("is-active", letter);
        document.getElementById("desk-ledger-btn").setAttribute("aria-pressed", letter ? "false" : "true");
        document.getElementById("goto-letter-btn").setAttribute("aria-pressed", letter ? "true" : "false");
        document.querySelector(".upload-section").classList.toggle("hidden", letter);
        document.getElementById("results-section").classList.toggle("hidden", letter || !records.length);
        section.classList.toggle("hidden", !letter);
        window.scrollTo({ top: 0, behavior: "smooth" });
    }

    document.addEventListener("debt:ready", (event) => {
        records = event.detail || [];
        selected = new Set();
        fillStatusOptions();
        section.classList.add("hidden");
        render();
    });
});
