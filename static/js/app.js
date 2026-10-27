/**
 * SSO Debt Report System - Frontend Application Logic
 */

document.addEventListener("DOMContentLoaded", () => {
    // DOM Elements
    const wcfStatusBadge = document.getElementById("wcf-status-badge");
    const wcfStatusText = document.getElementById("wcf-status-text");
    const dropZone = document.getElementById("drop-zone");
    const fileInput = document.getElementById("file-input");
    const browseBtn = document.getElementById("browse-btn");
    const selectedFileName = document.getElementById("selected-file-name");
    const inputFilesList = document.getElementById("input-files-list");
    const processBtn = document.getElementById("process-btn");
    const progressContainer = document.getElementById("progress-container");
    const progressStatus = document.getElementById("progress-status");

    // Results Elements
    const resultsSection = document.getElementById("results-section");
    const kpiPrincipal = document.getElementById("kpi-principal");
    const kpiEstablishments = document.getElementById("kpi-establishments");
    const kpiRawCount = document.getElementById("kpi-raw-count");
    const kpiGrandTotal = document.getElementById("kpi-grand-total");
    const kpiSurchargeInfo = document.getElementById("kpi-surcharge-info");
    const kpiMatchRate = document.getElementById("kpi-match-rate");
    const kpiMatchCounts = document.getElementById("kpi-match-counts");
    const downloadBtn = document.getElementById("download-btn");
    const tableBody = document.getElementById("table-body");
    const tableCountLabel = document.getElementById("table-count-label");
    const searchInput = document.getElementById("search-input");

    // Master Data Modal Elements
    const openMasterModalBtn = document.getElementById("open-master-modal-btn");
    const masterModal = document.getElementById("master-modal");
    const closeMasterModalBtn = document.getElementById("close-master-modal-btn");
    const cancelMasterBtn = document.getElementById("cancel-master-btn");
    const masterFilePath = document.getElementById("master-file-path");
    const masterRecordsCount = document.getElementById("master-records-count");
    const masterLastModified = document.getElementById("master-last-modified");
    const masterFileSize = document.getElementById("master-file-size");
    const masterFileInput = document.getElementById("master-file-input");
    const browseMasterBtn = document.getElementById("browse-master-btn");
    const selectedMasterName = document.getElementById("selected-master-name");
    const saveMasterBtn = document.getElementById("save-master-btn");
    const reloadDiskBtn = document.getElementById("reload-disk-btn");
    const masterStatusMsg = document.getElementById("master-status-msg");

    let currentSelectedFile = null;
    let currentSelectedExisting = null;
    let currentMasterFile = null;
    let allTableRecords = [];
    let hasSsoChecked = false;
    let isSsoChecking = false;
    let cancelSsoCheckFlag = false;
    let currentEditingRow = null;
    let currentEditingPeriods = [];

    // SSO CORE & Period Editor Elements
    const ssoStartYearInput = document.getElementById("sso-start-year");
    const ssoCheckBtn = document.getElementById("sso-check-btn");
    const ssoCancelBtn = document.getElementById("sso-cancel-btn");
    const periodSourceSwitch = document.getElementById("period-source-switch");
    const periodSourceExcelBtn = document.getElementById("period-source-excel-btn");
    const periodSourceSsoBtn = document.getElementById("period-source-sso-btn");
    const ssoProgressWrap = document.getElementById("sso-progress-wrap");
    const ssoProgressBar = document.getElementById("sso-progress-bar");
    const ssoProgressStatus = document.getElementById("sso-progress-status");
    const ssoSummaryBanner = document.getElementById("sso-summary-banner");

    const editPeriodsModal = document.getElementById("edit-periods-modal");
    const closeEditPeriodsBtn = document.getElementById("close-edit-periods-btn");
    const cancelEditPeriodsBtn = document.getElementById("cancel-edit-periods-btn");
    const saveEditPeriodsBtn = document.getElementById("save-edit-periods-btn");
    const editPeriodAcc = document.getElementById("edit-period-acc");
    const editPeriodName = document.getElementById("edit-period-name");
    const editPeriodExcelRef = document.getElementById("edit-period-excel-ref");
    const editPeriodSsoRef = document.getElementById("edit-period-sso-ref");
    const editPeriodCount = document.getElementById("edit-period-count");
    const editPeriodChips = document.getElementById("edit-period-chips");
    const restoreExcelPeriodsBtn = document.getElementById("restore-excel-periods-btn");
    const restoreSsoPeriodsBtn = document.getElementById("restore-sso-periods-btn");
    const clearAllPeriodsBtn = document.getElementById("clear-all-periods-btn");
    const addPeriodMonth = document.getElementById("add-period-month");
    const addPeriodYear = document.getElementById("add-period-year");
    const addSinglePeriodBtn = document.getElementById("add-single-period-btn");
    const addPeriodText = document.getElementById("add-period-text");
    const addTextPeriodBtn = document.getElementById("add-text-period-btn");
    const editPeriodPreviewB = document.getElementById("edit-period-preview-b");
    const editPeriodPreviewLetter = document.getElementById("edit-period-preview-letter");

    const THAI_MONTHS = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];

    function normalizeAndSortPeriods(list) {
        const parsed = [];
        (list || []).forEach((raw) => {
            const parts = String(raw || "").trim().split("/");
            if (parts.length !== 2) return;
            const m = Number(parts[0]);
            const y = Number(parts[1]);
            if (Number.isInteger(m) && m >= 1 && m <= 12 && Number.isInteger(y) && y >= 2400 && y <= 2700) {
                parsed.push([y, m]);
            }
        });
        parsed.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
        const unique = [];
        parsed.forEach((item) => {
            const prev = unique[unique.length - 1];
            if (!prev || prev[0] !== item[0] || prev[1] !== item[1]) {
                unique.push(`${String(item[1]).padStart(2, "0")}/${item[0]}`);
            }
        });
        return unique;
    }

    function formatPeriodRangesB(periodList) {
        const clean = normalizeAndSortPeriods(periodList);
        if (!clean.length) return "ชำระครบแล้ว (0 งวด)";
        const tuples = clean.map((p) => {
            const [m, y] = p.split("/").map(Number);
            return [y, m];
        });
        const groups = [[tuples[0]]];
        for (let i = 1; i < tuples.length; i++) {
            const prev = groups[groups.length - 1].at(-1);
            const cur = tuples[i];
            if ((cur[0] - prev[0]) * 12 + (cur[1] - prev[1]) === 1) {
                groups[groups.length - 1].push(cur);
            } else {
                groups.push([cur]);
            }
        }
        return groups.map((g) => {
            if (g.length === 1) return `${THAI_MONTHS[g[0][1]]} ${g[0][0]}`;
            const start = g[0];
            const end = g[g.length - 1];
            return `${THAI_MONTHS[start[1]]} ${start[0]} - ${THAI_MONTHS[end[1]]} ${end[0]}`;
        }).join(", ");
    }

    function formatPeriodRangesLetter(periodList) {
        const clean = normalizeAndSortPeriods(periodList);
        if (!clean.length) return "— (ไม่มีงวดค้างชำระ)";
        const tuples = clean.map((p) => {
            const [m, y] = p.split("/").map(Number);
            return [y, m];
        });
        const groups = [[tuples[0]]];
        for (let i = 1; i < tuples.length; i++) {
            const prev = groups[groups.length - 1].at(-1);
            const cur = tuples[i];
            if ((cur[0] - prev[0]) * 12 + (cur[1] - prev[1]) === 1) {
                groups[groups.length - 1].push(cur);
            } else {
                groups.push([cur]);
            }
        }
        return groups.map((g) => {
            const [sY, sM] = g[0];
            const [eY, eM] = g[g.length - 1];
            if (g.length === 1) return `${THAI_MONTHS[sM]} ${sY}`;
            if (sY === eY) return `${THAI_MONTHS[sM]}-${THAI_MONTHS[eM]} ${eY}`;
            return `${THAI_MONTHS[sM]} ${sY}-${THAI_MONTHS[eM]} ${eY}`;
        }).join(", ");
    }

    // Helper for XSS protection
    function escapeHtml(str) {
        if (str === null || str === undefined) return "";
        return String(str)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    // 1. Initial System Check & File List
    async function loadSystemInfo() {
        try {
            const res = await fetch("/api/system-info");
            const data = await res.json();

            // WCF Master Data Status
            if (data.wcf_connected) {
                wcfStatusBadge.className = "badge badge-connected";
                wcfStatusText.textContent = `WCF Master Data: พร้อมใช้งาน (${data.wcf_records_count.toLocaleString()} สปก.)`;
            } else {
                wcfStatusBadge.className = "badge badge-error";
                wcfStatusText.textContent = "WCF Master Data: ไม่พบไฟล์ฐานข้อมูล";
            }

            // Update Modal Stats
            if (masterFilePath) masterFilePath.textContent = data.wcf_source_path || "data/DATA-EMP-WCF.xlsx";
            if (masterRecordsCount) masterRecordsCount.textContent = `${data.wcf_records_count.toLocaleString()} แห่ง`;
            if (masterLastModified) masterLastModified.textContent = data.wcf_last_modified || "-";
            if (masterFileSize) masterFileSize.textContent = `${data.wcf_size_kb || 0} KB`;

            // Input Files List
            renderInputFiles(data.input_files);
        } catch (err) {
            console.error("System info error:", err);
            wcfStatusBadge.className = "badge badge-error";
            wcfStatusText.textContent = "ไม่สามารถเชื่อมต่อ Server ได้";
        }
    }

    function renderInputFiles(files) {
        if (!files || files.length === 0) {
            inputFilesList.innerHTML = '<div class="loading-placeholder">ไม่พบไฟล์ในโฟลเดอร์ Input</div>';
            return;
        }

        inputFilesList.innerHTML = "";
        files.forEach((f, idx) => {
            const item = document.createElement("div");
            item.className = "file-item" + (idx === 0 ? " selected" : "");
            item.innerHTML = `
                <span class="file-item-name" title="${escapeHtml(f.filename)}">
                    📄 ${escapeHtml(f.filename)}
                </span>
                <span class="file-item-size">${escapeHtml(f.size_kb)} KB</span>
            `;

            item.addEventListener("click", () => {
                document.querySelectorAll(".file-item").forEach(el => el.classList.remove("selected"));
                item.classList.add("selected");
                currentSelectedExisting = f.filename;
                currentSelectedFile = null;
                selectedFileName.classList.add("hidden");
                processBtn.disabled = false;
            });

            inputFilesList.appendChild(item);
        });

        // Auto select first file
        if (files.length > 0) {
            currentSelectedExisting = files[0].filename;
            processBtn.disabled = false;
        }
    }

    // 2. Drag & Drop Event Listeners
    browseBtn.addEventListener("click", () => fileInput.click());
    dropZone.addEventListener("click", (e) => {
        if (e.target !== browseBtn) fileInput.click();
    });

    fileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handleFileSelect(e.target.files[0]);
        }
    });

    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("drag-over");
    });

    dropZone.addEventListener("dragleave", () => {
        dropZone.classList.remove("drag-over");
    });

    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropZone.classList.remove("drag-over");
        if (e.dataTransfer.files.length > 0) {
            handleFileSelect(e.dataTransfer.files[0]);
        }
    });

    function handleFileSelect(file) {
        currentSelectedFile = file;
        currentSelectedExisting = null;
        document.querySelectorAll(".file-item").forEach(el => el.classList.remove("selected"));

        selectedFileName.textContent = `เลือกไฟล์: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        selectedFileName.classList.remove("hidden");
        processBtn.disabled = false;
    }

    // 3. Process Button Click Handler
    processBtn.addEventListener("click", async () => {
        if (!currentSelectedFile && !currentSelectedExisting) return;

        processBtn.disabled = true;
        progressContainer.classList.remove("hidden");
        progressStatus.textContent = "กำลังประมวลผลข้อมูลและเชื่อมโยงกับฐานข้อมูล WCF...";

        try {
            let res;
            if (currentSelectedFile) {
                const formData = new FormData();
                formData.append("file", currentSelectedFile);
                res = await fetch("/api/upload-and-process", {
                    method: "POST",
                    body: formData
                });
            } else {
                res = await fetch("/api/process-existing", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ filename: currentSelectedExisting })
                });
            }

            const data = await res.json();
            if (!res.ok || !data.success) {
                throw new Error(data.detail || data.error || "ประมวลผลล้มเหลว");
            }

            renderResults(data);
        } catch (err) {
            alert("เกิดข้อผิดพลาด: " + err.message);
        } finally {
            progressContainer.classList.add("hidden");
            processBtn.disabled = false;
        }
    });

    // 4. Render Results & Preview Table
    function renderResults(data) {
        // KPIs
        kpiPrincipal.textContent = formatCurrency(data.total_principal_debt);
        kpiEstablishments.textContent = data.total_establishments.toLocaleString();
        kpiRawCount.textContent = `จากรายงานต้นฉบับ ${data.total_raw_records.toLocaleString()} แถว`;
        kpiGrandTotal.textContent = formatCurrency(data.grand_total_debt);
        kpiSurchargeInfo.textContent = `เงินเพิ่มรวม ${formatCurrency(data.total_surcharge)} บาท`;
        kpiMatchRate.textContent = `${data.match_rate_pct}%`;
        kpiMatchCounts.textContent = `จับคู่สำเร็จ ${data.matched_wcf_count} / ${data.total_establishments} แห่ง`;

        // Download Link
        downloadBtn.href = `/api/download/${encodeURIComponent(data.output_filename)}`;

        // Table Records
        allTableRecords = (data.preview_data || []).map((r) => {
            const orig = Array.isArray(r["งวดดิบ_ต้นฉบับ"])
                ? normalizeAndSortPeriods(r["งวดดิบ_ต้นฉบับ"])
                : normalizeAndSortPeriods(r["งวดดิบ"] || []);
            return {
                ...r,
                "งวดดิบ": normalizeAndSortPeriods(r["งวดดิบ"] || orig),
                "งวดดิบ_ต้นฉบับ": orig,
                "แหล่งข้อมูลงวด": r["แหล่งข้อมูลงวด"] || "Excel",
            };
        });
        hasSsoChecked = false;
        if (periodSourceSwitch) periodSourceSwitch.classList.add("hidden");
        if (ssoSummaryBanner) {
            ssoSummaryBanner.classList.add("hidden");
            ssoSummaryBanner.textContent = "";
        }
        if (ssoProgressWrap) ssoProgressWrap.classList.add("hidden");

        tableCountLabel.textContent = `${allTableRecords.length} รายการ`;
        renderTable(getFilteredTableRecords());

        document.dispatchEvent(new CustomEvent("debt:ready", { detail: allTableRecords }));

        // Show section & scroll
        resultsSection.classList.remove("hidden");
        resultsSection.scrollIntoView({ behavior: "smooth" });
    }

    function getFilteredTableRecords() {
        const query = (searchInput?.value || "").trim().toLowerCase();
        if (!query) return allTableRecords;
        return allTableRecords.filter(r => {
            return (
                (r["เลขอ้างอิง"] && r["เลขอ้างอิง"].toLowerCase().includes(query)) ||
                (r["เลขที่บัญชี"] && r["เลขที่บัญชี"].toLowerCase().includes(query)) ||
                (r["สาขา"] && r["สาขา"].toLowerCase().includes(query)) ||
                (r["ชื่อสถานประกอบการ (จริง)"] && r["ชื่อสถานประกอบการ (จริง)"].toLowerCase().includes(query)) ||
                (r["ชื่อในรายงานหนี้"] && r["ชื่อในรายงานหนี้"].toLowerCase().includes(query)) ||
                (r["ตำบล"] && r["ตำบล"].toLowerCase().includes(query)) ||
                (r["อำเภอ"] && r["อำเภอ"].toLowerCase().includes(query)) ||
                (r["จังหวัด"] && r["จังหวัด"].toLowerCase().includes(query))
            );
        });
    }

    function getSourceBadgeHtml(source) {
        if (source === "SSO CORE") {
            return `<span class="source-tag source-tag-sso">SSO CORE</span>`;
        }
        if (source === "แก้ไขเอง") {
            return `<span class="source-tag source-tag-custom">แก้ไขเอง</span>`;
        }
        return `<span class="source-tag source-tag-excel">Excel</span>`;
    }

    function renderTable(records) {
        tableBody.innerHTML = "";
        if (records.length === 0) {
            tableBody.innerHTML = `<tr><td colspan="14" style="text-align: center; padding: 24px; color: var(--text-muted);">ไม่พบข้อมูลที่ตรงกับเงื่อนไขค้นหา</td></tr>`;
            return;
        }

        records.forEach((row, idx) => {
            const tr = document.createElement("tr");
            const statusVal = row["สถานะ สปก."] || "default";
            const statusClass = `status-${escapeHtml(statusVal)}`;
            const reportName = row["ชื่อในรายงานหนี้"] || "";
            const nameDisplay = row["ชื่อสถานประกอบการ (จริง)"]
                ? escapeHtml(row["ชื่อสถานประกอบการ (จริง)"])
                : `${escapeHtml(reportName)} <span style="color: var(--text-muted); font-style: italic;">[ไม่พบใน WCF - ตรวจสอบด้วยตนเอง]</span>`;

            const pCount = Number(row["จำนวนงวดค้าง"] ?? 0);
            const periodStr = row["งวดเดือนค้างชำระ"] || "";
            const periodCellHtml = pCount === 0
                ? `<span class="paid-zero-badge">${escapeHtml(periodStr || "ชำระครบแล้ว (0 งวด)")}</span> ${getSourceBadgeHtml(row["แหล่งข้อมูลงวด"])}`
                : `${escapeHtml(periodStr)} ${getSourceBadgeHtml(row["แหล่งข้อมูลงวด"])}`;
            const refKey = row["เลขอ้างอิง"] || `${row["เลขที่บัญชี"]};${row["สาขา"]}`;

            tr.innerHTML = `
                <td class="text-center cell-mono">${idx + 1}</td>
                <td class="cell-mono"><strong>${escapeHtml(row["เลขอ้างอิง"] || "")}</strong></td>
                <td><strong>${nameDisplay}</strong></td>
                <td>${escapeHtml(row["ตำบล"] || "-")}</td>
                <td>${escapeHtml(row["อำเภอ"] || "-")}</td>
                <td>${escapeHtml(row["จังหวัด"] || "-")}</td>
                <td class="text-center cell-mono">${escapeHtml(row["รหัสไปรษณีย์"] || "-")}</td>
                <td class="text-center cell-mono">${escapeHtml(pCount)}</td>
                <td>${periodCellHtml}</td>
                <td class="text-center">
                    <button type="button" class="btn-edit-period" data-edit-ref="${escapeHtml(refKey)}" title="แก้ไขงวดค้างชำระ">✏️ แก้ไขงวด</button>
                </td>
                <td class="text-right cell-debt">${formatCurrency(row["ยอดหนี้ไม่รวมเงินเพิ่ม"])}</td>
                <td class="text-right cell-mono">${formatCurrency(row["เงินเพิ่ม"])}</td>
                <td class="text-right cell-total">${formatCurrency(row["ยอดค้างชำระรวมทั้งสิ้น"])}</td>
                <td class="text-center"><span class="status-pill ${statusClass}">${escapeHtml(row["สถานะ สปก."] || "-")}</span></td>
            `;
            tableBody.appendChild(tr);
        });

        tableBody.querySelectorAll("button[data-edit-ref]").forEach((btn) => {
            btn.addEventListener("click", () => {
                openEditPeriodsModal(btn.getAttribute("data-edit-ref"));
            });
        });
    }

    // 5. Live Search Filter
    searchInput.addEventListener("input", () => {
        const filtered = getFilteredTableRecords();
        renderTable(filtered);
        const query = (searchInput.value || "").trim();
        tableCountLabel.textContent = query
            ? `${filtered.length} / ${allTableRecords.length} รายการ`
            : `${allTableRecords.length} รายการ`;
    });

    function formatCurrency(val) {
        if (val === null || val === undefined) return "0.00";
        return Number(val).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    // 6. Master Data Modal Handlers
    function openMasterModal() {
        loadSystemInfo();
        masterStatusMsg.classList.add("hidden");
        masterStatusMsg.textContent = "";
        selectedMasterName.classList.add("hidden");
        selectedMasterName.textContent = "";
        saveMasterBtn.disabled = true;
        currentMasterFile = null;
        masterModal.classList.remove("hidden");
    }

    function closeMasterModal() {
        masterModal.classList.add("hidden");
    }

    if (openMasterModalBtn) openMasterModalBtn.addEventListener("click", openMasterModal);
    if (wcfStatusBadge) wcfStatusBadge.addEventListener("click", openMasterModal);
    if (closeMasterModalBtn) closeMasterModalBtn.addEventListener("click", closeMasterModal);
    if (cancelMasterBtn) cancelMasterBtn.addEventListener("click", closeMasterModal);

    masterModal.addEventListener("click", (e) => {
        if (e.target === masterModal) closeMasterModal();
    });

    if (browseMasterBtn) browseMasterBtn.addEventListener("click", () => masterFileInput.click());

    if (masterFileInput) {
        masterFileInput.addEventListener("change", (e) => {
            if (e.target.files.length > 0) {
                currentMasterFile = e.target.files[0];
                selectedMasterName.textContent = `ไฟล์ที่เลือก: ${currentMasterFile.name} (${(currentMasterFile.size / 1024).toFixed(1)} KB)`;
                selectedMasterName.classList.remove("hidden");
                saveMasterBtn.disabled = false;
            }
        });
    }

    if (saveMasterBtn) {
        saveMasterBtn.addEventListener("click", async () => {
            if (!currentMasterFile) return;

            saveMasterBtn.disabled = true;
            masterStatusMsg.className = "master-status-msg";
            masterStatusMsg.textContent = "กำลังอัปโหลดและประมวลผล Master Data...";
            masterStatusMsg.classList.remove("hidden");

            try {
                const formData = new FormData();
                formData.append("file", currentMasterFile);

                const res = await fetch("/api/upload-master-data", {
                    method: "POST",
                    body: formData
                });
                const data = await res.json();
                if (!res.ok || !data.success) {
                    throw new Error(data.detail || data.message || "อัปโหลด Master Data ล้มเหลว");
                }

                masterStatusMsg.className = "master-status-msg success";
                masterStatusMsg.textContent = data.message;
                await loadSystemInfo();
                setTimeout(closeMasterModal, 2000);
            } catch (err) {
                masterStatusMsg.className = "master-status-msg error";
                masterStatusMsg.textContent = "เกิดข้อผิดพลาด: " + err.message;
                saveMasterBtn.disabled = false;
            }
        });
    }

    if (reloadDiskBtn) {
        reloadDiskBtn.addEventListener("click", async () => {
            reloadDiskBtn.disabled = true;
            masterStatusMsg.className = "master-status-msg";
            masterStatusMsg.textContent = "กำลังตรวจสอบและโหลดข้อมูลจากดิสก์...";
            masterStatusMsg.classList.remove("hidden");

            try {
                const res = await fetch("/api/reload-master-data", { method: "POST" });
                const data = await res.json();
                if (!res.ok || !data.success) {
                    throw new Error(data.detail || data.message || "รีโหลดข้อมูลล้มเหลว");
                }

                masterStatusMsg.className = "master-status-msg success";
                masterStatusMsg.textContent = data.message;
                await loadSystemInfo();
            } catch (err) {
                masterStatusMsg.className = "master-status-msg error";
                masterStatusMsg.textContent = "เกิดข้อผิดพลาด: " + err.message;
            } finally {
                reloadDiskBtn.disabled = false;
            }
        });
    }

    // 7. SSO CORE Overdue Check & Source Switching
    async function syncRecordsAndReExport(sourceLabel) {
        renderTable(getFilteredTableRecords());
        document.dispatchEvent(new CustomEvent("debt:updated", { detail: allTableRecords }));

        try {
            const baseName = currentSelectedFile
                ? currentSelectedFile.name.replace(/\.[^.]+$/, "")
                : (currentSelectedExisting ? currentSelectedExisting.replace(/\.[^.]+$/, "") : "Updated");
            const res = await fetch("/api/records/re-export", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    records: allTableRecords,
                    source_filename: `${baseName}_${sourceLabel || "Updated"}`
                })
            });
            if (res.ok) {
                const data = await res.json();
                if (data.output_filename) {
                    downloadBtn.href = `/api/download/${encodeURIComponent(data.output_filename)}`;
                }
            }
        } catch (err) {
            console.warn("Re-export warning:", err);
        }
    }

    function applyPeriodSourceMode(mode) {
        allTableRecords.forEach((row) => {
            if (mode === "SSO CORE" && Array.isArray(row["งวดดิบ_sso"])) {
                const ssoP = normalizeAndSortPeriods(row["งวดดิบ_sso"]);
                row["งวดดิบ"] = ssoP;
                row["จำนวนงวดค้าง"] = ssoP.length;
                row["งวดเดือนค้างชำระ"] = formatPeriodRangesB(ssoP);
                row["แหล่งข้อมูลงวด"] = "SSO CORE";
            } else if (mode === "Excel") {
                const origP = normalizeAndSortPeriods(row["งวดดิบ_ต้นฉบับ"] || []);
                row["งวดดิบ"] = origP;
                row["จำนวนงวดค้าง"] = origP.length;
                row["งวดเดือนค้างชำระ"] = formatPeriodRangesB(origP);
                row["แหล่งข้อมูลงวด"] = "Excel";
            }
        });

        if (periodSourceExcelBtn && periodSourceSsoBtn) {
            periodSourceExcelBtn.classList.toggle("is-active", mode === "Excel");
            periodSourceSsoBtn.classList.toggle("is-active", mode === "SSO CORE");
        }
        syncRecordsAndReExport(mode === "SSO CORE" ? "SSOCORE" : "Excel");
    }

    if (periodSourceExcelBtn) {
        periodSourceExcelBtn.addEventListener("click", () => applyPeriodSourceMode("Excel"));
    }
    if (periodSourceSsoBtn) {
        periodSourceSsoBtn.addEventListener("click", () => applyPeriodSourceMode("SSO CORE"));
    }

    if (ssoCancelBtn) {
        ssoCancelBtn.addEventListener("click", () => {
            cancelSsoCheckFlag = true;
            ssoProgressStatus.textContent = "กำลังหยุดการตรวจสอบ...";
        });
    }

    if (ssoCheckBtn) {
        ssoCheckBtn.addEventListener("click", async () => {
            if (!allTableRecords.length || isSsoChecking) return;

            const startYear = Number(ssoStartYearInput?.value) || 2568;
            isSsoChecking = true;
            cancelSsoCheckFlag = false;
            ssoCheckBtn.disabled = true;
            if (ssoCancelBtn) ssoCancelBtn.classList.remove("hidden");
            if (ssoProgressWrap) ssoProgressWrap.classList.remove("hidden");
            if (ssoProgressBar) {
                ssoProgressBar.classList.remove("indeterminate");
                ssoProgressBar.style.width = "2%";
            }
            if (ssoSummaryBanner) ssoSummaryBanner.classList.add("hidden");

            const total = allTableRecords.length;
            const batchSize = 5;
            let checkedCount = 0;
            let overdueFirmCount = 0;
            let paidZeroCount = 0;

            try {
                for (let i = 0; i < total; i += batchSize) {
                    if (cancelSsoCheckFlag) break;
                    const slice = allTableRecords.slice(i, i + batchSize);
                    const pct = Math.max(5, Math.round((i / total) * 100));
                    if (ssoProgressBar) ssoProgressBar.style.width = `${pct}%`;
                    if (ssoProgressStatus) {
                        ssoProgressStatus.textContent = `กำลังตรวจสอบกับ SSO CORE (${i + 1}–${Math.min(i + slice.length, total)} จาก ${total} แห่ง)...`;
                    }

                    const payload = {
                        start_year: startYear,
                        accounts: slice.map((r) => ({
                            acc: String(r["เลขที่บัญชี"] || ""),
                            branch: String(r["สาขา"] || "000000")
                        }))
                    };

                    const res = await fetch("/api/sso-core/check", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify(payload)
                    });
                    const data = await res.json();
                    if (!res.ok || !data.success) {
                        throw new Error(data.detail || "ตรวจสอบ SSO CORE ล้มเหลว");
                    }

                    const byKey = new Map();
                    (data.results || []).forEach((item) => {
                        byKey.set(`${item.acc};${item.branch}`, item);
                    });

                    slice.forEach((row) => {
                        const key = `${row["เลขที่บัญชี"]};${row["สาขา"] || "000000"}`;
                        const hit = byKey.get(key);
                        if (hit) {
                            const ssoPeriods = normalizeAndSortPeriods(hit.overdue_periods || []);
                            row["งวดดิบ_sso"] = ssoPeriods;
                            row["งวดดิบ"] = [...ssoPeriods];
                            row["จำนวนงวดค้าง"] = ssoPeriods.length;
                            row["งวดเดือนค้างชำระ"] = hit.period_text_b || formatPeriodRangesB(ssoPeriods);
                            row["แหล่งข้อมูลงวด"] = "SSO CORE";
                            row["สถานะตรวจสอบ_SSO"] = hit.sso_check_badge || "";
                            checkedCount++;
                            if (ssoPeriods.length > 0) overdueFirmCount++;
                            else paidZeroCount++;
                        }
                    });

                    renderTable(getFilteredTableRecords());
                }

                hasSsoChecked = checkedCount > 0;
                if (hasSsoChecked && periodSourceSwitch) {
                    periodSourceSwitch.classList.remove("hidden");
                    if (periodSourceExcelBtn) periodSourceExcelBtn.classList.remove("is-active");
                    if (periodSourceSsoBtn) periodSourceSsoBtn.classList.add("is-active");
                }

                if (ssoProgressBar) ssoProgressBar.style.width = "100%";
                if (ssoSummaryBanner) {
                    const statusPrefix = cancelSsoCheckFlag ? "หยุดตรวจสอบชั่วคราว:" : "✅ ตรวจสอบ SSO CORE สำเร็จ:";
                    ssoSummaryBanner.textContent = `${statusPrefix} ตรวจสอบแล้ว ${checkedCount}/${total} แห่ง · พบยังค้างชำระ ${overdueFirmCount} แห่ง · ชำระครบแล้ว (0 งวด) ${paidZeroCount} แห่ง (ขณะนี้ระบบอ้างอิงงวดตาม SSO CORE สำหรับพิมพ์หนังสือและส่งออก Excel)`;
                    ssoSummaryBanner.classList.remove("hidden");
                }

                await syncRecordsAndReExport("SSOCORE");
            } catch (err) {
                const msg = err.message || "ไม่สามารถเชื่อมต่อระบบ SSO CORE ได้";
                if (ssoSummaryBanner) {
                    ssoSummaryBanner.textContent = `⚠️ ${msg}`;
                    ssoSummaryBanner.classList.remove("hidden");
                }
                alert(msg.startsWith("ไม่สามารถ") ? msg : `เกิดข้อผิดพลาดในการตรวจสอบ SSO CORE: ${msg}`);
            } finally {
                isSsoChecking = false;
                ssoCheckBtn.disabled = false;
                if (ssoCancelBtn) ssoCancelBtn.classList.add("hidden");
                if (ssoProgressWrap) ssoProgressWrap.classList.add("hidden");
            }
        });
    }

    // 8. Edit Overdue Periods Modal Logic
    function parseFlexiblePeriodText(rawText) {
        const added = [];
        const tokens = String(rawText || "")
            .split(/[,;\s]+/)
            .map((t) => t.trim())
            .filter(Boolean);

        tokens.forEach((tok) => {
            const rangeMatch = tok.match(/^(\d{1,2})\/(\d{4})-(\d{1,2})\/(\d{4})$/);
            if (rangeMatch) {
                let m1 = Number(rangeMatch[1]);
                let y1 = Number(rangeMatch[2]);
                const m2 = Number(rangeMatch[3]);
                const y2 = Number(rangeMatch[4]);
                let safety = 0;
                while ((y1 < y2 || (y1 === y2 && m1 <= m2)) && safety < 240) {
                    if (m1 >= 1 && m1 <= 12) {
                        added.push(`${String(m1).padStart(2, "0")}/${y1}`);
                    }
                    m1++;
                    if (m1 > 12) {
                        m1 = 1;
                        y1++;
                    }
                    safety++;
                }
                return;
            }
            const singleMatch = tok.match(/^(\d{1,2})\/(\d{4})$/);
            if (singleMatch) {
                const m = Number(singleMatch[1]);
                const y = Number(singleMatch[2]);
                if (m >= 1 && m <= 12 && y >= 2400 && y <= 2700) {
                    added.push(`${String(m).padStart(2, "0")}/${y}`);
                }
            }
        });
        return normalizeAndSortPeriods(added);
    }

    function renderEditModalChips() {
        currentEditingPeriods = normalizeAndSortPeriods(currentEditingPeriods);
        if (editPeriodCount) editPeriodCount.textContent = String(currentEditingPeriods.length);
        if (editPeriodPreviewB) editPeriodPreviewB.textContent = formatPeriodRangesB(currentEditingPeriods);
        if (editPeriodPreviewLetter) editPeriodPreviewLetter.textContent = formatPeriodRangesLetter(currentEditingPeriods);

        if (!editPeriodChips) return;
        if (!currentEditingPeriods.length) {
            editPeriodChips.innerHTML = `<span style="color: #64748b; font-size: 12.5px;">ไม่มีงวดค้างชำระ (0 งวด - ชำระครบแล้ว)</span>`;
            return;
        }
        editPeriodChips.innerHTML = currentEditingPeriods.map((p) => `
            <span class="period-chip">
                ${escapeHtml(p)}
                <button type="button" class="period-chip-remove" data-remove-period="${escapeHtml(p)}" title="ลบงวด ${escapeHtml(p)}">&times;</button>
            </span>
        `).join("");

        editPeriodChips.querySelectorAll("button[data-remove-period]").forEach((btn) => {
            btn.addEventListener("click", () => {
                const target = btn.getAttribute("data-remove-period");
                currentEditingPeriods = currentEditingPeriods.filter((x) => x !== target);
                renderEditModalChips();
            });
        });
    }

    function openEditPeriodsModal(refKey) {
        const row = allTableRecords.find((r) => {
            const k1 = r["เลขอ้างอิง"] || `${r["เลขที่บัญชี"]};${r["สาขา"]}`;
            const k2 = `${r["เลขที่บัญชี"]}|${r["สาขา"]}`;
            return k1 === refKey || k2 === refKey;
        });
        if (!row || !editPeriodsModal) return;

        currentEditingRow = row;
        currentEditingPeriods = normalizeAndSortPeriods(row["งวดดิบ"] || []);

        const origPeriods = normalizeAndSortPeriods(row["งวดดิบ_ต้นฉบับ"] || []);
        const ssoPeriods = Array.isArray(row["งวดดิบ_sso"]) ? normalizeAndSortPeriods(row["งวดดิบ_sso"]) : null;

        if (editPeriodAcc) editPeriodAcc.textContent = `${row["เลขที่บัญชี"] || "-"} / สาขา ${row["สาขา"] || "000000"}`;
        if (editPeriodName) editPeriodName.textContent = row["ชื่อสถานประกอบการ (จริง)"] || row["ชื่อในรายงานหนี้"] || "-";
        if (editPeriodExcelRef) {
            editPeriodExcelRef.textContent = origPeriods.length
                ? `${origPeriods.join(", ")} (${origPeriods.length} งวด)`
                : "ไม่มี (0 งวด)";
        }
        if (editPeriodSsoRef) {
            if (ssoPeriods === null) {
                editPeriodSsoRef.textContent = "ยังไม่ได้กดตรวจสอบ SSO CORE";
                if (restoreSsoPeriodsBtn) restoreSsoPeriodsBtn.classList.add("hidden");
            } else {
                editPeriodSsoRef.textContent = ssoPeriods.length
                    ? `${ssoPeriods.join(", ")} (${ssoPeriods.length} งวด)`
                    : "ชำระครบแล้ว (0 งวด)";
                if (restoreSsoPeriodsBtn) restoreSsoPeriodsBtn.classList.remove("hidden");
            }
        }
        if (addPeriodText) addPeriodText.value = "";
        renderEditModalChips();
        editPeriodsModal.classList.remove("hidden");
    }

    function closeEditPeriodsModal() {
        if (editPeriodsModal) editPeriodsModal.classList.add("hidden");
        currentEditingRow = null;
    }

    if (closeEditPeriodsBtn) closeEditPeriodsBtn.addEventListener("click", closeEditPeriodsModal);
    if (cancelEditPeriodsBtn) cancelEditPeriodsBtn.addEventListener("click", closeEditPeriodsModal);
    if (editPeriodsModal) {
        editPeriodsModal.addEventListener("click", (e) => {
            if (e.target === editPeriodsModal) closeEditPeriodsModal();
        });
    }

    if (addSinglePeriodBtn) {
        addSinglePeriodBtn.addEventListener("click", () => {
            const m = addPeriodMonth?.value || "01";
            const y = Number(addPeriodYear?.value || 2569);
            if (y >= 2400 && y <= 2700) {
                currentEditingPeriods.push(`${String(m).padStart(2, "0")}/${y}`);
                renderEditModalChips();
            }
        });
    }

    if (addTextPeriodBtn) {
        addTextPeriodBtn.addEventListener("click", () => {
            const parsed = parseFlexiblePeriodText(addPeriodText?.value || "");
            if (!parsed.length) {
                alert("กรุณาระบุงวดในรูปแบบ ดด/ปปปป เช่น 01/2569, 02/2569 หรือช่วง 01/2569-03/2569");
                return;
            }
            currentEditingPeriods.push(...parsed);
            if (addPeriodText) addPeriodText.value = "";
            renderEditModalChips();
        });
    }

    if (restoreExcelPeriodsBtn) {
        restoreExcelPeriodsBtn.addEventListener("click", () => {
            if (!currentEditingRow) return;
            currentEditingPeriods = normalizeAndSortPeriods(currentEditingRow["งวดดิบ_ต้นฉบับ"] || []);
            renderEditModalChips();
        });
    }

    if (restoreSsoPeriodsBtn) {
        restoreSsoPeriodsBtn.addEventListener("click", () => {
            if (!currentEditingRow || !Array.isArray(currentEditingRow["งวดดิบ_sso"])) return;
            currentEditingPeriods = normalizeAndSortPeriods(currentEditingRow["งวดดิบ_sso"]);
            renderEditModalChips();
        });
    }

    if (clearAllPeriodsBtn) {
        clearAllPeriodsBtn.addEventListener("click", () => {
            currentEditingPeriods = [];
            renderEditModalChips();
        });
    }

    if (saveEditPeriodsBtn) {
        saveEditPeriodsBtn.addEventListener("click", async () => {
            if (!currentEditingRow) return;
            const finalPeriods = normalizeAndSortPeriods(currentEditingPeriods);
            const origPeriods = normalizeAndSortPeriods(currentEditingRow["งวดดิบ_ต้นฉบับ"] || []);
            const ssoPeriods = Array.isArray(currentEditingRow["งวดดิบ_sso"])
                ? normalizeAndSortPeriods(currentEditingRow["งวดดิบ_sso"])
                : null;

            currentEditingRow["งวดดิบ"] = finalPeriods;
            currentEditingRow["จำนวนงวดค้าง"] = finalPeriods.length;
            currentEditingRow["งวดเดือนค้างชำระ"] = formatPeriodRangesB(finalPeriods);

            if (ssoPeriods && JSON.stringify(finalPeriods) === JSON.stringify(ssoPeriods)) {
                currentEditingRow["แหล่งข้อมูลงวด"] = "SSO CORE";
            } else if (JSON.stringify(finalPeriods) === JSON.stringify(origPeriods)) {
                currentEditingRow["แหล่งข้อมูลงวด"] = "Excel";
            } else {
                currentEditingRow["แหล่งข้อมูลงวด"] = "แก้ไขเอง";
            }

            closeEditPeriodsModal();
            await syncRecordsAndReExport("Edited");
        });
    }

    document.addEventListener("debt:open-edit-period", (e) => {
        if (e.detail && e.detail.refKey) {
            openEditPeriodsModal(e.detail.refKey);
        }
    });

    // Start
    loadSystemInfo();
});
