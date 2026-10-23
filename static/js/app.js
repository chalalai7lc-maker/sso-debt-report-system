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
        allTableRecords = data.preview_data || [];
        tableCountLabel.textContent = `${allTableRecords.length} รายการ`;
        renderTable(allTableRecords);

        document.dispatchEvent(new CustomEvent("debt:ready", { detail: allTableRecords }));

        // Show section & scroll
        resultsSection.classList.remove("hidden");
        resultsSection.scrollIntoView({ behavior: "smooth" });
    }

    function renderTable(records) {
        tableBody.innerHTML = "";
        if (records.length === 0) {
            tableBody.innerHTML = `<tr><td colspan="13" style="text-align: center; padding: 24px; color: var(--text-muted);">ไม่พบข้อมูลที่ตรงกับเงื่อนไขค้นหา</td></tr>`;
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

            tr.innerHTML = `
                <td class="text-center cell-mono">${idx + 1}</td>
                <td class="cell-mono"><strong>${escapeHtml(row["เลขอ้างอิง"] || "")}</strong></td>
                <td><strong>${nameDisplay}</strong></td>
                <td>${escapeHtml(row["ตำบล"] || "-")}</td>
                <td>${escapeHtml(row["อำเภอ"] || "-")}</td>
                <td>${escapeHtml(row["จังหวัด"] || "-")}</td>
                <td class="text-center cell-mono">${escapeHtml(row["รหัสไปรษณีย์"] || "-")}</td>
                <td class="text-center cell-mono">${escapeHtml(row["จำนวนงวดค้าง"] ?? "")}</td>
                <td>${escapeHtml(row["งวดเดือนค้างชำระ"] || "")}</td>
                <td class="text-right cell-debt">${formatCurrency(row["ยอดหนี้ไม่รวมเงินเพิ่ม"])}</td>
                <td class="text-right cell-mono">${formatCurrency(row["เงินเพิ่ม"])}</td>
                <td class="text-right cell-total">${formatCurrency(row["ยอดค้างชำระรวมทั้งสิ้น"])}</td>
                <td class="text-center"><span class="status-pill ${statusClass}">${escapeHtml(row["สถานะ สปก."] || "-")}</span></td>
            `;
            tableBody.appendChild(tr);
        });
    }

    // 5. Live Search Filter
    searchInput.addEventListener("input", (e) => {
        const query = e.target.value.trim().toLowerCase();
        if (!query) {
            renderTable(allTableRecords);
            tableCountLabel.textContent = `${allTableRecords.length} รายการ`;
            return;
        }

        const filtered = allTableRecords.filter(r => {
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

        renderTable(filtered);
        tableCountLabel.textContent = `${filtered.length} / ${allTableRecords.length} รายการ`;
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

    // Start
    loadSystemInfo();
});
