// monitoring-clean.js (v2 - real upload flow)
// Wire this up in dashboard.html: <script src="/monitoring-clean.js"></script>

(function () {
  const btn = document.getElementById("cleanMonitoringBtn");
  const modal = document.getElementById("cleanConfirmModal");
  const modalDetail = document.getElementById("cleanModalDetail");
  const cancelBtn = document.getElementById("cleanModalCancel");
  const confirmBtn = document.getElementById("cleanModalConfirm");

  if (!btn) return;

  // Hidden file input, created once, reused on every click
  const fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.accept = ".xlsx";
  fileInput.style.display = "none";
  document.body.appendChild(fileInput);

  let pendingFileId = null; // set after a successful upload+inspect

  function openModal(detailText) {
    modalDetail.textContent = detailText;
    modal.classList.add("visible");
    if (window.lucide) window.lucide.createIcons();
  }

  function closeModal() {
    modal.classList.remove("visible");
  }

  function setBtnState(text, disabled) {
    btn.textContent = text;
    btn.disabled = !!disabled;
  }

  const DEFAULT_BTN_HTML = '<i data-lucide="sparkles" style="width:14px;height:14px;"></i> Clean Monitoring Data';

  function resetBtn() {
    btn.innerHTML = DEFAULT_BTN_HTML;
    btn.disabled = false;
    if (window.lucide) window.lucide.createIcons();
  }

  btn.addEventListener("click", () => {
    fileInput.value = ""; // allow re-selecting the same filename
    fileInput.click();
  });

  function startTimer(onTick) {
    const t0 = Date.now();
    const id = setInterval(() => onTick(((Date.now() - t0) / 1000).toFixed(0)), 1000);
    return () => clearInterval(id);
  }

  fileInput.addEventListener("change", async () => {
    const file = fileInput.files[0];
    if (!file) return;

    const stopTimer = startTimer((s) => setBtnState(`Analyzing... (${s}s)`, true));

    try {
      const formData = new FormData();
      formData.append("file", file);

      const res = await fetch("/api/monitoring/upload", { method: "POST", body: formData });
      const data = await res.json();
      stopTimer();

      if (!res.ok) throw new Error(data.error || "Upload failed");

      pendingFileId = data.fileId;

      if (!data.dirty) {
        resetBtn();
        btn.textContent = "Already clean ✓";
        setTimeout(resetBtn, 1500);
        return;
      }

      const mergedTotal = data.totalMergedCells;
      const phantomSheets = data.sheets.filter(s => s.phantomRows).length;
      const detail = `Found ${mergedTotal} merged-cell artifact(s) across ${data.sheetCount} sheets` +
        (phantomSheets ? ` (${phantomSheets} sheet(s) have inflated row counts from formatting).` : ".") +
        ` Cleaning will unmerge cells, trim empty rows, and overwrite monitoring.xlsx + dashboard-cache.json. ` +
        `This can take up to a minute for large files — please don't close this tab.`;

      resetBtn();
      openModal(detail);
    } catch (err) {
      stopTimer();
      console.error("upload/inspect failed:", err);
      alert("Upload failed: " + err.message);
      resetBtn();
    }
  });

  cancelBtn.addEventListener("click", () => {
    closeModal();
    pendingFileId = null;
  });

  confirmBtn.addEventListener("click", async () => {
    if (!pendingFileId) return closeModal();

    confirmBtn.disabled = true;
    cancelBtn.disabled = true;
    const stopTimer = startTimer((s) => {
      confirmBtn.textContent = `Cleaning... (${s}s)`;
    });

    try {
      const res = await fetch("/api/monitoring/clean", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fileId: pendingFileId }),
      });
      const result = await res.json();
      stopTimer();

      if (!res.ok || !result.ok) throw new Error(result.error || "Cleanup failed");

      closeModal();
      pendingFileId = null;

      btn.textContent = `Cleaned ✓ (${result.rawRows} rows)`;
      setTimeout(resetBtn, 2000);

      // Reload the dashboard so charts reflect the newly cleaned data
      if (typeof window.loadDashboard === "function") {
        window.loadDashboard(true);
      } else {
        location.reload();
      }
    } catch (err) {
      stopTimer();
      console.error("clean failed:", err);
      alert("Cleanup failed: " + err.message);
    } finally {
      confirmBtn.disabled = false;
      cancelBtn.disabled = false;
      confirmBtn.textContent = "Yes, clean it";
    }
  });

  modal.addEventListener("click", (e) => {
    if (e.target === modal) {
      closeModal();
      pendingFileId = null;
    }
  });
})();