/**
 * scanner.js
 * PO Scanner — Camera feed, zone-based OCR, Excel lookup
 *
 * Zone calibration: the % values in ZONES below match CSS zone positions.
 * After testing with your actual PO template, adjust these to match exactly.
 */

"use strict";

// ─── Zone Definitions ───────────────────────────────────────────────────────
// Default zones — overridden by calibration data saved in localStorage.
const DEFAULT_ZONES = {
  poNumber: { top: 0.03, left: 0.76, width: 0.22, height: 0.07 },
  supplier: { top: 0.03, left: 0.02, width: 0.30, height: 0.07 },
  items:    { top: 0.22, left: 0.02, width: 0.96, height: 0.45 },
};

const ZONES_KEY = "purchasing-tool-zones";

function loadZones() {
  try {
    const saved = localStorage.getItem(ZONES_KEY);
    if (saved) return JSON.parse(saved);
  } catch(e) {}
  return DEFAULT_ZONES;
}

let ZONES = loadZones();

// ─── State ──────────────────────────────────────────────────────────────────
let stream = null;
let currentFacingMode = "environment"; // rear camera default
let allDevices = [];
let tesseractWorker = null;
let isWorkerReady = false;
let currentMode = "camera"; // "camera" | "upload"
let uploadedImageData = null;

const video     = document.getElementById("camera-feed");
const canvas    = document.getElementById("capture-canvas");
const ctx       = canvas.getContext("2d");

// ─── Toast Notifications ────────────────────────────────────────────────────
function toast(msg, type = "info", duration = 3500) {
  const container = document.getElementById("toasts");
  const el = document.createElement("div");
  el.className = `toast toast-${type}`;
  el.innerHTML = `<span>${msg}</span>`;
  container.appendChild(el);
  setTimeout(() => el.remove(), duration);
}

// ─── Mode Toggle ────────────────────────────────────────────────────────────
function setMode(mode) {
  currentMode = mode;
  document.getElementById("camera-panel").style.display  = mode === "camera" ? "block" : "none";
  document.getElementById("upload-panel").style.display  = mode === "upload" ? "block" : "none";
  document.getElementById("tab-camera").classList.toggle("active", mode === "camera");
  document.getElementById("tab-upload").classList.toggle("active", mode === "upload");
  if (mode !== "camera") stopCamera();
}

// ─── Camera ─────────────────────────────────────────────────────────────────
async function enumerateCameras() {
  try {
    const devices = await navigator.mediaDevices.enumerateDevices();
    allDevices = devices.filter(d => d.kind === "videoinput");
    const sel = document.getElementById("camera-select");
    sel.innerHTML = "";
    allDevices.forEach((d, i) => {
      const opt = document.createElement("option");
      opt.value = d.deviceId;
      opt.textContent = d.label || `Camera ${i + 1}`;
      sel.appendChild(opt);
    });
  } catch (e) {
    console.warn("Could not enumerate cameras:", e);
  }
}

async function startCamera() {
  document.getElementById("no-camera-msg").style.display = "none";

  const constraints = {
    video: {
      facingMode: { ideal: currentFacingMode },
      width:  { ideal: 1920 },
      height: { ideal: 1080 },
    },
    audio: false,
  };

  // If a specific device is selected in the dropdown, use it
  const sel = document.getElementById("camera-select");
  if (sel.value) {
    delete constraints.video.facingMode;
    constraints.video.deviceId = { exact: sel.value };
  }

  try {
    if (stream) stopCamera();
    stream = await navigator.mediaDevices.getUserMedia(constraints);
    video.srcObject = stream;
    video.style.display = "block";

    video.onloadedmetadata = () => {
      canvas.width  = video.videoWidth;
      canvas.height = video.videoHeight;
    };

    document.getElementById("rec-dot").classList.add("live");
    document.getElementById("cam-status").textContent = "Live";
    document.getElementById("capture-btn").disabled = false;
    document.getElementById("stop-btn").disabled = false;
    document.getElementById("cam-tip").textContent = "Align PO within the brackets, then tap Capture";

    await enumerateCameras();
    await initTesseract();
  } catch (err) {
    console.error("Camera error:", err);
    toast("Camera access denied or unavailable. Try the Upload tab instead.", "error", 5000);
    document.getElementById("no-camera-msg").style.display = "flex";
  }
}

function stopCamera() {
  if (stream) {
    stream.getTracks().forEach(t => t.stop());
    stream = null;
  }
  video.srcObject = null;
  document.getElementById("rec-dot").classList.remove("live");
  document.getElementById("cam-status").textContent = "Camera inactive";
  document.getElementById("capture-btn").disabled = true;
  document.getElementById("stop-btn").disabled = true;
  document.getElementById("no-camera-msg").style.display = "flex";
}

function switchCamera() {
  currentFacingMode = currentFacingMode === "environment" ? "user" : "environment";
  if (stream) startCamera();
}

// Allow selecting camera from dropdown
document.getElementById("camera-select").addEventListener("change", () => {
  if (stream) startCamera();
});

// ─── Tesseract Initialization ────────────────────────────────────────────────
async function initTesseract() {
  if (isWorkerReady) return;
  updateOCRStatus("Loading Tesseract engine...", 10);

  try {
    tesseractWorker = await Tesseract.createWorker("eng", 1, {
      logger: m => {
        if (m.status === "recognizing text") {
          updateOCRStatus(`Recognizing... ${Math.round(m.progress * 100)}%`, m.progress * 100);
        }
      },
    });

    // Configure for printed document text
    await tesseractWorker.setParameters({
      tessedit_char_whitelist: "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.,/-: ()",
      preserve_interword_spaces: "1",
    });

    isWorkerReady = true;
    updateOCRStatus("Ready", 100);
    document.getElementById("ocr-engine-badge").textContent = "Tesseract.js · Ready";
    document.getElementById("ocr-engine-badge").className = "badge badge-green";
    setTimeout(() => { document.getElementById("ocr-progress-wrap").classList.remove("visible"); }, 1000);
  } catch (e) {
    console.error("Tesseract init failed:", e);
    updateOCRStatus("Engine failed to load", 0);
    toast("OCR engine failed to load. Check internet connection.", "error");
  }
}

function updateOCRStatus(msg, pct) {
  document.getElementById("ocr-progress-wrap").classList.add("visible");
  document.getElementById("ocr-progress-fill").style.width = pct + "%";
  document.getElementById("ocr-progress-label").textContent = msg;
  document.getElementById("ocr-status-text").textContent = msg;
}

// ─── Capture Frame from Video ────────────────────────────────────────────────
function captureFrame() {
  canvas.width  = video.videoWidth  || 1280;
  canvas.height = video.videoHeight || 720;
  ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL("image/png");
}

// ─── Crop a zone from an image (dataURL) → returns cropped dataURL ───────────
function cropZone(imageDataURL, zone) {
  return new Promise(resolve => {
    const img = new Image();
    img.onload = () => {
      const w = img.width;
      const h = img.height;

      // Add small padding for better OCR accuracy
      const pad = 4;
      const x = Math.max(0, Math.floor(zone.left  * w) - pad);
      const y = Math.max(0, Math.floor(zone.top   * h) - pad);
      const cw = Math.min(w - x, Math.ceil(zone.width  * w) + pad * 2);
      const ch = Math.min(h - y, Math.ceil(zone.height * h) + pad * 2);

      // Upscale 2× for better OCR on small zones
      const scale = 2;
      const offscreen = document.createElement("canvas");
      offscreen.width  = cw * scale;
      offscreen.height = ch * scale;
      const octx = offscreen.getContext("2d");

      // Slight contrast boost helps OCR on printed text
      octx.filter = "contrast(1.15) brightness(1.05)";
      octx.drawImage(img, x, y, cw, ch, 0, 0, cw * scale, ch * scale);
      resolve(offscreen.toDataURL("image/png"));
    };
    img.src = imageDataURL;
  });
}

// ─── Run OCR on a cropped zone ───────────────────────────────────────────────
async function ocrZone(imageDataURL, zone) {
  if (!isWorkerReady) await initTesseract();
  const cropped = await cropZone(imageDataURL, zone);
  const result  = await tesseractWorker.recognize(cropped);
  return {
    text: result.data.text.trim().replace(/\n+/g, " ").replace(/\s+/g, " "),
    confidence: result.data.confidence,
  };
}

// ─── Main: Capture from camera and scan ──────────────────────────────────────
async function captureAndScan() {
  if (!stream) { toast("Start the camera first", "warn"); return; }

  const imageData = captureFrame();
  document.getElementById("scan-line").classList.add("scanning");
  await runOCR(imageData);
  setTimeout(() => document.getElementById("scan-line").classList.remove("scanning"), 2200);
}

// ─── Main: Scan uploaded image ───────────────────────────────────────────────
async function scanUploadedImage() {
  if (!uploadedImageData) { toast("No image loaded", "warn"); return; }
  await runOCR(uploadedImageData);
}

// ─── Core OCR Pipeline ───────────────────────────────────────────────────────
async function runOCR(imageDataURL) {
  if (!isWorkerReady) {
    toast("OCR engine not ready yet, please wait...", "warn");
    await initTesseract();
  }

  // Highlight active zones
  setZoneState("zone-po",       "active");
  setZoneState("zone-supplier", "");
  setZoneState("zone-items",    "");
  updateOCRStatus("Scanning PO Number...", 15);
  document.getElementById("ocr-engine-badge").textContent = "Scanning...";
  document.getElementById("ocr-engine-badge").className = "badge badge-yellow";

  try {
    // Scan zones sequentially
    const poResult = await ocrZone(imageDataURL, ZONES.poNumber);
    setFieldValue("val-po", poResult.text, poResult.confidence, "po-conf");
    setZoneState("zone-po", "done");
    setZoneState("zone-supplier", "active");

    updateOCRStatus("Scanning Supplier Name...", 40);
    const supplierResult = await ocrZone(imageDataURL, ZONES.supplier);
    setFieldValue("val-supplier", supplierResult.text, supplierResult.confidence, "supplier-conf");
    setZoneState("zone-supplier", "done");
    setZoneState("zone-items", "active");

    updateOCRStatus("Scanning Items & Amounts...", 65);
    const itemsResult = await ocrZone(imageDataURL, ZONES.items);
    document.getElementById("val-items").value = itemsResult.text;
    setZoneState("zone-items", "done");

    updateOCRStatus("Scan complete ✓", 100);
    document.getElementById("ocr-engine-badge").textContent = "Tesseract.js · Done";
    document.getElementById("ocr-engine-badge").className = "badge badge-green";
    toast("OCR complete! Review extracted data.", "success");

    // Auto-lookup after scan
    setTimeout(lookupPO, 500);
  } catch (err) {
    console.error("OCR error:", err);
    updateOCRStatus("OCR failed — try better lighting", 0);
    document.getElementById("ocr-engine-badge").className = "badge badge-red";
    toast("OCR failed. Ensure good lighting and try again.", "error");
  }
}

function setZoneState(id, state) {
  const el = document.getElementById(id);
  if (!el) return;
  el.classList.remove("active", "done");
  if (state) el.classList.add(state);
}

function setFieldValue(inputId, text, confidence, badgeId) {
  document.getElementById(inputId).value = text;
  const badge = document.getElementById(badgeId);
  if (badge) {
    const pct = Math.round(confidence);
    badge.textContent = `${pct}%`;
    badge.className = `badge ${pct >= 80 ? "badge-green" : pct >= 60 ? "badge-yellow" : "badge-red"}`;
  }
}

// ─── Excel Lookup ────────────────────────────────────────────────────────────
async function lookupPO() {
  const poNumber = document.getElementById("val-po").value.trim();
  if (!poNumber) { toast("Enter or scan a PO number first", "warn"); return; }

  const resultCard = document.getElementById("result-card");
  resultCard.style.display = "block";
  document.getElementById("result-body").innerHTML = `
    <div style="padding:20px;text-align:center;color:var(--text-2);">
      <div style="font-size:24px;margin-bottom:8px;">🔍</div>
      Looking up ${poNumber}...
    </div>`;
  document.getElementById("result-badge").textContent = "Searching";
  document.getElementById("result-badge").className = "badge badge-grey";
  document.getElementById("result-actions").innerHTML = "";

  try {
    const res  = await fetch(`/api/po/${encodeURIComponent(poNumber)}`);
    const data = await res.json();

    if (data.found) {
      showMatchResult(data.row, data.sheet);
    } else {
      showNoMatchResult(poNumber);
    }
  } catch (e) {
    document.getElementById("result-badge").textContent = "Server Error";
    document.getElementById("result-badge").className = "badge badge-red";
    document.getElementById("result-body").innerHTML = `
      <div style="padding:16px;color:var(--red);font-size:13px;">
        ⚠️ Could not connect to server. Is <code>node server.js</code> running?
      </div>`;
  }
}

function showMatchResult(row, sheet) {
  const scannedSupplier = document.getElementById("val-supplier").value.trim();
  const scannedItems    = document.getElementById("val-items").value.trim();

  const xlSupplier = row["SUPPLIER'S NAME"] || "";
  const xlAmount   = row["TOTAL AMOUNT"]    || "";
  const xlStatus   = row["PURCHASE ORDER STATUS"] || "";

  // Simple comparison
  const supplierMatch = normalize(scannedSupplier) === normalize(xlSupplier)
    || normalize(xlSupplier).includes(normalize(scannedSupplier))
    || normalize(scannedSupplier).includes(normalize(xlSupplier));

  const hasMismatch = scannedSupplier && !supplierMatch;

  document.getElementById("result-badge").textContent = hasMismatch ? "⚠ Discrepancy" : "✓ Match";
  document.getElementById("result-badge").className = `badge ${hasMismatch ? "badge-red" : "badge-green"}`;

  const rows = [
    { field: "PO Number",       scanned: document.getElementById("val-po").value, excel: row["PO NUMBER"]       || row["PO Number"]       || "—", match: true },
    { field: "Supplier Name",   scanned: scannedSupplier, excel: xlSupplier,       match: supplierMatch },
    { field: "Sheet (Month)",   scanned: "—",             excel: sheet,            match: true },
    { field: "PO Status",       scanned: "—",             excel: xlStatus,         match: true },
    { field: "Total Amount",    scanned: "—",             excel: xlAmount,         match: true },
    { field: "Requesting Dept", scanned: "—",             excel: row["REQUESTING DEPT."] || "—", match: true },
  ];

  let tableHTML = `
    <div class="table-wrap" style="border-radius:0;border:none;">
      <table>
        <thead>
          <tr>
            <th>Field</th>
            <th>Scanned</th>
            <th>In Excel</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
  `;

  rows.forEach(r => {
    const statusIcon = r.scanned === "—" ? "—" : (r.match ? "✓" : "✗");
    const cls = r.scanned === "—" ? "" : (r.match ? "diff-row-match" : "diff-row-mismatch");
    tableHTML += `
      <tr>
        <td class="mono" style="font-size:11px;text-transform:uppercase;letter-spacing:.04em;color:var(--text-2);">${r.field}</td>
        <td style="font-family:var(--font-mono);font-size:12px;">${r.scanned || "—"}</td>
        <td style="font-family:var(--font-mono);font-size:12px;">${r.excel || "—"}</td>
        <td class="${cls}" style="font-family:var(--font-mono);font-size:13px;">${statusIcon}</td>
      </tr>`;
  });

  tableHTML += "</tbody></table></div>";

  if (hasMismatch) {
    tableHTML += `<div style="padding:12px 16px;background:#3a1a1a;border-top:1px solid #5a2d2d;font-size:13px;color:var(--red);">
      ⚠️ Supplier name mismatch detected. Verify the physical PO before proceeding.
    </div>`;
  } else {
    tableHTML += `<div style="padding:12px 16px;background:#1a3a1e;border-top:1px solid #2d5a32;font-size:13px;color:var(--green);">
      ✅ PO found in monitoring file — sheet: ${sheet}
    </div>`;
  }

  document.getElementById("result-body").innerHTML = tableHTML;

  const actions = document.getElementById("result-actions");
  actions.innerHTML = `
    <button class="btn btn-secondary" style="font-size:12px;" onclick="copyToClipboard()">📋 Copy PO Data</button>
  `;

  toast(hasMismatch ? "Discrepancy found! Check result." : "PO matched successfully.", hasMismatch ? "warn" : "success");
}

function showNoMatchResult(poNumber) {
  document.getElementById("result-badge").textContent = "Not Found";
  document.getElementById("result-badge").className = "badge badge-yellow";
  document.getElementById("result-body").innerHTML = `
    <div style="padding:20px 16px;">
      <div style="font-size:28px;margin-bottom:8px;">🔎</div>
      <div style="font-size:14px;font-weight:600;margin-bottom:6px;">PO not found in monitoring file</div>
      <div style="font-size:13px;color:var(--text-2);">
        <strong>${poNumber}</strong> was not found in any monthly sheet. 
        You can add it as a new row below.
      </div>
    </div>`;

  document.getElementById("result-actions").innerHTML = `
    <button class="btn btn-primary" onclick="openAddModal()">➕ Add New Row</button>
  `;

  toast(`PO "${poNumber}" not found. You can add it.`, "warn");
}

// ─── Add New Row Modal ───────────────────────────────────────────────────────
const COLUMNS = [
  "PR DATE", "PR DATE RECEIVED", "PR NO.", "REQUESTING DEPT.",
  "PO DATE", "PO NUMBER", "END USER/S", "SUPPLIER'S NAME",
  "ITEM CODE", "ITEM DESCRIPTION", "SPECIFICATIONS", "QTY", "UoM",
  "UNIT PRICE", "AMOUNT", "TOTAL AMOUNT", "PR REQUIRED DATE",
  "DATE DELIVERED", "REMARKS", "PURCHASE ORDER STATUS",
  "ITEMS/SERVICES", "ORIGINAL PRICE", "TOTAL COST SAVINGS"
];

function openAddModal() {
  const prefill = {
    "PO NUMBER":       document.getElementById("val-po").value,
    "SUPPLIER'S NAME": document.getElementById("val-supplier").value,
    "REQUESTING DEPT.":document.getElementById("val-dept").value,
    "PO DATE":         document.getElementById("val-po-date").value,
    "ITEMS/SERVICES":  document.getElementById("val-items").value,
  };

  const container = document.getElementById("add-form-fields");
  container.innerHTML = "";

  COLUMNS.forEach(col => {
    const div = document.createElement("div");
    div.innerHTML = `
      <div class="form-label">${col}</div>
      <input class="form-input" type="text" id="addrow-${sanitizeId(col)}"
        value="${escapeHTML(prefill[col] || "")}"
        placeholder="${col}" />
    `;
    container.appendChild(div);
  });

  document.getElementById("add-modal").classList.add("open");
}

async function saveNewRow() {
  const month = document.getElementById("val-month").value;
  const rowData = {};

  COLUMNS.forEach(col => {
    const el = document.getElementById("addrow-" + sanitizeId(col));
    if (el) rowData[col] = el.value;
  });

  try {
    const res = await fetch(`/api/sheet/${month}/add`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(rowData),
    });
    const data = await res.json();
    if (data.success) {
      toast("Row added to " + month + " sheet successfully!", "success");
      closeModal("add-modal");
      document.getElementById("result-card").style.display = "none";
    } else {
      toast("Failed to save row", "error");
    }
  } catch (e) {
    toast("Server error — could not save row", "error");
  }
}

// ─── File Upload ─────────────────────────────────────────────────────────────
function handleFileSelect(event) {
  const file = event.target.files[0];
  if (file) loadImageFile(file);
}

function handleDrop(event) {
  event.preventDefault();
  event.currentTarget.style.borderColor = "var(--border)";
  const file = event.dataTransfer.files[0];
  if (file && file.type.startsWith("image/")) {
    loadImageFile(file);
  } else {
    toast("Please drop an image file", "warn");
  }
}

function loadImageFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    uploadedImageData = e.target.result;
    document.getElementById("preview-img").src = uploadedImageData;
    document.getElementById("upload-preview").style.display = "block";
    document.getElementById("drop-zone").style.display = "none";
    initTesseract(); // pre-load engine
  };
  reader.readAsDataURL(file);
}

function clearUpload() {
  uploadedImageData = null;
  document.getElementById("upload-preview").style.display = "none";
  document.getElementById("drop-zone").style.display = "block";
  document.getElementById("file-input").value = "";
}

// ─── Utilities ───────────────────────────────────────────────────────────────
function clearFields() {
  ["val-po", "val-supplier", "val-items", "val-dept"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
  ["po-conf", "supplier-conf"].forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.textContent = ""; el.className = "badge badge-grey"; }
  });
  ["field-po", "field-supplier", "field-items"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.className = "ocr-field";
  });
  document.getElementById("result-card").style.display = "none";
  setZoneState("zone-po", ""); setZoneState("zone-supplier", ""); setZoneState("zone-items", "");
}

function normalize(str) {
  return String(str || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function sanitizeId(str) {
  return str.replace(/[^a-zA-Z0-9]/g, "_");
}

function escapeHTML(str) {
  return String(str).replace(/"/g, "&quot;");
}

function closeModal(id) {
  document.getElementById(id).classList.remove("open");
}

function copyToClipboard() {
  const po = document.getElementById("val-po").value;
  const supplier = document.getElementById("val-supplier").value;
  const items = document.getElementById("val-items").value;
  const text = `PO Number: ${po}\nSupplier: ${supplier}\nItems: ${items}`;
  navigator.clipboard.writeText(text).then(() => toast("Copied to clipboard", "success"));
}

// ─── Close modal on overlay click ────────────────────────────────────────────
document.querySelectorAll(".modal-overlay").forEach(overlay => {
  overlay.addEventListener("click", e => {
    if (e.target === overlay) overlay.classList.remove("open");
  });
});

// ─── Set default month to current month ──────────────────────────────────────
const MONTH_NAMES = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
document.getElementById("val-month").value = MONTH_NAMES[new Date().getMonth()];