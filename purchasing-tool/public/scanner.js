/**
 * scanner.js — Merged capture + calibrate + scan flow
 *
 * Step 1: Capture  — live camera OR file upload
 * Step 2: Calibrate — drag/resize zone boxes on the frozen image
 * Step 3: Scan     — OCR runs, results appear, Excel lookup
 */

"use strict";

// ─── Zones ───────────────────────────────────────────────────────────────────
const ZONES_KEY = "purchasing-tool-zones";

const DEFAULT_ZONES = {
  poNumber: { top: 0.03, left: 0.76, width: 0.22, height: 0.07 },
  supplier: { top: 0.03, left: 0.02, width: 0.30, height: 0.07 },
  items:    { top: 0.22, left: 0.02, width: 0.96, height: 0.45 },
};

const ZONE_META = {
  poNumber: { label: "PO Number",      color: "#388bfd" },
  supplier: { label: "Supplier Name",  color: "#3fb950" },
  items:    { label: "Items & Amounts",color: "#d29922" },
};

function loadZones() {
  try { const s = localStorage.getItem(ZONES_KEY); if (s) return JSON.parse(s); } catch(e) {}
  return JSON.parse(JSON.stringify(DEFAULT_ZONES));
}
function saveZonesToStorage() {
  localStorage.setItem(ZONES_KEY, JSON.stringify(zones));
}

let zones = loadZones();

// ─── State ───────────────────────────────────────────────────────────────────
let currentStep       = 1;   // 1=capture, 2=calibrate, 3=scan
let stream            = null;
let currentFacingMode = "environment";
let capturedImageData = null;

let tesseractWorker   = null;
let isWorkerReady     = false;

// drag/resize
let dragState = null;

// ─── DOM refs ────────────────────────────────────────────────────────────────
const videoEl       = document.getElementById("camera-feed");
const frozenImg     = document.getElementById("frozen-image");
const captureCanvas = document.getElementById("capture-canvas");
const captureCtx    = captureCanvas.getContext("2d");
const viewportWrap  = document.getElementById("viewport-wrap");
const viewportEmpty = document.getElementById("viewport-empty");

// ─── Toast ───────────────────────────────────────────────────────────────────
function toast(msg, type = "info", duration = 3500) {
  const el = document.createElement("div");
  el.className = `toast toast-${type}`;
  el.textContent = msg;
  document.getElementById("toasts").appendChild(el);
  setTimeout(() => el.remove(), duration);
}

// ═════════════════════════════════════════════════════════════════════════════
// STEP MANAGEMENT
// ═════════════════════════════════════════════════════════════════════════════
function goToStep(n) {
  currentStep = n;

  // Update step indicators
  [1,2,3].forEach(i => {
    const s = document.getElementById(`step-${i}`);
    const num = document.getElementById(`step-num-${i}`);
    s.classList.remove("active","done");
    if (i < n) { s.classList.add("done"); num.innerHTML = `<i data-lucide="check" style="width:10px;height:10px;"></i>`; }
    else if (i === n) s.classList.add("active");
  });
  if (window.lucide) lucide.createIcons();

  if (n === 1) renderStep1();
  if (n === 2) renderStep2();
  if (n === 3) renderStep3();
}

// ═════════════════════════════════════════════════════════════════════════════
// STEP 1 — CAPTURE
// ═════════════════════════════════════════════════════════════════════════════
function renderStep1() {
  setTitle("Ready to capture", false);
  setHeaderControls(`
    <div class="mode-tabs">
      <button class="mode-tab active" id="tab-cam" onclick="switchInputMode('camera')">
        <i data-lucide="camera" style="width:13px;height:13px;"></i> Camera
      </button>
      <button class="mode-tab" id="tab-upload" onclick="switchInputMode('upload')">
        <i data-lucide="upload" style="width:13px;height:13px;"></i> Upload
      </button>
    </div>
  `);
  setToolbar(`
    <button class="btn btn-primary" id="capture-btn" onclick="doCapture()" disabled>
      <i data-lucide="aperture" style="width:14px;height:14px;"></i> Capture Photo
    </button>
    <button class="btn btn-secondary" id="stop-btn" onclick="stopCamera()" disabled>
      <i data-lucide="square" style="width:13px;height:13px;"></i> Stop
    </button>
    <input type="file" id="file-input" accept="image/*" style="display:none" onchange="handleFileSelect(event)">
    <span class="toolbar-hint" id="toolbar-hint">Start camera or upload an image</span>
  `);
  hideCalibrateBar();
  showBrackets(true);
  setViewportInteractive(false);
  showZones(false);

  // Show empty state if no image yet
  if (!capturedImageData) {
    viewportEmpty.style.display = "flex";
    frozenImg.style.display = "none";
    videoEl.style.display = "none";
  }
  if (window.lucide) lucide.createIcons();
}

function switchInputMode(mode) {
  document.getElementById("tab-cam").classList.toggle("active", mode === "camera");
  document.getElementById("tab-upload").classList.toggle("active", mode === "upload");
  if (mode === "camera") {
    startCamera();
  } else {
    stopCamera();
    document.getElementById("file-input").click();
  }
  if (window.lucide) lucide.createIcons();
}

async function startCamera() {
  const constraints = {
    video: { facingMode: { ideal: currentFacingMode }, width: { ideal: 1920 }, height: { ideal: 1080 } },
    audio: false,
  };
  try {
    if (stream) stopCamera();
    stream = await navigator.mediaDevices.getUserMedia(constraints);
    videoEl.srcObject = stream;
    videoEl.style.display = "block";
    frozenImg.style.display = "none";
    viewportEmpty.style.display = "none";
    videoEl.onloadedmetadata = () => {
      captureCanvas.width  = videoEl.videoWidth;
      captureCanvas.height = videoEl.videoHeight;
    };
    setTitle("Live", true);
    document.getElementById("capture-btn").disabled = false;
    document.getElementById("stop-btn").disabled = false;
    document.getElementById("toolbar-hint").textContent = "Align PO with the brackets, then Capture";
    initTesseract();
  } catch(err) {
    toast("Camera access denied. Try uploading an image instead.", "error", 5000);
  }
  if (window.lucide) lucide.createIcons();
}

function stopCamera() {
  if (stream) { stream.getTracks().forEach(t => t.stop()); stream = null; }
  videoEl.srcObject = null;
  videoEl.style.display = "none";
  setTitle("Camera stopped", false);
  const btn = document.getElementById("capture-btn");
  const stop = document.getElementById("stop-btn");
  if (btn) btn.disabled = true;
  if (stop) stop.disabled = true;
}

function switchCamera() {
  currentFacingMode = currentFacingMode === "environment" ? "user" : "environment";
  if (stream) startCamera();
}

function doCapture() {
  if (!stream) return;
  captureCanvas.width  = videoEl.videoWidth  || 1280;
  captureCanvas.height = videoEl.videoHeight || 720;
  captureCtx.drawImage(videoEl, 0, 0, captureCanvas.width, captureCanvas.height);
  capturedImageData = captureCanvas.toDataURL("image/png");
  stopCamera();
  loadFrozenImage(capturedImageData);
  goToStep(2);
}

function handleFileSelect(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    capturedImageData = e.target.result;
    loadFrozenImage(capturedImageData);
    goToStep(2);
  };
  reader.readAsDataURL(file);
}

function loadFrozenImage(src) {
  frozenImg.src = src;
  frozenImg.style.display = "block";
  videoEl.style.display = "none";
  viewportEmpty.style.display = "none";
}

// ═════════════════════════════════════════════════════════════════════════════
// STEP 2 — CALIBRATE
// ═════════════════════════════════════════════════════════════════════════════
function renderStep2() {
  setTitle("Calibrate scan zones", false);
  setHeaderControls(`
    <button class="btn btn-secondary" style="font-size:12px;display:flex;align-items:center;gap:5px;" onclick="resetZones()">
      <i data-lucide="rotate-ccw" style="width:12px;height:12px;"></i> Reset
    </button>
    <button class="btn btn-secondary" style="font-size:12px;display:flex;align-items:center;gap:5px;" onclick="goToStep(1)">
      <i data-lucide="arrow-left" style="width:12px;height:12px;"></i> Retake
    </button>
  `);
  setToolbar(`
    <button class="btn btn-primary" onclick="confirmCalibration()" style="display:flex;align-items:center;gap:7px;">
      <i data-lucide="check" style="width:14px;height:14px;"></i> Zones Look Good — Scan Now
    </button>
    <span class="toolbar-hint">Drag boxes to reposition, pull edges to resize</span>
  `);
  showCalibrateBar();
  showBrackets(false);
  setViewportInteractive(true);
  buildZoneBoxes();
  showZones(true);
  if (window.lucide) lucide.createIcons();
}

function buildZoneBoxes() {
  // Remove old zone boxes
  viewportWrap.querySelectorAll(".zone-box").forEach(el => el.remove());

  Object.entries(zones).forEach(([key, z]) => {
    const meta = ZONE_META[key];
    const box = document.createElement("div");
    box.className = "zone-box";
    box.dataset.zone = key;
    positionZoneBox(box, z);

    // Label
    const lbl = document.createElement("div");
    lbl.className = "zone-label";
    lbl.textContent = meta.label;
    box.appendChild(lbl);

    // Resize handles
    ["nw","n","ne","e","se","s","sw","w"].forEach(dir => {
      const h = document.createElement("div");
      h.className = `rh rh-${dir}`;
      h.dataset.handle = dir;
      h.addEventListener("mousedown",  e => startResize(e, key, dir));
      h.addEventListener("touchstart", e => startResizeTouch(e, key, dir), { passive: false });
      box.appendChild(h);
    });

    box.addEventListener("mousedown",  e => startMove(e, key));
    box.addEventListener("touchstart", e => startMoveTouch(e, key), { passive: false });

    viewportWrap.appendChild(box);
  });
}

function positionZoneBox(box, z) {
  box.style.left   = (z.left   * 100) + "%";
  box.style.top    = (z.top    * 100) + "%";
  box.style.width  = (z.width  * 100) + "%";
  box.style.height = (z.height * 100) + "%";
}

function updateZoneBox(key) {
  const box = viewportWrap.querySelector(`.zone-box[data-zone="${key}"]`);
  if (box) positionZoneBox(box, zones[key]);
}

function resetZones() {
  zones = JSON.parse(JSON.stringify(DEFAULT_ZONES));
  buildZoneBoxes();
  toast("Zones reset to defaults.", "info");
}

function confirmCalibration() {
  saveZonesToStorage();
  toast("Zones saved. Starting scan...", "success");
  goToStep(3);
}

// ─── Drag / Resize (mouse) ────────────────────────────────────────────────────
function getWrapRect() { return viewportWrap.getBoundingClientRect(); }

function startMove(e, key) {
  if (e.target.classList.contains("rh")) return;
  e.preventDefault();
  const r = getWrapRect();
  dragState = { type:"move", zone:key, startX:e.clientX, startY:e.clientY,
    startZone:{...zones[key]}, wrapW:r.width, wrapH:r.height };
}
function startResize(e, key, handle) {
  e.preventDefault(); e.stopPropagation();
  const r = getWrapRect();
  dragState = { type:"resize", zone:key, handle, startX:e.clientX, startY:e.clientY,
    startZone:{...zones[key]}, wrapW:r.width, wrapH:r.height };
}
function startMoveTouch(e, key) {
  if (e.target.classList.contains("rh")) return;
  e.preventDefault();
  const t = e.touches[0]; const r = getWrapRect();
  dragState = { type:"move", zone:key, startX:t.clientX, startY:t.clientY,
    startZone:{...zones[key]}, wrapW:r.width, wrapH:r.height };
}
function startResizeTouch(e, key, handle) {
  e.preventDefault(); e.stopPropagation();
  const t = e.touches[0]; const r = getWrapRect();
  dragState = { type:"resize", zone:key, handle, startX:t.clientX, startY:t.clientY,
    startZone:{...zones[key]}, wrapW:r.width, wrapH:r.height };
}

document.addEventListener("mousemove", e => { if (dragState) handleDrag(e.clientX, e.clientY); });
document.addEventListener("mouseup",   () => { dragState = null; });
document.addEventListener("touchmove", e => {
  if (!dragState) return;
  e.preventDefault();
  handleDrag(e.touches[0].clientX, e.touches[0].clientY);
}, { passive: false });
document.addEventListener("touchend", () => { dragState = null; });

function handleDrag(cx, cy) {
  if (!dragState) return;
  const dx = (cx - dragState.startX) / dragState.wrapW;
  const dy = (cy - dragState.startY) / dragState.wrapH;
  const sz = dragState.startZone;
  const key = dragState.zone;
  let { left, top, width, height } = sz;

  if (dragState.type === "move") {
    left = clamp(sz.left + dx, 0, 1 - sz.width);
    top  = clamp(sz.top  + dy, 0, 1 - sz.height);
  } else {
    const h = dragState.handle;
    if (h.includes("e")) width  = clamp(sz.width  + dx, 0.02, 1 - sz.left);
    if (h.includes("s")) height = clamp(sz.height + dy, 0.02, 1 - sz.top);
    if (h.includes("w")) {
      const nl = clamp(sz.left + dx, 0, sz.left + sz.width - 0.02);
      width = sz.left + sz.width - nl; left = nl;
    }
    if (h.includes("n")) {
      const nt = clamp(sz.top + dy, 0, sz.top + sz.height - 0.02);
      height = sz.top + sz.height - nt; top = nt;
    }
  }

  zones[key] = { left, top, width, height };
  updateZoneBox(key);
}

function clamp(v, min, max) { return Math.max(min, Math.min(max, v)); }

// ═════════════════════════════════════════════════════════════════════════════
// STEP 3 — SCAN
// ═════════════════════════════════════════════════════════════════════════════
function renderStep3() {
  setTitle("Scanning...", false);
  setHeaderControls(`
    <button class="btn btn-secondary" style="font-size:12px;display:flex;align-items:center;gap:5px;" onclick="goToStep(2)">
      <i data-lucide="move" style="width:12px;height:12px;"></i> Adjust Zones
    </button>
    <button class="btn btn-secondary" style="font-size:12px;display:flex;align-items:center;gap:5px;" onclick="goToStep(1)">
      <i data-lucide="refresh-cw" style="width:12px;height:12px;"></i> New Scan
    </button>
  `);
  setToolbar(`<span class="toolbar-hint" id="toolbar-hint">Running OCR...</span>`);
  hideCalibrateBar();
  showBrackets(false);
  setViewportInteractive(false);
  showZones(true);
  if (window.lucide) lucide.createIcons();

  // Kick off OCR
  runOCR(capturedImageData);
}

// ─── Zone state highlight ─────────────────────────────────────────────────────
function setZoneState(key, state) {
  const box = viewportWrap.querySelector(`.zone-box[data-zone="${key}"]`);
  if (!box) return;
  box.classList.remove("zone-active","zone-done");
  if (state) box.classList.add(`zone-${state}`);
}

// ─── Tesseract ────────────────────────────────────────────────────────────────
async function initTesseract() {
  if (isWorkerReady) return;
  updateOCRStatus("Loading OCR engine...", 10);
  try {
    tesseractWorker = await Tesseract.createWorker("eng", 1, {
      logger: m => {
        if (m.status === "recognizing text")
          updateOCRStatus(`Recognizing... ${Math.round(m.progress * 100)}%`, m.progress * 100);
      },
    });
    await tesseractWorker.setParameters({
      tessedit_char_whitelist: "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.,/-: ()",
      preserve_interword_spaces: "1",
    });
    isWorkerReady = true;
    updateOCRStatus("Ready", 100);
    document.getElementById("ocr-engine-badge").textContent = "Tesseract.js · Ready";
    document.getElementById("ocr-engine-badge").className = "badge badge-green";
    setTimeout(() => document.getElementById("ocr-progress-wrap").classList.remove("visible"), 1000);
  } catch(e) {
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

async function cropZone(imageDataURL, zone) {
  return new Promise(resolve => {
    const img = new Image();
    img.onload = () => {
      const w = img.width, h = img.height, pad = 4;
      const x  = Math.max(0, Math.floor(zone.left   * w) - pad);
      const y  = Math.max(0, Math.floor(zone.top    * h) - pad);
      const cw = Math.min(w - x, Math.ceil(zone.width  * w) + pad * 2);
      const ch = Math.min(h - y, Math.ceil(zone.height * h) + pad * 2);
      const off = document.createElement("canvas");
      off.width = cw * 2; off.height = ch * 2;
      const octx = off.getContext("2d");
      octx.filter = "contrast(1.15) brightness(1.05)";
      octx.drawImage(img, x, y, cw, ch, 0, 0, cw * 2, ch * 2);
      resolve(off.toDataURL("image/png"));
    };
    img.src = imageDataURL;
  });
}

async function ocrZone(imageDataURL, zone) {
  if (!isWorkerReady) await initTesseract();
  const cropped = await cropZone(imageDataURL, zone);
  const result  = await tesseractWorker.recognize(cropped);
  return { text: result.data.text.trim().replace(/\s+/g, " "), confidence: result.data.confidence };
}

async function runOCR(imageDataURL) {
  if (!isWorkerReady) await initTesseract();

  document.getElementById("scan-line").classList.add("scanning");
  document.getElementById("ocr-engine-badge").textContent = "Scanning...";
  document.getElementById("ocr-engine-badge").className = "badge badge-yellow";

  try {
    setZoneState("poNumber","active");
    updateOCRStatus("Scanning PO Number...", 15);
    const po = await ocrZone(imageDataURL, zones.poNumber);
    setFieldValue("val-po", po.text, po.confidence, "po-conf");
    setZoneState("poNumber","done");

    setZoneState("supplier","active");
    updateOCRStatus("Scanning Supplier Name...", 45);
    const sup = await ocrZone(imageDataURL, zones.supplier);
    setFieldValue("val-supplier", sup.text, sup.confidence, "supplier-conf");
    setZoneState("supplier","done");

    setZoneState("items","active");
    updateOCRStatus("Scanning Items & Amounts...", 70);
    const itm = await ocrZone(imageDataURL, zones.items);
    document.getElementById("val-items").value = itm.text;
    setZoneState("items","done");

    document.getElementById("scan-line").classList.remove("scanning");
    updateOCRStatus("Scan complete", 100);
    document.getElementById("ocr-engine-badge").textContent = "Tesseract.js · Done";
    document.getElementById("ocr-engine-badge").className = "badge badge-green";
    setTitle("Scan complete", false);

    const hint = document.getElementById("toolbar-hint");
    if (hint) hint.textContent = "Review extracted data on the right, then check in Excel";

    toast("OCR complete! Review extracted data.", "success");
    setTimeout(lookupPO, 400);

  } catch(err) {
    console.error("OCR error:", err);
    document.getElementById("scan-line").classList.remove("scanning");
    updateOCRStatus("OCR failed", 0);
    document.getElementById("ocr-engine-badge").className = "badge badge-red";
    toast("OCR failed. Go back and adjust zones, or try better lighting.", "error", 5000);
    setTitle("Scan failed — adjust zones and retry", false);
  }
  if (window.lucide) lucide.createIcons();
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

// ─── Excel Lookup ─────────────────────────────────────────────────────────────
async function lookupPO() {
  const poNumber = document.getElementById("val-po").value.trim();
  if (!poNumber) { toast("Enter or scan a PO number first", "warn"); return; }

  const rc = document.getElementById("result-card");
  rc.style.display = "block";
  document.getElementById("result-body").innerHTML = `<div style="padding:20px;text-align:center;color:var(--text-2);">Looking up ${poNumber}...</div>`;
  document.getElementById("result-badge").textContent = "Searching";
  document.getElementById("result-badge").className = "badge badge-grey";
  document.getElementById("result-actions").innerHTML = "";

  try {
    const data = await fetch(`/api/po/${encodeURIComponent(poNumber)}`).then(r => r.json());
    data.found ? showMatchResult(data.row, data.sheet) : showNoMatchResult(poNumber);
  } catch(e) {
    document.getElementById("result-badge").textContent = "Server Error";
    document.getElementById("result-badge").className = "badge badge-red";
    document.getElementById("result-body").innerHTML = `<div style="padding:16px;color:var(--red);font-size:13px;">Could not connect to server. Is node server.js running?</div>`;
  }
}

function showMatchResult(row, sheet) {
  const scanned = document.getElementById("val-supplier").value.trim();
  const xlSup   = row["SUPPLIER'S NAME"] || "";
  const match   = normalize(xlSup).includes(normalize(scanned)) || normalize(scanned).includes(normalize(xlSup));
  const mismatch = scanned && !match;

  document.getElementById("result-badge").textContent = mismatch ? "Discrepancy" : "Match";
  document.getElementById("result-badge").className   = `badge ${mismatch ? "badge-red" : "badge-green"}`;

  const rows = [
    { f:"PO Number",    s:document.getElementById("val-po").value, x:row["PO NUMBER"]||"—", m:true },
    { f:"Supplier",     s:scanned, x:xlSup, m:match },
    { f:"Sheet",        s:"—", x:sheet, m:true },
    { f:"Status",       s:"—", x:row["PURCHASE ORDER STATUS"]||"—", m:true },
    { f:"Total Amount", s:"—", x:row["TOTAL AMOUNT"]||"—", m:true },
    { f:"Dept",         s:"—", x:row["REQUESTING DEPT."]||"—", m:true },
  ];

  let html = `<div class="table-wrap" style="border-radius:0;border:none;"><table>
    <thead><tr><th>Field</th><th>Scanned</th><th>In Excel</th><th></th></tr></thead><tbody>`;
  rows.forEach(r => {
    const icon = r.s==="—" ? "—" : (r.m ? "✓" : "✗");
    const cls  = r.s==="—" ? "" : (r.m ? "diff-row-match" : "diff-row-mismatch");
    html += `<tr>
      <td class="mono" style="font-size:10px;text-transform:uppercase;color:var(--text-2);">${r.f}</td>
      <td style="font-family:var(--font-mono);font-size:11px;">${r.s||"—"}</td>
      <td style="font-family:var(--font-mono);font-size:11px;">${r.x||"—"}</td>
      <td class="${cls}" style="font-family:var(--font-mono);font-size:12px;">${icon}</td>
    </tr>`;
  });
  html += `</tbody></table></div>`;
  html += mismatch
    ? `<div style="padding:10px 14px;background:#3a1a1a;border-top:1px solid #5a2d2d;font-size:12px;color:var(--red);">Supplier name mismatch — verify before proceeding.</div>`
    : `<div style="padding:10px 14px;background:#1a3a1e;border-top:1px solid #2d5a32;font-size:12px;color:var(--green);">PO found — sheet: ${sheet}</div>`;

  document.getElementById("result-body").innerHTML = html;
  document.getElementById("result-actions").innerHTML = `
    <button class="btn btn-secondary" style="font-size:11px;display:flex;align-items:center;gap:5px;" onclick="copyToClipboard()">
      <i data-lucide="clipboard" style="width:12px;height:12px;"></i> Copy
    </button>`;
  if (window.lucide) lucide.createIcons();
  toast(mismatch ? "Discrepancy found!" : "PO matched.", mismatch ? "warn" : "success");
}

function showNoMatchResult(poNumber) {
  document.getElementById("result-badge").textContent = "Not Found";
  document.getElementById("result-badge").className = "badge badge-yellow";
  document.getElementById("result-body").innerHTML = `
    <div style="padding:18px 14px;">
      <div style="font-size:13px;font-weight:600;margin-bottom:5px;">PO not found in monitoring file</div>
      <div style="font-size:12px;color:var(--text-2);"><strong>${poNumber}</strong> was not found in any monthly sheet.</div>
    </div>`;
  document.getElementById("result-actions").innerHTML = `
    <button class="btn btn-primary" onclick="openAddModal()" style="display:flex;align-items:center;gap:6px;font-size:12px;">
      <i data-lucide="plus-circle" style="width:13px;height:13px;"></i> Add New Row
    </button>`;
  if (window.lucide) lucide.createIcons();
  toast(`PO "${poNumber}" not found.`, "warn");
}

// ─── Add Row Modal ────────────────────────────────────────────────────────────
const COLUMNS = [
  "PR DATE","PR DATE RECEIVED","PR NO.","REQUESTING DEPT.",
  "PO DATE","PO NUMBER","END USER/S","SUPPLIER'S NAME",
  "ITEM CODE","ITEM DESCRIPTION","SPECIFICATIONS","QTY","UoM",
  "UNIT PRICE","AMOUNT","TOTAL AMOUNT","PR REQUIRED DATE",
  "DATE DELIVERED","REMARKS","PURCHASE ORDER STATUS",
  "ITEMS/SERVICES","ORIGINAL PRICE","TOTAL COST SAVINGS"
];

function openAddModal() {
  const prefill = {
    "PO NUMBER":        document.getElementById("val-po").value,
    "SUPPLIER'S NAME":  document.getElementById("val-supplier").value,
    "REQUESTING DEPT.": document.getElementById("val-dept").value,
    "PO DATE":          document.getElementById("val-po-date").value,
    "ITEMS/SERVICES":   document.getElementById("val-items").value,
  };
  const container = document.getElementById("add-form-fields");
  container.innerHTML = "";
  COLUMNS.forEach(col => {
    const div = document.createElement("div");
    div.innerHTML = `<div class="form-label">${col}</div>
      <input class="form-input" type="text" id="addrow-${sanitizeId(col)}"
        value="${escapeHTML(prefill[col]||"")}" placeholder="${col}" />`;
    container.appendChild(div);
  });
  document.getElementById("add-modal").classList.add("open");
}

async function saveNewRow() {
  const month = document.getElementById("val-month").value;
  const rowData = {};
  COLUMNS.forEach(col => {
    const el = document.getElementById("addrow-"+sanitizeId(col));
    if (el) rowData[col] = el.value;
  });
  try {
    const data = await fetch(`/api/sheet/${month}/add`, {
      method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(rowData)
    }).then(r => r.json());
    if (data.success) { toast("Row added to "+month+" sheet!", "success"); closeModal("add-modal"); document.getElementById("result-card").style.display="none"; }
    else toast("Failed to save row","error");
  } catch(e) { toast("Server error","error"); }
}

// ─── Utilities ────────────────────────────────────────────────────────────────
function setTitle(text, live) {
  document.getElementById("panel-title-text").textContent = text;
  document.getElementById("rec-dot").classList.toggle("live", live);
}

function setHeaderControls(html) {
  document.getElementById("header-controls").innerHTML = html;
  if (window.lucide) lucide.createIcons();
}

function setToolbar(html) {
  document.getElementById("viewport-toolbar").innerHTML = html;
  if (window.lucide) lucide.createIcons();
}

function showBrackets(show) {
  ["brackets-tl","brackets-tr","brackets-bl","brackets-br"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = show ? "block" : "none";
  });
}

function setViewportInteractive(interactive) {
  viewportWrap.classList.toggle("calibrate-active",   interactive);
  viewportWrap.classList.toggle("calibrate-inactive", !interactive);
}

function showZones(show) {
  viewportWrap.querySelectorAll(".zone-box").forEach(el => {
    el.style.display = show ? "block" : "none";
  });
}

function showCalibrateBar()  { document.getElementById("calibrate-bar").classList.add("visible"); }
function hideCalibrateBar()  { document.getElementById("calibrate-bar").classList.remove("visible"); }

function clearFields() {
  ["val-po","val-supplier","val-items","val-dept"].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = "";
  });
  ["po-conf","supplier-conf"].forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.textContent = ""; el.className = "badge badge-grey"; }
  });
  document.getElementById("result-card").style.display = "none";
}

function normalize(str) { return String(str||"").toLowerCase().replace(/[^a-z0-9]/g,""); }
function sanitizeId(str){ return str.replace(/[^a-zA-Z0-9]/g,"_"); }
function escapeHTML(str){ return String(str).replace(/"/g,"&quot;"); }
function closeModal(id) { document.getElementById(id).classList.remove("open"); }
function copyToClipboard() {
  const text = `PO: ${document.getElementById("val-po").value}\nSupplier: ${document.getElementById("val-supplier").value}\nItems: ${document.getElementById("val-items").value}`;
  navigator.clipboard.writeText(text).then(() => toast("Copied","success"));
}

document.querySelectorAll(".modal-overlay").forEach(o => {
  o.addEventListener("click", e => { if (e.target===o) o.classList.remove("open"); });
});

// ─── Init ─────────────────────────────────────────────────────────────────────
const MONTH_NAMES = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
document.getElementById("val-month").value = MONTH_NAMES[new Date().getMonth()];
goToStep(1);
initTesseract(); // pre-load in background
