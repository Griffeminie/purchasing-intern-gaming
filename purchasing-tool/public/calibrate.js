/**
 * calibrate.js
 * Full field calibration with perspective correction via OpenCV.js.
 *
 * Key fix: zones are drawn over a <canvas> that always fills the full
 * container width, so percentage positions are always accurate regardless
 * of the original image dimensions.
 */

"use strict";

const ZONES_KEY = "purchasing-tool-zones";

// ─── Field group definitions ─────────────────────────────────────────────────
const FIELD_GROUPS = [
  {
    id: "header", label: "Header Info", color: "#388bfd",
    fields: [
      { key:"po_number",     label:"PO Number",              excelCol:"PO NUMBER",            default:{ top:0.175, left:0.565, width:0.410, height:0.038 } },
      { key:"po_date",       label:"PO Date",                excelCol:"PO DATE",              default:{ top:0.213, left:0.565, width:0.210, height:0.033 } },
      { key:"supplier_name", label:"Supplier Name",          excelCol:"SUPPLIER'S NAME",      default:{ top:0.175, left:0.040, width:0.470, height:0.038 } },
      { key:"supplier_addr", label:"Supplier Address",       excelCol:"REMARKS",              default:{ top:0.213, left:0.040, width:0.470, height:0.040 } },
    ]
  },
  {
    id: "reference", label: "Reference Numbers", color: "#a5a0f5",
    fields: [
      { key:"ro_no",         label:"RO No. / PR No.",        excelCol:"PR NO.",               default:{ top:0.246, left:0.565, width:0.210, height:0.033 } },
      { key:"job_order",     label:"Job Order No.",          excelCol:"END USER/S",           default:{ top:0.246, left:0.775, width:0.200, height:0.033 } },
      { key:"cost_center",   label:"Cost Center",            excelCol:"ITEM CODE",            default:{ top:0.279, left:0.565, width:0.210, height:0.040 } },
      { key:"account_no",    label:"Account No.",            excelCol:"SPECIFICATIONS",       default:{ top:0.279, left:0.775, width:0.200, height:0.040 } },
    ]
  },
  {
    id: "delivery", label: "Delivery Info", color: "#3fb950",
    fields: [
      { key:"deliver_to",    label:"Deliver To",             excelCol:"REQUESTING DEPT.",     default:{ top:0.323, left:0.040, width:0.380, height:0.060 } },
      { key:"deliver_date",  label:"Deliver Not Later Than", excelCol:"PR REQUIRED DATE",     default:{ top:0.323, left:0.420, width:0.220, height:0.040 } },
      { key:"date_delivered",label:"Date Delivered",         excelCol:"DATE DELIVERED",       default:{ top:0.363, left:0.420, width:0.220, height:0.033 } },
      { key:"invoice_no",    label:"Invoice No.",            excelCol:"PR DATE RECEIVED",     default:{ top:0.323, left:0.565, width:0.410, height:0.040 } },
    ]
  },
  {
    id: "items", label: "Items Table", color: "#d29922",
    fields: [
      { key:"item_no",       label:"Item No.",               excelCol:"ITEM CODE",            default:{ top:0.430, left:0.040, width:0.055, height:0.400 } },
      { key:"qty",           label:"Quantity",               excelCol:"QTY",                  default:{ top:0.430, left:0.095, width:0.075, height:0.400 } },
      { key:"unit",          label:"Unit (UoM)",             excelCol:"UoM",                  default:{ top:0.430, left:0.170, width:0.060, height:0.400 } },
      { key:"description",   label:"Item Description",       excelCol:"ITEM DESCRIPTION",     default:{ top:0.430, left:0.230, width:0.420, height:0.400 } },
      { key:"unit_price",    label:"Unit Price",             excelCol:"UNIT PRICE",           default:{ top:0.430, left:0.650, width:0.155, height:0.400 } },
      { key:"amount",        label:"Amount",                 excelCol:"AMOUNT",               default:{ top:0.430, left:0.805, width:0.165, height:0.400 } },
    ]
  },
  {
    id: "totals", label: "Totals & Terms", color: "#f0883e",
    fields: [
      { key:"total_amount",  label:"Total Amount",           excelCol:"TOTAL AMOUNT",         default:{ top:0.852, left:0.805, width:0.165, height:0.038 } },
      { key:"original_price",label:"Original Price",         excelCol:"ORIGINAL PRICE",       default:{ top:0.815, left:0.650, width:0.155, height:0.033 } },
      { key:"cost_savings",  label:"Total Cost Savings",     excelCol:"TOTAL COST SAVINGS",   default:{ top:0.833, left:0.805, width:0.165, height:0.033 } },
      { key:"po_status",     label:"PO Status",              excelCol:"PURCHASE ORDER STATUS",default:{ top:0.888, left:0.280, width:0.460, height:0.033 } },
      { key:"items_services",label:"Items / Services",       excelCol:"ITEMS/SERVICES",       default:{ top:0.930, left:0.040, width:0.450, height:0.033 } },
    ]
  },
  {
    id: "dates", label: "PR Dates (manual — above form)", color: "#f85149",
    fields: [
      { key:"pr_date",       label:"PR Date",                excelCol:"PR DATE",              default:{ top:0.010, left:0.040, width:0.280, height:0.025 } },
      { key:"pr_received",   label:"PR Date Received",       excelCol:"PR DATE RECEIVED",     default:{ top:0.010, left:0.340, width:0.280, height:0.025 } },
    ]
  },
];

// Flat field map
const FIELDS = {};
FIELD_GROUPS.forEach(g => g.fields.forEach(f => { FIELDS[f.key] = { ...f, groupColor: g.color }; }));

// ─── State ───────────────────────────────────────────────────────────────────
let zones        = loadZones();
let selectedKey  = null;
let zonesVisible = true;
let dragState    = null;
let cvReady      = false;

// Canvas refs
let displayCanvas, displayCtx, workCanvas;
// Original image data stored for re-draw
let originalImageData = null; // ImageData of the raw loaded image
let currentImageDataUrl = null; // data URL currently shown on canvas

const values = {};
Object.keys(FIELDS).forEach(k => values[k] = "");

// ─── OpenCV callbacks ─────────────────────────────────────────────────────────
function onOpenCvReady() {
  cvReady = true;
  console.log("OpenCV.js ready");
  // Update rectify button if image already loaded
  const btn = document.getElementById("rectify-btn");
  if (btn) btn.disabled = false;
}

function onOpenCvFailed() {
  console.warn("OpenCV.js failed to load — rectification unavailable");
  const bar = document.getElementById("rectify-bar");
  if (bar) {
    bar.innerHTML = `<i data-lucide="alert-triangle" style="width:14px;height:14px;"></i>
      <span style="flex:1;">Auto-straighten unavailable (OpenCV failed to load). You can still calibrate manually.</span>`;
    bar.classList.add("visible");
    if (window.lucide) lucide.createIcons();
  }
}

// ─── Load / Save ─────────────────────────────────────────────────────────────
function loadZones() {
  try { const s = localStorage.getItem(ZONES_KEY); if (s) return JSON.parse(s); } catch(e) {}
  const d = {};
  Object.values(FIELDS).forEach(f => { d[f.key] = { ...f.default }; });
  return d;
}

function saveAllZones() {
  localStorage.setItem(ZONES_KEY, JSON.stringify(zones));
  toast("All zones saved! Scanner will use these positions.", "success");
}

function resetZones() {
  if (!confirm("Reset all zones to default positions?")) return;
  const d = {};
  Object.values(FIELDS).forEach(f => { d[f.key] = { ...f.default }; });
  zones = d;
  rebuildBoxes();
  if (selectedKey) fillCoordEditor(selectedKey);
  toast("Zones reset to defaults.", "info");
}

// ─── Image loading ────────────────────────────────────────────────────────────
function loadRefImage(event) {
  const file = event.target.files[0];
  if (!file) return;
  toast("Loading image...", "info", 1500);
  const reader = new FileReader();
  reader.onload = e => {
    const img = document.getElementById("ref-image");
    img.src = e.target.result;
    currentImageDataUrl = e.target.result;
  };
  reader.readAsDataURL(file);
}

function onImageLoad() {
  const img = document.getElementById("ref-image");
  displayCanvas = document.getElementById("display-canvas");
  workCanvas    = document.getElementById("work-canvas");
  displayCtx    = displayCanvas.getContext("2d");

  // Draw image to display canvas at full resolution
  displayCanvas.width  = img.naturalWidth;
  displayCanvas.height = img.naturalHeight;
  displayCtx.drawImage(img, 0, 0);

  // Store original for reset
  originalImageData = displayCtx.getImageData(0, 0, displayCanvas.width, displayCanvas.height);

  document.getElementById("img-placeholder").style.display = "none";
  document.getElementById("rectify-bar").classList.add("visible");
  document.getElementById("canvas-hint").textContent = "Click a field → drag to reposition · pull corners to resize";

  if (window.lucide) lucide.createIcons();
  rebuildBoxes();
}

function useOriginal() {
  document.getElementById("rectify-bar").classList.remove("visible");
  rebuildBoxes();
}

// ─── Perspective Correction ───────────────────────────────────────────────────
async function rectifyDocument() {
  if (!cvReady) {
    toast("OpenCV still loading, please wait a moment...", "warn");
    return;
  }

  const overlay = document.getElementById("opencv-overlay");
  const msgEl   = document.getElementById("opencv-msg");
  overlay.classList.add("visible");

  try {
    msgEl.textContent = "Detecting document edges...";
    await sleep(50); // let UI update

    // Read from current display canvas
    const src = cv.imread(displayCanvas);
    const gray = new cv.Mat();
    const blurred = new cv.Mat();
    const edges = new cv.Mat();

    // Convert to grayscale and find edges
    cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY);
    cv.GaussianBlur(gray, blurred, new cv.Size(5, 5), 0);
    cv.Canny(blurred, edges, 50, 150);

    msgEl.textContent = "Finding document contour...";
    await sleep(50);

    // Dilate edges to close gaps
    const kernel = cv.Mat.ones(3, 3, cv.CV_8U);
    const dilated = new cv.Mat();
    cv.dilate(edges, dilated, kernel);

    // Find contours
    const contours  = new cv.MatVector();
    const hierarchy = new cv.Mat();
    cv.findContours(dilated, contours, hierarchy, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE);

    // Find the largest contour that looks like a quadrilateral
    let bestContour = null;
    let bestArea    = 0;
    const imgArea   = src.rows * src.cols;

    for (let i = 0; i < contours.size(); i++) {
      const contour = contours.get(i);
      const area    = cv.contourArea(contour);
      if (area < imgArea * 0.10) continue; // too small

      // Approximate polygon
      const peri    = cv.arcLength(contour, true);
      const approx  = new cv.Mat();
      cv.approxPolyDP(contour, approx, 0.02 * peri, true);

      if (approx.rows === 4 && area > bestArea) {
        bestArea    = area;
        bestContour = approx;
      } else {
        approx.delete();
      }
    }

    if (!bestContour) {
      // No quad found — try a looser search: just use bounding rect of largest contour
      msgEl.textContent = "Using bounding box fallback...";
      await sleep(50);

      let biggestContour = null;
      let biggestArea = 0;
      for (let i = 0; i < contours.size(); i++) {
        const c = contours.get(i);
        const a = cv.contourArea(c);
        if (a > biggestArea) { biggestArea = a; biggestContour = c; }
      }

      if (biggestContour) {
        const rect = cv.boundingRect(biggestContour);
        // Crop to that rect instead of full warp
        const cropped = src.roi(rect);
        const dstCanvas = document.createElement("canvas");
        dstCanvas.width  = rect.width;
        dstCanvas.height = rect.height;
        cv.imshow(dstCanvas, cropped);
        displayCanvas.width  = dstCanvas.width;
        displayCanvas.height = dstCanvas.height;
        displayCtx.drawImage(dstCanvas, 0, 0);
        cropped.delete();
        toast("Document cropped to bounding area. Adjust zones if needed.", "info", 4000);
      } else {
        toast("Could not detect document outline. Try better lighting or higher contrast.", "warn", 5000);
      }

      cleanup();
      overlay.classList.remove("visible");
      rebuildBoxes();
      return;
    }

    msgEl.textContent = "Applying perspective correction...";
    await sleep(50);

    // Order the 4 corners: top-left, top-right, bottom-right, bottom-left
    const corners = orderCorners(bestContour);

    // Compute output dimensions from the warped rectangle
    const w1 = dist(corners[0], corners[1]);
    const w2 = dist(corners[3], corners[2]);
    const h1 = dist(corners[0], corners[3]);
    const h2 = dist(corners[1], corners[2]);
    const outW = Math.round(Math.max(w1, w2));
    const outH = Math.round(Math.max(h1, h2));

    // Source points (detected corners)
    const srcPts = cv.matFromArray(4, 1, cv.CV_32FC2, [
      corners[0].x, corners[0].y,
      corners[1].x, corners[1].y,
      corners[2].x, corners[2].y,
      corners[3].x, corners[3].y,
    ]);

    // Destination points (flat rectangle)
    const dstPts = cv.matFromArray(4, 1, cv.CV_32FC2, [
      0,    0,
      outW, 0,
      outW, outH,
      0,    outH,
    ]);

    const M   = cv.getPerspectiveTransform(srcPts, dstPts);
    const dst = new cv.Mat();
    const dsize = new cv.Size(outW, outH);
    cv.warpPerspective(src, dst, M, dsize, cv.INTER_LINEAR, cv.BORDER_CONSTANT, new cv.Scalar());

    // Draw result to display canvas
    displayCanvas.width  = outW;
    displayCanvas.height = outH;
    cv.imshow(displayCanvas, dst);

    // Update stored image data URL from corrected canvas
    currentImageDataUrl = displayCanvas.toDataURL("image/png");

    // Cleanup
    dst.delete(); M.delete(); srcPts.delete(); dstPts.delete(); bestContour.delete();
    cleanup();

    toast("Document straightened! Zones should now align correctly.", "success");

  } catch(err) {
    console.error("Rectification error:", err);
    toast("Rectification failed: " + err.message, "error", 5000);
  }

  overlay.classList.remove("visible");
  document.getElementById("rectify-bar").classList.remove("visible");
  rebuildBoxes();

  function cleanup() {
    try { src.delete(); gray.delete(); blurred.delete(); edges.delete(); kernel.delete(); dilated.delete(); contours.delete(); hierarchy.delete(); } catch(e) {}
  }
}

// Order 4 corners: top-left, top-right, bottom-right, bottom-left
function orderCorners(mat) {
  const pts = [];
  for (let i = 0; i < 4; i++) {
    pts.push({ x: mat.data32S[i * 2], y: mat.data32S[i * 2 + 1] });
  }
  // Sort by sum (top-left = smallest sum, bottom-right = largest)
  pts.sort((a, b) => (a.x + a.y) - (b.x + b.y));
  const tl = pts[0], br = pts[3];
  // Of remaining two, smaller x-y diff = top-right
  const mid = [pts[1], pts[2]];
  mid.sort((a, b) => (a.x - a.y) - (b.x - b.y));
  const tr = mid[0], bl = mid[1];
  return [tl, tr, br, bl];
}

function dist(a, b) {
  return Math.sqrt((b.x - a.x) ** 2 + (b.y - a.y) ** 2);
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ─── Render zone boxes ────────────────────────────────────────────────────────
function getWrap() { return document.getElementById("canvas-wrap"); }

function pct(v) { return (v * 100).toFixed(3) + "%"; }

function rebuildBoxes() {
  getWrap().querySelectorAll(".zone-box").forEach(el => el.remove());
  if (!displayCanvas || displayCanvas.width === 0) return;

  Object.entries(FIELDS).forEach(([key, field]) => {
    const z = zones[key];
    if (!z) return;

    const box = document.createElement("div");
    box.className = "zone-box" + (selectedKey === key ? " selected" : "");
    box.dataset.key = key;
    box.style.borderColor = field.groupColor;
    box.style.background  = hexToRgba(field.groupColor, 0.08);
    box.style.display     = zonesVisible ? "block" : "none";
    positionBox(box, z);

    const lbl = document.createElement("div");
    lbl.className = "zone-label";
    lbl.textContent = field.label;
    box.appendChild(lbl);

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
    box.addEventListener("click",      () => selectField(key));

    getWrap().appendChild(box);
  });
}

function positionBox(box, z) {
  box.style.left   = pct(z.left);
  box.style.top    = pct(z.top);
  box.style.width  = pct(z.width);
  box.style.height = pct(z.height);
}

function updateBox(key) {
  const box = getWrap().querySelector(`.zone-box[data-key="${key}"]`);
  if (box) positionBox(box, zones[key]);
}

function hexToRgba(hex, a) {
  const r=parseInt(hex.slice(1,3),16), g=parseInt(hex.slice(3,5),16), b=parseInt(hex.slice(5,7),16);
  return `rgba(${r},${g},${b},${a})`;
}

// ─── Field selection ──────────────────────────────────────────────────────────
function selectField(key) {
  selectedKey = key;

  getWrap().querySelectorAll(".zone-box").forEach(el => {
    el.classList.toggle("selected", el.dataset.key === key);
  });
  document.querySelectorAll(".zone-row").forEach(el => {
    el.classList.toggle("active", el.dataset.key === key);
  });
  document.querySelectorAll(".zone-value-wrap").forEach(el => {
    el.style.display = el.dataset.key === key ? "block" : "none";
  });
  document.querySelectorAll(".coord-editor").forEach(el => {
    el.style.display = el.dataset.key === key ? "grid" : "none";
  });
  fillCoordEditor(key);
  const row = document.querySelector(`.zone-row[data-key="${key}"]`);
  if (row) row.scrollIntoView({ behavior:"smooth", block:"nearest" });
}

// ─── Sidebar ─────────────────────────────────────────────────────────────────
function buildSidebar() {
  const container = document.getElementById("zone-groups");
  container.innerHTML = "";

  FIELD_GROUPS.forEach(group => {
    const groupDiv = document.createElement("div");
    groupDiv.className = "zone-group";

    const header = document.createElement("div");
    header.className = "zone-group-header";
    header.innerHTML = `
      <div class="group-dot" style="background:${group.color};"></div>
      <span style="flex:1;">${group.label}</span>
      <i data-lucide="chevron-down" style="width:13px;height:13px;opacity:0.5;"></i>`;
    groupDiv.appendChild(header);

    const body = document.createElement("div");
    body.className = "zone-group-body";

    group.fields.forEach(field => {
      const row = document.createElement("div");
      row.className = "zone-row";
      row.dataset.key = field.key;
      row.innerHTML = `
        <div class="zone-row-dot" style="background:${group.color};"></div>
        <span class="zone-row-label">${field.label}</span>
        <span class="zone-row-shortcut">${field.excelCol.substring(0,8)}</span>`;
      row.addEventListener("click", () => selectField(field.key));
      body.appendChild(row);

      const valWrap = document.createElement("div");
      valWrap.className = "zone-value-wrap";
      valWrap.dataset.key = field.key;
      valWrap.style.display = "none";
      valWrap.innerHTML = `
        <input class="zone-value-input" type="text" id="val-${field.key}"
          placeholder="Extracted value will appear here..."
          value="${values[field.key] || ""}"
          oninput="values['${field.key}'] = this.value" />
        <div style="font-size:10px;color:var(--text-3);margin-top:3px;font-family:var(--font-mono);">
          Excel column: ${field.excelCol}
        </div>`;
      body.appendChild(valWrap);

      const coordEd = document.createElement("div");
      coordEd.className = "coord-editor";
      coordEd.dataset.key = field.key;
      coordEd.style.display = "none";
      coordEd.innerHTML = `
        <div class="coord-wrap"><div class="coord-lbl">Left %</div>  <input class="coord-inp" type="number" id="ci-left-${field.key}"   min="0" max="99"  step="0.1" oninput="applyCoord('${field.key}')"></div>
        <div class="coord-wrap"><div class="coord-lbl">Top %</div>   <input class="coord-inp" type="number" id="ci-top-${field.key}"    min="0" max="99"  step="0.1" oninput="applyCoord('${field.key}')"></div>
        <div class="coord-wrap"><div class="coord-lbl">Width %</div> <input class="coord-inp" type="number" id="ci-width-${field.key}"  min="1" max="100" step="0.1" oninput="applyCoord('${field.key}')"></div>
        <div class="coord-wrap"><div class="coord-lbl">Height %</div><input class="coord-inp" type="number" id="ci-height-${field.key}" min="1" max="100" step="0.1" oninput="applyCoord('${field.key}')"></div>`;
      body.appendChild(coordEd);
    });

    groupDiv.appendChild(body);
    container.appendChild(groupDiv);
  });

  if (window.lucide) lucide.createIcons();
}

function fillCoordEditor(key) {
  const z = zones[key]; if (!z) return;
  const s = (id, val) => { const el = document.getElementById(id); if (el) el.value = (val*100).toFixed(1); };
  s(`ci-left-${key}`, z.left); s(`ci-top-${key}`, z.top);
  s(`ci-width-${key}`, z.width); s(`ci-height-${key}`, z.height);
}

function applyCoord(key) {
  const g = id => parseFloat(document.getElementById(id)?.value || 0) / 100;
  const left=g(`ci-left-${key}`), top=g(`ci-top-${key}`), width=g(`ci-width-${key}`), height=g(`ci-height-${key}`);
  if ([left,top,width,height].some(isNaN)) return;
  zones[key] = { left:clamp(left,0,0.99), top:clamp(top,0,0.99), width:clamp(width,0.01,1-left), height:clamp(height,0.01,1-top) };
  updateBox(key);
}

// ─── Drag / Resize ────────────────────────────────────────────────────────────
function getWrapRect() { return document.getElementById("canvas-wrap").getBoundingClientRect(); }

function startMove(e, key) {
  if (e.target.classList.contains("rh")) return;
  e.preventDefault(); selectField(key);
  const r = getWrapRect();
  dragState = { type:"move", key, startX:e.clientX, startY:e.clientY, startZone:{...zones[key]}, wrapW:r.width, wrapH:r.height };
}
function startResize(e, key, handle) {
  e.preventDefault(); e.stopPropagation();
  const r = getWrapRect();
  dragState = { type:"resize", key, handle, startX:e.clientX, startY:e.clientY, startZone:{...zones[key]}, wrapW:r.width, wrapH:r.height };
}
function startMoveTouch(e, key) {
  if (e.target.classList.contains("rh")) return;
  e.preventDefault(); selectField(key);
  const t=e.touches[0], r=getWrapRect();
  dragState = { type:"move", key, startX:t.clientX, startY:t.clientY, startZone:{...zones[key]}, wrapW:r.width, wrapH:r.height };
}
function startResizeTouch(e, key, handle) {
  e.preventDefault(); e.stopPropagation();
  const t=e.touches[0], r=getWrapRect();
  dragState = { type:"resize", key, handle, startX:t.clientX, startY:t.clientY, startZone:{...zones[key]}, wrapW:r.width, wrapH:r.height };
}

document.addEventListener("mousemove", e => { if (dragState) handleDrag(e.clientX, e.clientY); });
document.addEventListener("mouseup",   () => { dragState = null; });
document.addEventListener("touchmove", e => { if (!dragState) return; e.preventDefault(); handleDrag(e.touches[0].clientX, e.touches[0].clientY); }, { passive:false });
document.addEventListener("touchend",  () => { dragState = null; });

function handleDrag(cx, cy) {
  if (!dragState) return;
  const dx=(cx-dragState.startX)/dragState.wrapW, dy=(cy-dragState.startY)/dragState.wrapH;
  const sz=dragState.startZone, key=dragState.key;
  let {left,top,width,height}=sz;

  if (dragState.type==="move") {
    left=clamp(sz.left+dx,0,1-sz.width); top=clamp(sz.top+dy,0,1-sz.height);
  } else {
    const h=dragState.handle;
    if (h.includes("e")) width =clamp(sz.width +dx,0.01,1-sz.left);
    if (h.includes("s")) height=clamp(sz.height+dy,0.01,1-sz.top);
    if (h.includes("w")) { const nl=clamp(sz.left+dx,0,sz.left+sz.width-0.01);  width=sz.left+sz.width-nl;   left=nl; }
    if (h.includes("n")) { const nt=clamp(sz.top +dy,0,sz.top +sz.height-0.01); height=sz.top+sz.height-nt; top=nt; }
  }
  zones[key]={left,top,width,height};
  updateBox(key);
  fillCoordEditor(key);
}

function clamp(v,mn,mx){return Math.max(mn,Math.min(mx,v));}

// ─── Toggle zones ─────────────────────────────────────────────────────────────
function toggleZones() {
  zonesVisible = !zonesVisible;
  getWrap().querySelectorAll(".zone-box").forEach(el => { el.style.display = zonesVisible ? "block" : "none"; });
  document.getElementById("toggle-btn").innerHTML =
    `<i data-lucide="${zonesVisible?"eye":"eye-off"}" style="width:12px;height:12px;" id="toggle-icon"></i> ${zonesVisible?"Hide All":"Show All"}`;
  if (window.lucide) lucide.createIcons();
}

// ─── Toast ────────────────────────────────────────────────────────────────────
function toast(msg, type="info", duration=3200) {
  const el=document.createElement("div");
  el.className=`toast toast-${type}`; el.textContent=msg;
  document.getElementById("toasts").appendChild(el);
  setTimeout(()=>el.remove(), duration);
}

// ─── Init ─────────────────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
  buildSidebar();
  try {
    const saved = sessionStorage.getItem("purchasing-tool-last-scan-image");
    if (saved) {
      const img = document.getElementById("ref-image");
      img.src = saved;
      toast("Image loaded from scanner. Try Auto-Straighten first.", "info", 4000);
    }
  } catch(e) { console.warn("Could not load scanner image:", e); }
});
