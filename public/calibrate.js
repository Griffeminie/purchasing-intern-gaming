/**
 * calibrate.js
 * Full field calibration — one draggable zone per Excel column.
 * Zones are color-grouped by section, matching the physical PO layout.
 *
 * Default positions are pre-set based on the standard California Clothing / Lightstar PO template.
 * Adjust by dragging on your actual PO photo.
 */

"use strict";

const ZONES_KEY = "purchasing-tool-zones";

// ─── Field definitions ────────────────────────────────────────────────────────
// Each field has:
//   key       — unique ID, matches Excel column name (sanitized)
//   label     — display name
//   excelCol  — exact Excel column header
//   group     — visual grouping
//   color     — hex color for the zone box
//   default   — { top, left, width, height } as 0-1 fractions of image size
//               Pre-mapped to the PO template in the uploaded sample image

const FIELD_GROUPS = [
  {
    id:    "header",
    label: "Header Info",
    color: "#388bfd",
    fields: [
      { key:"po_number",     label:"PO Number",              excelCol:"PO NUMBER",            default:{ top:0.175, left:0.565, width:0.410, height:0.038 } },
      { key:"po_date",       label:"PO Date",                excelCol:"PO DATE",              default:{ top:0.213, left:0.565, width:0.210, height:0.033 } },
      { key:"supplier_name", label:"Supplier Name",          excelCol:"SUPPLIER'S NAME",      default:{ top:0.175, left:0.040, width:0.470, height:0.038 } },
      { key:"supplier_addr", label:"Supplier Address",       excelCol:"REMARKS",              default:{ top:0.213, left:0.040, width:0.470, height:0.040 } },
    ]
  },
  {
    id:    "reference",
    label: "Reference Numbers",
    color: "#a5a0f5",
    fields: [
      { key:"ro_no",         label:"RO No. / PR No.",        excelCol:"PR NO.",               default:{ top:0.246, left:0.565, width:0.210, height:0.033 } },
      { key:"job_order",     label:"Job Order No.",          excelCol:"END USER/S",           default:{ top:0.246, left:0.775, width:0.200, height:0.033 } },
      { key:"cost_center",   label:"Cost Center",            excelCol:"ITEM CODE",            default:{ top:0.279, left:0.565, width:0.210, height:0.040 } },
      { key:"account_no",    label:"Account No.",            excelCol:"SPECIFICATIONS",       default:{ top:0.279, left:0.775, width:0.200, height:0.040 } },
    ]
  },
  {
    id:    "delivery",
    label: "Delivery Info",
    color: "#3fb950",
    fields: [
      { key:"deliver_to",    label:"Deliver To",             excelCol:"REQUESTING DEPT.",     default:{ top:0.323, left:0.040, width:0.380, height:0.060 } },
      { key:"deliver_date",  label:"Deliver Not Later Than", excelCol:"PR REQUIRED DATE",     default:{ top:0.323, left:0.420, width:0.220, height:0.040 } },
      { key:"date_delivered",label:"Date Delivered",         excelCol:"DATE DELIVERED",       default:{ top:0.363, left:0.420, width:0.220, height:0.033 } },
      { key:"invoice_no",    label:"Invoice No.",            excelCol:"PR DATE RECEIVED",     default:{ top:0.323, left:0.565, width:0.410, height:0.040 } },
    ]
  },
  {
    id:    "items",
    label: "Items Table",
    color: "#d29922",
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
    id:    "totals",
    label: "Totals & Terms",
    color: "#f0883e",
    fields: [
      { key:"total_amount",  label:"Total Amount",           excelCol:"TOTAL AMOUNT",         default:{ top:0.852, left:0.805, width:0.165, height:0.038 } },
      { key:"original_price",label:"Original Price",         excelCol:"ORIGINAL PRICE",       default:{ top:0.815, left:0.650, width:0.155, height:0.033 } },
      { key:"cost_savings",  label:"Total Cost Savings",     excelCol:"TOTAL COST SAVINGS",   default:{ top:0.833, left:0.805, width:0.165, height:0.033 } },
      { key:"po_status",     label:"PO Status",              excelCol:"PURCHASE ORDER STATUS",default:{ top:0.888, left:0.280, width:0.460, height:0.033 } },
      { key:"items_services",label:"Items / Services",       excelCol:"ITEMS/SERVICES",       default:{ top:0.930, left:0.040, width:0.450, height:0.033 } },
    ]
  },
  {
    id:    "dates",
    label: "PR Dates (manual — placed above form)",
    color: "#f85149",
    fields: [
      { key:"pr_date",       label:"PR Date",                excelCol:"PR DATE",              default:{ top:0.010, left:0.040, width:0.280, height:0.025 } },
      { key:"pr_received",   label:"PR Date Received",       excelCol:"PR DATE RECEIVED",     default:{ top:0.010, left:0.340, width:0.280, height:0.025 } },
    ]
  },
];

// ─── Build flat field map ─────────────────────────────────────────────────────
const FIELDS = {}; // key → { ...field, group }
FIELD_GROUPS.forEach(g => {
  g.fields.forEach(f => {
    FIELDS[f.key] = { ...f, groupColor: g.color, groupId: g.id };
  });
});

// ─── State ───────────────────────────────────────────────────────────────────
let zones       = loadZones();   // key → { top, left, width, height }
let selectedKey = null;
let zonesVisible = true;
let imageLoaded  = false;
let dragState    = null;

// Extracted values from OCR / manual entry
const values = {};
Object.keys(FIELDS).forEach(k => values[k] = "");

// ─── Load / Save ──────────────────────────────────────────────────────────────
function loadZones() {
  try {
    const s = localStorage.getItem(ZONES_KEY);
    if (s) return JSON.parse(s);
  } catch(e) {}
  // Build defaults
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
  const reader = new FileReader();
  reader.onload = e => {
    const img = document.getElementById("ref-image");
    img.src = e.target.result;
    img.style.display = "block";
    document.getElementById("img-placeholder").style.display = "none";
  };
  reader.readAsDataURL(file);
}

function onImageLoad() {
  imageLoaded = true;
  document.getElementById("canvas-hint").textContent = "Click a field → drag to reposition · pull corners to resize";
  rebuildBoxes();
}

// ─── Render zone boxes ────────────────────────────────────────────────────────
function getWrap() { return document.getElementById("canvas-wrap"); }

function pct(v) { return (v * 100).toFixed(3) + "%"; }

function rebuildBoxes() {
  getWrap().querySelectorAll(".zone-box").forEach(el => el.remove());
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

    // Label
    const lbl = document.createElement("div");
    lbl.className = "zone-label";
    lbl.textContent = field.label;
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
  const r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
  return `rgba(${r},${g},${b},${a})`;
}

// ─── Field selection ──────────────────────────────────────────────────────────
function selectField(key) {
  const prev = selectedKey;
  selectedKey = key;

  // Update box selected state
  getWrap().querySelectorAll(".zone-box").forEach(el => {
    el.classList.toggle("selected", el.dataset.key === key);
  });

  // Update sidebar rows
  document.querySelectorAll(".zone-row").forEach(el => {
    el.classList.toggle("active", el.dataset.key === key);
  });

  // Show/hide value inputs and coord editors
  document.querySelectorAll(".zone-value-wrap").forEach(el => {
    el.style.display = el.dataset.key === key ? "block" : "none";
  });
  document.querySelectorAll(".coord-editor").forEach(el => {
    el.style.display = el.dataset.key === key ? "grid" : "none";
  });

  fillCoordEditor(key);

  // Scroll zone into view in sidebar
  const row = document.querySelector(`.zone-row[data-key="${key}"]`);
  if (row) row.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

// ─── Sidebar rendering ────────────────────────────────────────────────────────
function buildSidebar() {
  const container = document.getElementById("zone-groups");
  container.innerHTML = "";

  FIELD_GROUPS.forEach(group => {
    const groupDiv = document.createElement("div");
    groupDiv.className = "zone-group";

    // Group header (toggles visibility of group)
    const header = document.createElement("div");
    header.className = "zone-group-header";
    header.innerHTML = `
      <div class="group-dot" style="background:${group.color};"></div>
      <span style="flex:1;">${group.label}</span>
      <i data-lucide="chevron-down" style="width:13px;height:13px;opacity:0.5;"></i>
    `;
    groupDiv.appendChild(header);

    const body = document.createElement("div");
    body.className = "zone-group-body";

    group.fields.forEach(field => {
      // Zone row (clickable)
      const row = document.createElement("div");
      row.className = "zone-row";
      row.dataset.key = field.key;
      row.innerHTML = `
        <div class="zone-row-dot" style="background:${group.color};"></div>
        <span class="zone-row-label">${field.label}</span>
        <span class="zone-row-shortcut">${field.excelCol.substring(0,8)}</span>
      `;
      row.addEventListener("click", () => selectField(field.key));
      body.appendChild(row);

      // Value input (hidden until selected)
      const valWrap = document.createElement("div");
      valWrap.className = "zone-value-wrap";
      valWrap.dataset.key = field.key;
      valWrap.style.display = "none";
      valWrap.innerHTML = `
        <input class="zone-value-input" type="text" id="val-${field.key}"
          placeholder="Extracted value will appear here..."
          value="${values[field.key] || ""}"
          oninput="values['${field.key}'] = this.value"
        />
        <div style="font-size:10px;color:var(--text-3);margin-top:3px;font-family:var(--font-mono);">
          Excel column: ${field.excelCol}
        </div>
      `;
      body.appendChild(valWrap);

      // Coord editor (hidden until selected)
      const coordEd = document.createElement("div");
      coordEd.className = "coord-editor";
      coordEd.dataset.key = field.key;
      coordEd.style.display = "none";
      coordEd.innerHTML = `
        <div class="coord-wrap"><div class="coord-lbl">Left %</div>  <input class="coord-inp" type="number" id="ci-left-${field.key}"   min="0" max="99" step="0.1" oninput="applyCoord('${field.key}')"></div>
        <div class="coord-wrap"><div class="coord-lbl">Top %</div>   <input class="coord-inp" type="number" id="ci-top-${field.key}"    min="0" max="99" step="0.1" oninput="applyCoord('${field.key}')"></div>
        <div class="coord-wrap"><div class="coord-lbl">Width %</div> <input class="coord-inp" type="number" id="ci-width-${field.key}"  min="1" max="100" step="0.1" oninput="applyCoord('${field.key}')"></div>
        <div class="coord-wrap"><div class="coord-lbl">Height %</div><input class="coord-inp" type="number" id="ci-height-${field.key}" min="1" max="100" step="0.1" oninput="applyCoord('${field.key}')"></div>
      `;
      body.appendChild(coordEd);
    });

    groupDiv.appendChild(body);
    container.appendChild(groupDiv);
  });

  if (window.lucide) lucide.createIcons();
}

function fillCoordEditor(key) {
  const z = zones[key];
  if (!z) return;
  const set = (id, val) => { const el = document.getElementById(id); if (el) el.value = (val*100).toFixed(1); };
  set(`ci-left-${key}`,   z.left);
  set(`ci-top-${key}`,    z.top);
  set(`ci-width-${key}`,  z.width);
  set(`ci-height-${key}`, z.height);
}

function applyCoord(key) {
  const get = id => parseFloat(document.getElementById(id)?.value || 0) / 100;
  const left   = get(`ci-left-${key}`);
  const top    = get(`ci-top-${key}`);
  const width  = get(`ci-width-${key}`);
  const height = get(`ci-height-${key}`);
  if ([left,top,width,height].some(isNaN)) return;
  zones[key] = {
    left:   clamp(left,  0, 0.99),
    top:    clamp(top,   0, 0.99),
    width:  clamp(width, 0.01, 1 - left),
    height: clamp(height,0.01, 1 - top),
  };
  updateBox(key);
}

// ─── Drag / Resize ────────────────────────────────────────────────────────────
function getWrapRect() { return document.getElementById("canvas-wrap").getBoundingClientRect(); }

function startMove(e, key) {
  if (e.target.classList.contains("rh")) return;
  e.preventDefault();
  selectField(key);
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
  e.preventDefault();
  selectField(key);
  const t = e.touches[0]; const r = getWrapRect();
  dragState = { type:"move", key, startX:t.clientX, startY:t.clientY, startZone:{...zones[key]}, wrapW:r.width, wrapH:r.height };
}
function startResizeTouch(e, key, handle) {
  e.preventDefault(); e.stopPropagation();
  const t = e.touches[0]; const r = getWrapRect();
  dragState = { type:"resize", key, handle, startX:t.clientX, startY:t.clientY, startZone:{...zones[key]}, wrapW:r.width, wrapH:r.height };
}

document.addEventListener("mousemove", e => { if (dragState) handleDrag(e.clientX, e.clientY); });
document.addEventListener("mouseup",   () => { dragState = null; });
document.addEventListener("touchmove", e => {
  if (!dragState) return; e.preventDefault();
  handleDrag(e.touches[0].clientX, e.touches[0].clientY);
}, { passive: false });
document.addEventListener("touchend", () => { dragState = null; });

function handleDrag(cx, cy) {
  if (!dragState) return;
  const dx = (cx - dragState.startX) / dragState.wrapW;
  const dy = (cy - dragState.startY) / dragState.wrapH;
  const sz = dragState.startZone;
  const key = dragState.key;
  let { left, top, width, height } = sz;

  if (dragState.type === "move") {
    left = clamp(sz.left + dx, 0, 1 - sz.width);
    top  = clamp(sz.top  + dy, 0, 1 - sz.height);
  } else {
    const h = dragState.handle;
    if (h.includes("e")) width  = clamp(sz.width  + dx, 0.01, 1 - sz.left);
    if (h.includes("s")) height = clamp(sz.height + dy, 0.01, 1 - sz.top);
    if (h.includes("w")) {
      const nl = clamp(sz.left + dx, 0, sz.left + sz.width - 0.01);
      width = sz.left + sz.width - nl; left = nl;
    }
    if (h.includes("n")) {
      const nt = clamp(sz.top + dy, 0, sz.top + sz.height - 0.01);
      height = sz.top + sz.height - nt; top = nt;
    }
  }

  zones[key] = { left, top, width, height };
  updateBox(key);
  fillCoordEditor(key);
}

function clamp(v, mn, mx) { return Math.max(mn, Math.min(mx, v)); }

// ─── Toggle all zones ─────────────────────────────────────────────────────────
function toggleZones() {
  zonesVisible = !zonesVisible;
  getWrap().querySelectorAll(".zone-box").forEach(el => {
    el.style.display = zonesVisible ? "block" : "none";
  });
  const icon = document.getElementById("toggle-icon");
  icon.setAttribute("data-lucide", zonesVisible ? "eye" : "eye-off");
  document.getElementById("toggle-btn").innerHTML =
    `<i data-lucide="${zonesVisible ? 'eye' : 'eye-off'}" style="width:12px;height:12px;" id="toggle-icon"></i> ${zonesVisible ? 'Hide All' : 'Show All'}`;
  if (window.lucide) lucide.createIcons();
}

// ─── Toast ────────────────────────────────────────────────────────────────────
function toast(msg, type = "info", duration = 3200) {
  const el = document.createElement("div");
  el.className = `toast toast-${type}`;
  el.textContent = msg;
  document.getElementById("toasts").appendChild(el);
  setTimeout(() => el.remove(), duration);
}

// ─── Init ─────────────────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
  buildSidebar();

  // Load image passed from scanner via sessionStorage
  try {
    const saved = sessionStorage.getItem("purchasing-tool-last-scan-image");
    if (saved) {
      const img = document.getElementById("ref-image");
      img.src = saved;
      img.style.display = "block";
      document.getElementById("img-placeholder").style.display = "none";
      toast("Image loaded from scanner. Adjust zones and save.", "info", 4000);
    }
  } catch(e) { console.warn("Could not load scanner image:", e); }
});
