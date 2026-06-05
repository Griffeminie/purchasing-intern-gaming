/**
 * calibrate.js
 * Zone calibration tool — drag and resize OCR scan zones over a reference PO image.
 * Zones saved to localStorage and picked up automatically by scanner.js.
 */

"use strict";

const ZONES_KEY    = "purchasing-tool-zones";
const ZONE_STORAGE_KEY = ZONES_KEY; // same key scanner.js reads

const DEFAULT_ZONES = {
  poNumber: { top: 0.03, left: 0.76, width: 0.22, height: 0.07 },
  supplier: { top: 0.03, left: 0.02, width: 0.30, height: 0.07 },
  items:    { top: 0.22, left: 0.02, width: 0.96, height: 0.45 },
};

const ZONE_META = {
  poNumber: { label: "PO Number",    color: "#388bfd", dotClass: "zone-color-po" },
  supplier: { label: "Supplier Name",color: "#3fb950", dotClass: "zone-color-supplier" },
  items:    { label: "Items & Amounts", color: "#d29922", dotClass: "zone-color-items" },
};

// ── State ────────────────────────────────────────────────────────────────────
let zones = loadZones();          // { poNumber, supplier, items } — 0–1 fractions
let selectedZone = null;          // key of currently selected zone
let imageLoaded = false;
let zonesVisible = true;

// drag/resize state
let dragState = null; // { type: "move"|"resize", zone, handle, startX, startY, startZone }

// ── Load / Save ──────────────────────────────────────────────────────────────
function loadZones() {
  try {
    const s = localStorage.getItem(ZONES_KEY);
    if (s) return JSON.parse(s);
  } catch(e) {}
  return JSON.parse(JSON.stringify(DEFAULT_ZONES)); // deep copy
}

function saveZones() {
  localStorage.setItem(ZONES_KEY, JSON.stringify(zones));
  toast("Zones saved! Scanner will use these positions.", "success");
}

function resetZones() {
  if (!confirm("Reset all zones to default positions?")) return;
  zones = JSON.parse(JSON.stringify(DEFAULT_ZONES));
  renderAllZones();
  updateSidebar();
  if (selectedZone) fillEditor(selectedZone);
  toast("Zones reset to defaults.", "info");
}

// ── Image loading ─────────────────────────────────────────────────────────────
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
  document.getElementById("canvas-hint").textContent =
    "Drag zones to reposition · Drag edges/corners to resize";
  renderAllZones();
  updateSidebar();
}

// ── Render ────────────────────────────────────────────────────────────────────
function getWrap() { return document.getElementById("canvas-wrap"); }

function pct(v) { return (v * 100).toFixed(2) + "%"; }

function renderAllZones() {
  // Remove existing zone boxes
  getWrap().querySelectorAll(".zone-box").forEach(el => el.remove());
  Object.keys(zones).forEach(key => renderZone(key));
}

function renderZone(key) {
  const wrap = getWrap();
  const z = zones[key];
  const meta = ZONE_META[key];

  const box = document.createElement("div");
  box.className = "zone-box" + (selectedZone === key ? " selected" : "");
  box.dataset.zone = key;
  box.style.left    = pct(z.left);
  box.style.top     = pct(z.top);
  box.style.width   = pct(z.width);
  box.style.height  = pct(z.height);
  box.style.display = zonesVisible ? "block" : "none";

  // Label
  const label = document.createElement("div");
  label.className = "zone-box-label";
  label.textContent = meta.label;
  box.appendChild(label);

  // Resize handles
  ["nw","n","ne","e","se","s","sw","w"].forEach(dir => {
    const h = document.createElement("div");
    h.className = `resize-handle ${dir}`;
    h.dataset.handle = dir;
    h.addEventListener("mousedown",  e => startResize(e, key, dir));
    h.addEventListener("touchstart", e => startResizeTouch(e, key, dir), { passive: false });
    box.appendChild(h);
  });

  // Drag to move
  box.addEventListener("mousedown",  e => startMove(e, key));
  box.addEventListener("touchstart", e => startMoveTouch(e, key), { passive: false });

  // Select on click
  box.addEventListener("click", () => selectZone(key));

  wrap.appendChild(box);
}

function updateZoneBox(key) {
  const box = getWrap().querySelector(`.zone-box[data-zone="${key}"]`);
  if (!box) return;
  const z = zones[key];
  box.style.left   = pct(z.left);
  box.style.top    = pct(z.top);
  box.style.width  = pct(z.width);
  box.style.height = pct(z.height);
}

// ── Selection ─────────────────────────────────────────────────────────────────
function selectZone(key) {
  selectedZone = key;
  getWrap().querySelectorAll(".zone-box").forEach(el => {
    el.classList.toggle("selected", el.dataset.zone === key);
  });
  document.querySelectorAll(".zone-list-item").forEach(el => {
    el.classList.toggle("active", el.dataset.zone === key);
  });
  fillEditor(key);
  document.getElementById("zone-editor").style.display = "block";
}

// ── Sidebar zone list ─────────────────────────────────────────────────────────
function updateSidebar() {
  const list = document.getElementById("zone-list");
  list.innerHTML = "";

  Object.keys(ZONE_META).forEach(key => {
    const meta = ZONE_META[key];
    const z = zones[key];

    const item = document.createElement("div");
    item.className = "zone-list-item" + (selectedZone === key ? " active" : "");
    item.dataset.zone = key;
    item.innerHTML = `
      <div class="zone-color-dot ${meta.dotClass}"></div>
      <div style="flex:1">
        <div class="zone-list-label">${meta.label}</div>
        <div class="zone-list-coords" id="coords-${key}">
          L: ${(z.left*100).toFixed(1)}%  T: ${(z.top*100).toFixed(1)}%<br>
          W: ${(z.width*100).toFixed(1)}%  H: ${(z.height*100).toFixed(1)}%
        </div>
      </div>
      <i data-lucide="chevron-right" style="width:14px;height:14px;color:var(--text-3);"></i>
    `;
    item.addEventListener("click", () => selectZone(key));
    list.appendChild(item);
  });

  if (window.lucide) lucide.createIcons();
}

function updateCoordsDisplay(key) {
  const z = zones[key];
  const el = document.getElementById("coords-" + key);
  if (el) {
    el.innerHTML =
      `L: ${(z.left*100).toFixed(1)}%  T: ${(z.top*100).toFixed(1)}%<br>` +
      `W: ${(z.width*100).toFixed(1)}%  H: ${(z.height*100).toFixed(1)}%`;
  }
}

// ── Coordinate editor inputs ──────────────────────────────────────────────────
function fillEditor(key) {
  const z = zones[key];
  const meta = ZONE_META[key];
  document.getElementById("editor-title").textContent = meta.label;
  document.getElementById("editor-badge").textContent = key;
  document.getElementById("inp-left").value   = (z.left   * 100).toFixed(1);
  document.getElementById("inp-top").value    = (z.top    * 100).toFixed(1);
  document.getElementById("inp-width").value  = (z.width  * 100).toFixed(1);
  document.getElementById("inp-height").value = (z.height * 100).toFixed(1);
}

function applyCoordInput() {
  if (!selectedZone) return;
  const left   = parseFloat(document.getElementById("inp-left").value)   / 100;
  const top    = parseFloat(document.getElementById("inp-top").value)    / 100;
  const width  = parseFloat(document.getElementById("inp-width").value)  / 100;
  const height = parseFloat(document.getElementById("inp-height").value) / 100;

  if ([left,top,width,height].some(isNaN)) return;

  zones[selectedZone] = {
    left:   Math.max(0, Math.min(0.99, left)),
    top:    Math.max(0, Math.min(0.99, top)),
    width:  Math.max(0.01, Math.min(1 - left, width)),
    height: Math.max(0.01, Math.min(1 - top,  height)),
  };

  updateZoneBox(selectedZone);
  updateCoordsDisplay(selectedZone);
}

// ── Drag to Move (Mouse) ──────────────────────────────────────────────────────
function startMove(e, key) {
  if (e.target.classList.contains("resize-handle")) return;
  e.preventDefault();
  selectZone(key);

  const wrap = getWrap();
  const rect = wrap.getBoundingClientRect();

  dragState = {
    type: "move",
    zone: key,
    startX: e.clientX,
    startY: e.clientY,
    startZone: { ...zones[key] },
    wrapW: rect.width,
    wrapH: rect.height,
  };
}

// ── Drag to Resize (Mouse) ────────────────────────────────────────────────────
function startResize(e, key, handle) {
  e.preventDefault();
  e.stopPropagation();
  selectZone(key);

  const wrap = getWrap();
  const rect = wrap.getBoundingClientRect();

  dragState = {
    type: "resize",
    zone: key,
    handle,
    startX: e.clientX,
    startY: e.clientY,
    startZone: { ...zones[key] },
    wrapW: rect.width,
    wrapH: rect.height,
  };
}

// ── Mouse move / up ───────────────────────────────────────────────────────────
document.addEventListener("mousemove", e => {
  if (!dragState) return;
  handleDrag(e.clientX, e.clientY);
});

document.addEventListener("mouseup", () => { dragState = null; });

// ── Touch support ─────────────────────────────────────────────────────────────
function startMoveTouch(e, key) {
  if (e.target.classList.contains("resize-handle")) return;
  e.preventDefault();
  selectZone(key);
  const t = e.touches[0];
  const wrap = getWrap();
  const rect = wrap.getBoundingClientRect();
  dragState = {
    type: "move",
    zone: key,
    startX: t.clientX,
    startY: t.clientY,
    startZone: { ...zones[key] },
    wrapW: rect.width,
    wrapH: rect.height,
  };
}

function startResizeTouch(e, key, handle) {
  e.preventDefault();
  e.stopPropagation();
  selectZone(key);
  const t = e.touches[0];
  const wrap = getWrap();
  const rect = wrap.getBoundingClientRect();
  dragState = {
    type: "resize",
    zone: key,
    handle,
    startX: t.clientX,
    startY: t.clientY,
    startZone: { ...zones[key] },
    wrapW: rect.width,
    wrapH: rect.height,
  };
}

document.addEventListener("touchmove", e => {
  if (!dragState) return;
  e.preventDefault();
  const t = e.touches[0];
  handleDrag(t.clientX, t.clientY);
}, { passive: false });

document.addEventListener("touchend", () => { dragState = null; });

// ── Core drag handler ─────────────────────────────────────────────────────────
function handleDrag(clientX, clientY) {
  if (!dragState) return;

  const dx = (clientX - dragState.startX) / dragState.wrapW;
  const dy = (clientY - dragState.startY) / dragState.wrapH;
  const sz = dragState.startZone;
  const key = dragState.zone;
  let { left, top, width, height } = sz;

  if (dragState.type === "move") {
    left = clamp(sz.left + dx, 0, 1 - sz.width);
    top  = clamp(sz.top  + dy, 0, 1 - sz.height);
    zones[key] = { left, top, width, height };

  } else {
    // resize — depends on which handle is being dragged
    const h = dragState.handle;

    if (h.includes("e")) {
      width  = clamp(sz.width  + dx, 0.02, 1 - sz.left);
    }
    if (h.includes("s")) {
      height = clamp(sz.height + dy, 0.02, 1 - sz.top);
    }
    if (h.includes("w")) {
      const newLeft  = clamp(sz.left + dx, 0, sz.left + sz.width - 0.02);
      width  = sz.left + sz.width - newLeft;
      left   = newLeft;
    }
    if (h.includes("n")) {
      const newTop   = clamp(sz.top  + dy, 0, sz.top  + sz.height - 0.02);
      height = sz.top + sz.height - newTop;
      top    = newTop;
    }

    zones[key] = { left, top, width, height };
  }

  updateZoneBox(key);
  updateCoordsDisplay(key);
  if (selectedZone === key) fillEditor(key);
}

function clamp(v, min, max) { return Math.max(min, Math.min(max, v)); }

// ── Toggle zone visibility ────────────────────────────────────────────────────
function toggleAllZones() {
  zonesVisible = !zonesVisible;
  getWrap().querySelectorAll(".zone-box").forEach(el => {
    el.style.display = zonesVisible ? "block" : "none";
  });
  const btn = document.getElementById("toggle-btn");
  const icon = btn.querySelector("i");
  icon.setAttribute("data-lucide", zonesVisible ? "eye" : "eye-off");
  btn.querySelector("i").nextSibling
    ? btn.childNodes[btn.childNodes.length - 1].textContent = zonesVisible ? " Hide Zones" : " Show Zones"
    : null;
  btn.innerHTML = `<i data-lucide="${zonesVisible ? 'eye' : 'eye-off'}" style="width:13px;height:13px;"></i> ${zonesVisible ? 'Hide' : 'Show'} Zones`;
  if (window.lucide) lucide.createIcons();
}

// ── Toast ─────────────────────────────────────────────────────────────────────
function toast(msg, type = "info", duration = 3200) {
  const container = document.getElementById("toasts");
  const el = document.createElement("div");
  el.className = `toast toast-${type}`;
  el.textContent = msg;
  container.appendChild(el);
  setTimeout(() => el.remove(), duration);
}

// ── Init ──────────────────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
  updateSidebar();

  // Check if scanner.html passed an image via sessionStorage
  const LAST_IMAGE_KEY = "purchasing-tool-last-scan-image";
  try {
    const saved = sessionStorage.getItem(LAST_IMAGE_KEY);
    if (saved) {
      const img = document.getElementById("ref-image");
      img.src = saved;
      img.style.display = "block";
      document.getElementById("img-placeholder").style.display = "none";
      // Don't remove it — keeps working if user refreshes calibrate page
      toast("Image loaded from scanner. Adjust zones and save.", "info", 4000);
    }
  } catch(e) {
    console.warn("Could not load scanner image:", e);
  }
});
