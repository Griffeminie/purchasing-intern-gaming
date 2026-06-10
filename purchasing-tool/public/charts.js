/**
 * charts.js — Dashboard with localStorage caching
 * Loads instantly from cache, refreshes on demand.
 */
"use strict";

Chart.defaults.color = "#8b949e";
Chart.defaults.borderColor = "#30363d";
Chart.defaults.font.family = "'IBM Plex Mono', monospace";
Chart.defaults.font.size = 11;

const PALETTE = [
  "#1f6feb","#388bfd","#3fb950","#d29922",
  "#f0883e","#a5a0f5","#f85149","#56d364","#e3b341","#79c0ff",
];

let charts = {};

const CACHE_KEY      = "purchasing-tool-dashboard-cache";
const CACHE_META_KEY = "purchasing-tool-dashboard-meta";

// ─── Cache helpers ────────────────────────────────────────────────────────────
function saveCache(data) {
  try {
    localStorage.setItem(CACHE_KEY,      JSON.stringify(data));
    localStorage.setItem(CACHE_META_KEY, JSON.stringify({
      savedAt: Date.now(),
      rawRows: data.rawRows,
      sheets:  data.sheetsDetected,
    }));
  } catch(e) { console.warn("Cache save failed:", e); }
}

function loadCache() {
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch(e) { return null; }
}

function loadCacheMeta() {
  try {
    const raw = localStorage.getItem(CACHE_META_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch(e) { return null; }
}

function fmtAge(ms) {
  const s = Math.floor(ms / 1000);
  if (s < 60)    return `${s}s ago`;
  if (s < 3600)  return `${Math.floor(s / 60)}m ago`;
  if (s < 86400) return `${Math.floor(s / 3600)}h ago`;
  return `${Math.floor(s / 86400)}d ago`;
}

function updateCacheBadge() {
  const meta = loadCacheMeta();
  if (!meta) return;
  const age = Date.now() - meta.savedAt;

  const badge = document.getElementById("cache-badge");
  if (badge) {
    badge.textContent = `Last updated ${fmtAge(age)}`;
    badge.style.display = "inline";
  }

  const note = document.getElementById("data-note");
  if (note) note.textContent = `${meta.rawRows ?? "?"} records · ${(meta.sheets || []).join(", ")}`;
}

// ─── Debug log ────────────────────────────────────────────────────────────────
function log(msg, type = "info") {
  const panel = document.getElementById("debug-log");
  if (!panel) return;
  const ts = new Date().toLocaleTimeString();
  const colors = { info:"#8b949e", ok:"#3fb950", warn:"#d29922", error:"#f85149" };
  const line = document.createElement("div");
  line.style.cssText = `color:${colors[type]||colors.info};font-size:11px;padding:2px 0;border-bottom:1px solid var(--border-2);`;
  line.textContent = `[${ts}] ${msg}`;
  panel.appendChild(line);
  panel.scrollTop = panel.scrollHeight;
  console.log(`[Dashboard ${type}]`, msg);
}

// ─── Toast ────────────────────────────────────────────────────────────────────
function toast(msg, type = "info") {
  const el = document.createElement("div");
  el.className = `toast toast-${type}`;
  el.textContent = msg;
  document.getElementById("toasts").appendChild(el);
  setTimeout(() => el.remove(), 4000);
}

// ─── Format currency ──────────────────────────────────────────────────────────
function fmt(num) {
  if (!num || isNaN(num)) return "₱0";
  return "₱" + Number(num).toLocaleString("en-PH", { minimumFractionDigits:0, maximumFractionDigits:0 });
}

function hideLoading(id) {
  const el = document.getElementById(id);
  if (el) el.style.display = "none";
}

function showEmpty(loadingId, msg) {
  const el = document.getElementById(loadingId);
  if (el) { el.textContent = msg; el.style.display = "flex"; }
}

// ─── Main load ────────────────────────────────────────────────────────────────
async function loadDashboard(forceRefresh = false) {
  const debugPanel = document.getElementById("debug-panel");
  if (debugPanel) debugPanel.style.display = "block";
  const logEl = document.getElementById("debug-log");
  if (logEl) logEl.innerHTML = "";

  // ── Try cache first (unless forced refresh) ───────────────────────────────
  if (!forceRefresh) {
    const cached = loadCache();
    if (cached) {
      log("Loaded from local cache — instant!", "ok");
      renderAll(cached);
      updateCacheBadge();

      const meta = loadCacheMeta();
      if (meta && Date.now() - meta.savedAt > 30 * 60 * 1000) {
        log("Cache is over 30 minutes old — consider refreshing.", "warn");
        toast("Showing cached data. Hit Refresh to get latest from Excel.", "info");
      }
      return;
    }
    log("No cache found — fetching from server for the first time...");
  } else {
    log("Manual refresh requested — fetching from server...");
  }

  // ── Show loading states ───────────────────────────────────────────────────
  ["loading-monthly","loading-supplier","loading-status","loading-dept"].forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.textContent = "Loading..."; el.style.display = "flex"; }
  });

  const refreshBtn = document.getElementById("refresh-btn");
  if (refreshBtn) {
    refreshBtn.disabled = true;
    refreshBtn.innerHTML = `<i data-lucide="loader" style="width:14px;height:14px;"></i> Refreshing…`;
    if (window.lucide) lucide.createIcons();
  }

  const note = document.getElementById("data-note");
  if (note) note.textContent = "Fetching from Excel — this may take a moment…";

  try {
    log("Fetching /api/dashboard...");
    const res = await fetch("/api/dashboard");
    log(`Server responded: HTTP ${res.status}`, res.ok ? "ok" : "error");

    if (!res.ok) {
      const text = await res.text();
      log(`Error body: ${text.substring(0, 200)}`, "error");
      throw new Error(`HTTP ${res.status}: ${text.substring(0, 100)}`);
    }

    const data = await res.json();
    if (data.error) { log(`API error: ${data.error}`, "error"); throw new Error(data.error); }

    log(`Raw rows: ${data.rawRows ?? "?"}`, "ok");
    log(`Sheets detected: ${(data.sheetsDetected || []).join(", ")}`, "ok");

    const activeMonths = (data.monthlySpending || []).filter(m => m.total > 0);
    if (activeMonths.length) {
      log(`Months with data: ${activeMonths.map(m => `${m.month}=₱${m.total.toLocaleString()}`).join(", ")}`, "ok");
    } else {
      log("WARNING: No months have spending > 0", "warn");
    }

    if (data.supplierBreakdown?.length) {
      log(`Top supplier: ${data.supplierBreakdown[0]?.name} (₱${data.supplierBreakdown[0]?.total?.toLocaleString()})`, "ok");
    }

    saveCache(data);
    log("Saved to local cache — future loads will be instant.", "ok");
    updateCacheBadge();

    renderAll(data);
    toast("Dashboard updated and cached.", "success");

  } catch(e) {
    log(`FAILED: ${e.message}`, "error");
    toast("Refresh failed — showing cached data if available.", "error");

    const cached = loadCache();
    if (cached) {
      log("Falling back to cached data.", "warn");
      renderAll(cached);
      updateCacheBadge();
    } else {
      ["loading-monthly","loading-supplier","loading-status","loading-dept"].forEach(id => {
        showEmpty(id, "Error — see debug log");
      });
    }
  } finally {
    if (refreshBtn) {
      refreshBtn.disabled = false;
      refreshBtn.innerHTML = `<i data-lucide="refresh-cw" style="width:14px;height:14px;"></i> Refresh`;
      if (window.lucide) lucide.createIcons();
    }
  }
}

// ─── Raw API inspector ────────────────────────────────────────────────────────
async function inspectRaw() {
  const pre = document.getElementById("raw-response");
  pre.style.display = "block";
  pre.textContent = "Fetching...";
  try {
    const res  = await fetch("/api/dashboard");
    const data = await res.json();
    const summary = {
      rawRows:           data.rawRows,
      sheetsDetected:    data.sheetsDetected,
      monthlySpending:   data.monthlySpending,
      supplierBreakdown: data.supplierBreakdown?.slice(0, 5),
      statusCounts:      data.statusCounts,
      deptSpending:      data.deptSpending?.slice(0, 5),
      savingsData:       data.savingsData?.slice(0, 3),
      error:             data.error,
    };
    pre.textContent = JSON.stringify(summary, null, 2);
  } catch(e) {
    pre.textContent = "Failed: " + e.message;
  }
}

// ─── Render all charts ────────────────────────────────────────────────────────
function renderAll(data) {
  renderKPIs(data);
  renderMonthlyChart(data.monthlySpending);
  renderSupplierChart(data.supplierBreakdown);
  renderStatusChart(data.statusCounts);
  renderDeptChart(data.deptSpending);
  renderSavingsTable(data.savingsData);
}

// ─── KPI Cards ────────────────────────────────────────────────────────────────
function renderKPIs(data) {
  const total = data.monthlySpending.reduce((s, m) => s + (m.total || 0), 0);
  const totalSavings = (data.savingsData || []).reduce((s, r) => s + (parseFloat(r["TOTAL COST SAVINGS"]) || 0), 0);
  const poCount = Object.values(data.statusCounts).reduce((a, b) => a + b, 0);

  document.getElementById("kpi-total").textContent    = fmt(total);
  document.getElementById("kpi-po-count").textContent = poCount || data.rawRows || "—";
  document.getElementById("kpi-savings").textContent  = fmt(totalSavings);
  document.getElementById("kpi-suppliers").textContent = (data.supplierBreakdown || []).length || "—";
}

// ─── Monthly Spending ─────────────────────────────────────────────────────────
function renderMonthlyChart(monthlyData) {
  hideLoading("loading-monthly");

  if (!monthlyData?.some(m => m.total > 0)) {
    showEmpty("loading-monthly", "No spending data yet"); return;
  }

  if (charts.monthly) charts.monthly.destroy();
  const ctx = document.getElementById("chart-monthly").getContext("2d");

  charts.monthly = new Chart(ctx, {
    type: "bar",
    data: {
      labels: monthlyData.map(m => m.month),
      datasets: [{
        label: "Total Spending",
        data: monthlyData.map(m => m.total),
        backgroundColor: monthlyData.map(m => m.total > 0 ? "#1f6feb" : "#21262d"),
        borderColor:     monthlyData.map(m => m.total > 0 ? "#388bfd" : "#30363d"),
        borderWidth: 1, borderRadius: 4, hoverBackgroundColor: "#388bfd",
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend:{ display:false }, tooltip:{ callbacks:{ label: c => " " + fmt(c.parsed.y) } } },
      scales: {
        x: { grid:{ display:false } },
        y: { grid:{ color:"#21262d" }, ticks:{ callback: v => fmt(v) } },
      },
    },
  });
}

// ─── Supplier Doughnut ────────────────────────────────────────────────────────
function renderSupplierChart(supplierData) {
  if (!supplierData?.length) { showEmpty("loading-supplier", "No supplier data yet"); return; }
  hideLoading("loading-supplier");

  const top = supplierData.slice(0, 8);
  if (charts.supplier) charts.supplier.destroy();
  const ctx = document.getElementById("chart-supplier").getContext("2d");

  charts.supplier = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: top.map(s => s.name.length > 22 ? s.name.substring(0,20)+"…" : s.name),
      datasets: [{ data: top.map(s => s.total), backgroundColor: PALETTE, borderColor:"#161b22", borderWidth:2, hoverOffset:6 }],
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: "62%",
      plugins: {
        legend: { position:"right", labels:{ boxWidth:10, padding:10, font:{ size:10 } } },
        tooltip: { callbacks:{ label: c => ` ${c.label}: ${fmt(c.parsed)}` } },
      },
    },
  });
}

// ─── PO Status Doughnut ───────────────────────────────────────────────────────
function renderStatusChart(statusCounts) {
  const total = Object.values(statusCounts).reduce((a, b) => a + b, 0);
  if (!total) { showEmpty("loading-status", "No PO status data yet"); return; }
  hideLoading("loading-status");

  const labels = Object.keys(statusCounts).filter(k => statusCounts[k] > 0);
  const values = labels.map(k => statusCounts[k]);
  const colors = { "Served":"#3fb950", "For Delivery":"#388bfd", "Pending":"#d29922", "Cancelled":"#f85149" };

  if (charts.status) charts.status.destroy();
  const ctx = document.getElementById("chart-status").getContext("2d");

  charts.status = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels,
      datasets: [{ data:values, backgroundColor:labels.map(l=>colors[l]||"#8b949e"), borderColor:"#161b22", borderWidth:2, hoverOffset:6 }],
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: "60%",
      plugins: {
        legend: { position:"bottom", labels:{ boxWidth:10, padding:12, font:{ size:10 } } },
        tooltip: { callbacks:{ label: c => ` ${c.label}: ${c.parsed} (${Math.round(c.parsed/total*100)}%)` } },
      },
    },
  });
}

// ─── Department Bar ───────────────────────────────────────────────────────────
function renderDeptChart(deptData) {
  if (!deptData?.length) { showEmpty("loading-dept", "No department data yet"); return; }
  hideLoading("loading-dept");

  if (charts.dept) charts.dept.destroy();
  const ctx = document.getElementById("chart-dept").getContext("2d");

  charts.dept = new Chart(ctx, {
    type: "bar",
    data: {
      labels: deptData.map(d => d.name),
      datasets: [{
        label: "Spending",
        data: deptData.map(d => d.total),
        backgroundColor:"#1f6feb", borderColor:"#388bfd", borderWidth:1, borderRadius:4, hoverBackgroundColor:"#388bfd",
      }],
    },
    options: {
      indexAxis: "y", responsive: true, maintainAspectRatio: false,
      plugins: { legend:{ display:false }, tooltip:{ callbacks:{ label: c => " " + fmt(c.parsed.x) } } },
      scales: {
        x: { grid:{ color:"#21262d" }, ticks:{ callback: v => fmt(v) } },
        y: { grid:{ display:false } },
      },
    },
  });
}

// ─── Savings Table ────────────────────────────────────────────────────────────
function renderSavingsTable(savingsData) {
  const tbody = document.getElementById("savings-body");
  if (!savingsData?.length) {
    tbody.innerHTML = `<tr><td colspan="5" style="text-align:center;color:var(--text-3);padding:24px;">No savings data yet — fill in the Original Price column in your Excel to see comparisons</td></tr>`;
    return;
  }

  let totSavings = 0, totOriginal = 0;

  tbody.innerHTML = savingsData.filter(r => r["SUPPLIER'S NAME"]).map(row => {
    const original = parseFloat(row["ORIGINAL PRICE"])     || 0;
    const actual   = parseFloat(row["TOTAL AMOUNT"])       || 0;
    const savings  = parseFloat(row["TOTAL COST SAVINGS"]) || Math.max(0, original - actual);
    const pct      = original > 0 ? (savings / original * 100) : 0;
    totSavings  += savings;
    totOriginal += original;
    const cls = pct > 0 ? "diff-row-match" : (pct < 0 ? "diff-row-mismatch" : "");
    return `<tr>
      <td>${row["SUPPLIER'S NAME"]}</td>
      <td class="text-right mono">${fmt(original)}</td>
      <td class="text-right mono">${fmt(actual)}</td>
      <td class="text-right mono" style="color:var(--green)">${fmt(savings)}</td>
      <td class="text-right mono ${cls}">${pct.toFixed(1)}%</td>
    </tr>`;
  }).join("");

  const pct = totOriginal > 0 ? (totSavings / totOriginal * 100) : 0;
  tbody.innerHTML += `
    <tr style="border-top:2px solid var(--border);font-weight:700;">
      <td>TOTAL</td>
      <td class="text-right mono">${fmt(totOriginal)}</td>
      <td class="text-right mono">—</td>
      <td class="text-right mono" style="color:var(--green)">${fmt(totSavings)}</td>
      <td class="text-right mono" style="color:var(--green)">${pct.toFixed(1)}%</td>
    </tr>`;

  const kpiPct = document.getElementById("kpi-savings-pct");
  if (kpiPct && pct > 0) kpiPct.textContent = `${pct.toFixed(1)}% below original prices`;
}

document.addEventListener("DOMContentLoaded", () => loadDashboard(false));
