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

// ─── Cache helpers — server-side file, shared across all devices ──────────────
// Cache lives at purchasing-tool/data/dashboard-cache.json
// Pass ?refresh=1 to force a rebuild from Excel

function fmtAge(ms) {
  const s = Math.floor(ms / 1000);
  if (s < 60)    return `${s}s ago`;
  if (s < 3600)  return `${Math.floor(s / 60)}m ago`;
  if (s < 86400) return `${Math.floor(s / 3600)}h ago`;
  return `${Math.floor(s / 86400)}d ago`;
}

function updateCacheBadge(data) {
  const badge = document.getElementById("cache-badge");
  const note  = document.getElementById("data-note");

  const source = data.fromCache ? (data.stale ? "stale cache" : "cache") : "live";
  const age    = data.cachedAt ? fmtAge(Date.now() - data.cachedAt) : "unknown";

  if (badge) {
    badge.textContent = data.fromCache
      ? `Cached · saved ${age}`
      : `Live · just updated`;
    badge.style.background = data.stale ? "#3a2e0a" : (data.fromCache ? "" : "#0a2010");
    badge.style.color      = data.stale ? "var(--yellow)" : (data.fromCache ? "" : "var(--green)");
    badge.style.display    = "inline";
  }

  if (note) {
    note.textContent = `${data.rawRows ?? "?"} records · ${(data.sheetsDetected || []).join(", ")} · source: ${source}`;
  }
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

  log(forceRefresh ? "Manual refresh — rebuilding from SQLite..." : "Loading dashboard...");

  // Show loading states
  ["loading-monthly","loading-supplier","loading-status","loading-dept"].forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.textContent = "Loading..."; el.style.display = "flex"; }
  });

  const refreshBtn = document.getElementById("refresh-btn");
  if (refreshBtn) {
    refreshBtn.disabled = true;
    refreshBtn.innerHTML = `<i data-lucide="loader" style="width:14px;height:14px;"></i> ${forceRefresh ? "Refreshing…" : "Loading…"}`;
    if (window.lucide) lucide.createIcons();
  }

  const note = document.getElementById("data-note");
  if (forceRefresh && note) note.textContent = "Rebuilding from SQLite — this may take a moment…";

  try {
    const url = forceRefresh ? "/api/dashboard?refresh=1" : "/api/dashboard";
    log(`Fetching ${url}...`);
    const res = await fetch(url);
    log(`Server responded: HTTP ${res.status}`, res.ok ? "ok" : "error");

    if (!res.ok) {
      const text = await res.text();
      log(`Error: ${text.substring(0, 200)}`, "error");
      throw new Error(`HTTP ${res.status}`);
    }

    const data = await res.json();
    if (data.error) { log(`API error: ${data.error}`, "error"); throw new Error(data.error); }

    log(`Source: ${data.fromCache ? (data.stale ? "stale cache" : "server cache") : "fresh from SQLite"}`, data.fromCache ? "info" : "ok");
    log(`Raw rows: ${data.rawRows ?? "?"}`, "ok");
    log(`Sheets: ${(data.sheetsDetected || []).join(", ")}`, "ok");

    const activeMonths = (data.monthlySpending || []).filter(m => m.total > 0);
    if (activeMonths.length) {
      log(`Months with data: ${activeMonths.map(m => `${m.month}=₱${m.total.toLocaleString()}`).join(", ")}`, "ok");
    } else {
      log("WARNING: No months have spending > 0", "warn");
    }

    updateCacheBadge(data);
    renderAll(data);

    if (forceRefresh && !data.fromCache) {
      toast("Dashboard refreshed and cache updated.", "success");
    } else if (data.stale) {
      toast("Showing stale cache — refresh failed. Try again.", "warn");
    }

  } catch(e) {
    log(`FAILED: ${e.message}`, "error");
    toast("Could not load dashboard — see debug log.", "error");
    ["loading-monthly","loading-supplier","loading-status","loading-dept"].forEach(id => {
      showEmpty(id, "Error — see debug log");
    });
  } finally {
    if (refreshBtn) {
      refreshBtn.disabled = false;
      refreshBtn.innerHTML = `<i data-lucide="refresh-cw" style="width:14px;height:14px;"></i> Refresh`;
      if (window.lucide) lucide.createIcons();
    }
  }

  // Supplier count is intentionally a separate call — it comes from the
  // supplier masterlist (same source as suppliers.html), not from the PO/
  // monitoring data behind /api/dashboard, so it isn't limited to only the
  // suppliers that happen to have POs on record.
  loadSupplierCount();
}

// ─── Supplier count (masterlist, same source as suppliers.html) ───────────────
async function loadSupplierCount() {
  const el = document.getElementById("kpi-suppliers");
  if (!el) return;
  try {
    const res = await fetch("/api/suppliers");
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    const list = Array.isArray(data) ? data : data.suppliers;
    if (Array.isArray(list)) {
      el.textContent = list.length.toLocaleString("en-PH");
      log(`Supplier count: ${list.length} (from /api/suppliers)`, "ok");
    }
  } catch (e) {
    log(`Supplier count fetch failed: ${e.message}`, "warn");
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
  // Savings = sum of abs(negative items) across all months — calculated in db.js
  const totalSavings = data.monthlySpending.reduce((s, m) => s + (m.savings || 0), 0);
<<<<<<< HEAD
  const poCount = Object.values(data.statusCounts).reduce((a, b) => a + b, 0);
=======
  // totalPOs is grouped by db.js the same way Monitoring's table groups rows
  // (one PO per non-blank PO NUMBER, each blank-PO-NUMBER row counted on its
  // own) — summing statusCounts used to undercount because that bucket data
  // came from a query that silently dropped blank-PO-number/blank-status rows.
  const poCount = data.totalPOs ?? Object.values(data.statusCounts).reduce((a, b) => a + b, 0);
>>>>>>> c16602719a7c662c5a52c25710cffccb199f9ed6

  document.getElementById("kpi-total").textContent    = fmt(total);
  document.getElementById("kpi-net").textContent      = fmt(total - totalSavings);
  document.getElementById("kpi-po-count").textContent = poCount || data.rawRows || "—";
  document.getElementById("kpi-savings").textContent  = fmt(totalSavings);
  // kpi-suppliers is populated separately by loadSupplierCount() from the
  // supplier masterlist, not from supplierBreakdown (which only reflects
  // suppliers that have POs in the monitoring data).
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
      datasets: [{ data: top.map(s => s.total), backgroundColor: PALETTE, borderColor: PALETTE, borderWidth:0, hoverOffset:6 }],
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
      datasets: [{ data:values, backgroundColor:labels.map(l=>colors[l]||"#8b949e"), borderColor:labels.map(l=>colors[l]||"#8b949e"), borderWidth:0, hoverOffset:6 }],
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
// Savings = sum of negative-priced items (discounts) per supplier
// Calculated in db.js from qty * unit_price where result < 0
function renderSavingsTable(savingsData) {
  const tbody = document.getElementById("savings-body");

  const rows = (savingsData || []).filter(r => r["SUPPLIER'S NAME"] && parseFloat(r["TOTAL COST SAVINGS"]) > 0);

  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="3" style="text-align:center;color:var(--text-3);padding:24px;">No discount savings recorded yet — discounts appear as negative-priced items in the monitoring</td></tr>`;
    return;
  }

  let totAmount = 0, totSavings = 0;

  tbody.innerHTML = rows.map(row => {
    const actual  = parseFloat(row["TOTAL AMOUNT"])       || 0;
    const savings = parseFloat(row["TOTAL COST SAVINGS"]) || 0;
    const pct     = actual > 0 ? (savings / actual * 100) : 0;
    totAmount  += actual;
    totSavings += savings;
    return `<tr>
      <td>${row["SUPPLIER'S NAME"]}</td>
      <td class="text-right mono">${fmt(actual)}</td>
      <td class="text-right mono" style="color:var(--green)">${fmt(savings)}</td>
    </tr>`;
  }).join("");

  const pct = totAmount > 0 ? (totSavings / totAmount * 100) : 0;
  tbody.innerHTML += `
    <tr style="border-top:2px solid var(--border);font-weight:700;">
      <td>TOTAL</td>
      <td class="text-right mono">${fmt(totAmount)}</td>
      <td class="text-right mono" style="color:var(--green)">${fmt(totSavings)}</td>
    </tr>`;

  const kpiPct = document.getElementById("kpi-savings-pct");
  if (kpiPct && pct > 0) kpiPct.textContent = `${pct.toFixed(1)}% saved via discounts`;
}

document.addEventListener("DOMContentLoaded", () => loadDashboard(false));