/**
 * charts.js
 * Dashboard charts — reads from /api/dashboard
 */

"use strict";

// ─── Chart.js Defaults ──────────────────────────────────────────────────────
Chart.defaults.color = "#8b949e";
Chart.defaults.borderColor = "#30363d";
Chart.defaults.font.family = "'IBM Plex Mono', monospace";
Chart.defaults.font.size = 11;

const PALETTE = [
  "#1f6feb", "#388bfd", "#3fb950", "#d29922",
  "#f0883e", "#a5a0f5", "#f85149", "#56d364",
  "#e3b341", "#79c0ff",
];

let charts = {};

function toast(msg, type = "info") {
  const icons = { success: "✅", error: "❌", warn: "⚠️", info: "ℹ️" };
  const container = document.getElementById("toasts");
  const el = document.createElement("div");
  el.className = `toast toast-${type}`;
  el.innerHTML = `<span class="toast-icon">${icons[type]}</span><span>${msg}</span>`;
  container.appendChild(el);
  setTimeout(() => el.remove(), 3500);
}

function fmt(num) {
  if (!num || isNaN(num)) return "₱0";
  return "₱" + Number(num).toLocaleString("en-PH", { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function hideLoading(id) {
  const el = document.getElementById(id);
  if (el) el.style.display = "none";
}

// ─── Main load ───────────────────────────────────────────────────────────────
async function loadDashboard() {
  try {
    const res  = await fetch("/api/dashboard");
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    renderAll(data);
  } catch (e) {
    console.error(e);
    toast("Could not load dashboard data. Is server running?", "error");
    ["loading-monthly","loading-supplier","loading-status","loading-dept"].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.textContent = "Server offline";
    });
  }
}

function renderAll(data) {
  renderKPIs(data);
  renderMonthlyChart(data.monthlySpending);
  renderSupplierChart(data.supplierBreakdown);
  renderStatusChart(data.statusCounts);
  renderDeptChart(data.deptSpending);
  renderSavingsTable(data.savingsData);
}

// ─── KPI Cards ───────────────────────────────────────────────────────────────
function renderKPIs(data) {
  const total = data.monthlySpending.reduce((s, m) => s + m.total, 0);
  const totalSavings = data.savingsData.reduce((s, r) => {
    return s + (parseFloat(r["TOTAL COST SAVINGS"]) || 0);
  }, 0);
  const supplierCount = data.supplierBreakdown.length;
  const poCount = Object.values(data.statusCounts).reduce((a, b) => a + b, 0);

  document.getElementById("kpi-total").textContent     = fmt(total);
  document.getElementById("kpi-po-count").textContent  = poCount;
  document.getElementById("kpi-savings").textContent   = fmt(totalSavings);
  document.getElementById("kpi-suppliers").textContent = supplierCount;
}

// ─── Monthly Spending Bar Chart ──────────────────────────────────────────────
function renderMonthlyChart(monthlyData) {
  hideLoading("loading-monthly");
  const labels = monthlyData.map(m => m.month);
  const values = monthlyData.map(m => m.total);

  if (charts.monthly) charts.monthly.destroy();
  const ctx = document.getElementById("chart-monthly").getContext("2d");

  const gradient = ctx.createLinearGradient(0, 0, 0, 280);
  gradient.addColorStop(0, "rgba(31,111,235,0.4)");
  gradient.addColorStop(1, "rgba(31,111,235,0.02)");

  charts.monthly = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: "Total Spending",
        data: values,
        backgroundColor: values.map(v => v > 0 ? "#1f6feb" : "#21262d"),
        borderColor: values.map(v => v > 0 ? "#388bfd" : "#30363d"),
        borderWidth: 1,
        borderRadius: 4,
        hoverBackgroundColor: "#388bfd",
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => " " + fmt(ctx.parsed.y),
          },
        },
      },
      scales: {
        x: { grid: { display: false } },
        y: {
          grid: { color: "#21262d" },
          ticks: { callback: v => fmt(v) },
        },
      },
    },
  });
}

// ─── Supplier Doughnut ───────────────────────────────────────────────────────
function renderSupplierChart(supplierData) {
  hideLoading("loading-supplier");
  if (!supplierData.length) {
    document.getElementById("loading-supplier").style.display = "flex";
    document.getElementById("loading-supplier").textContent = "No data yet";
    return;
  }

  const top = supplierData.slice(0, 8);
  if (charts.supplier) charts.supplier.destroy();
  const ctx = document.getElementById("chart-supplier").getContext("2d");

  charts.supplier = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: top.map(s => s.name.length > 20 ? s.name.substring(0,18)+"…" : s.name),
      datasets: [{
        data: top.map(s => s.total),
        backgroundColor: PALETTE,
        borderColor: "#161b22",
        borderWidth: 2,
        hoverOffset: 6,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "62%",
      plugins: {
        legend: {
          position: "right",
          labels: { boxWidth: 10, padding: 10, font: { size: 10 } },
        },
        tooltip: {
          callbacks: { label: ctx => ` ${ctx.label}: ${fmt(ctx.parsed)}` },
        },
      },
    },
  });
}

// ─── PO Status Doughnut ──────────────────────────────────────────────────────
function renderStatusChart(statusCounts) {
  hideLoading("loading-status");
  const labels = Object.keys(statusCounts);
  const values = Object.values(statusCounts);
  const total  = values.reduce((a, b) => a + b, 0);

  if (!total) {
    document.getElementById("loading-status").style.display = "flex";
    document.getElementById("loading-status").textContent = "No data yet";
    return;
  }

  const colors = ["#3fb950", "#d29922", "#f85149", "#388bfd"];
  if (charts.status) charts.status.destroy();
  const ctx = document.getElementById("chart-status").getContext("2d");

  charts.status = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: colors,
        borderColor: "#161b22",
        borderWidth: 2,
        hoverOffset: 6,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "60%",
      plugins: {
        legend: {
          position: "bottom",
          labels: { boxWidth: 10, padding: 12, font: { size: 10 } },
        },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.label}: ${ctx.parsed} (${Math.round(ctx.parsed/total*100)}%)`,
          },
        },
      },
    },
  });
}

// ─── Department Horizontal Bar ───────────────────────────────────────────────
function renderDeptChart(deptData) {
  hideLoading("loading-dept");
  if (!deptData.length) {
    document.getElementById("loading-dept").style.display = "flex";
    document.getElementById("loading-dept").textContent = "No data yet";
    return;
  }

  if (charts.dept) charts.dept.destroy();
  const ctx = document.getElementById("chart-dept").getContext("2d");

  charts.dept = new Chart(ctx, {
    type: "bar",
    data: {
      labels: deptData.map(d => d.name),
      datasets: [{
        label: "Total Spending",
        data: deptData.map(d => d.total),
        backgroundColor: "#1f6feb",
        borderColor: "#388bfd",
        borderWidth: 1,
        borderRadius: 4,
        hoverBackgroundColor: "#388bfd",
      }],
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: { label: ctx => " " + fmt(ctx.parsed.x) },
        },
      },
      scales: {
        x: {
          grid: { color: "#21262d" },
          ticks: { callback: v => fmt(v) },
        },
        y: { grid: { display: false } },
      },
    },
  });
}

// ─── Savings Table ───────────────────────────────────────────────────────────
function renderSavingsTable(savingsData) {
  const tbody = document.getElementById("savings-body");

  if (!savingsData.length) {
    tbody.innerHTML = `<tr><td colspan="5" style="text-align:center;color:var(--text-3);padding:24px;">No savings data yet</td></tr>`;
    return;
  }

  let totalSavings = 0;
  let totalOriginal = 0;

  tbody.innerHTML = savingsData.map(row => {
    const original = parseFloat(row["ORIGINAL PRICE"])     || 0;
    const actual   = parseFloat(row["TOTAL AMOUNT"])       || 0;
    const savings  = parseFloat(row["TOTAL COST SAVINGS"]) || (original - actual);
    const pct      = original > 0 ? (savings / original * 100) : 0;

    totalSavings  += savings;
    totalOriginal += original;

    const pctCls = pct > 0 ? "diff-row-match" : (pct < 0 ? "diff-row-mismatch" : "");

    return `
      <tr>
        <td>${row["SUPPLIER'S NAME"] || "—"}</td>
        <td class="text-right mono">${fmt(original)}</td>
        <td class="text-right mono">${fmt(actual)}</td>
        <td class="text-right mono" style="color:var(--green)">${fmt(savings)}</td>
        <td class="text-right mono ${pctCls}">${pct.toFixed(1)}%</td>
      </tr>`;
  }).join("");

  // Summary row
  const overallPct = totalOriginal > 0 ? (totalSavings / totalOriginal * 100) : 0;
  tbody.innerHTML += `
    <tr style="border-top:2px solid var(--border);font-weight:700;">
      <td>TOTAL</td>
      <td class="text-right mono">${fmt(totalOriginal)}</td>
      <td class="text-right mono">—</td>
      <td class="text-right mono" style="color:var(--green)">${fmt(totalSavings)}</td>
      <td class="text-right mono" style="color:var(--green)">${overallPct.toFixed(1)}%</td>
    </tr>`;

  const kpiPct = document.getElementById("kpi-savings-pct");
  if (kpiPct) kpiPct.textContent = `${overallPct.toFixed(1)}% below original prices`;
}

// ─── Init ─────────────────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", loadDashboard);
