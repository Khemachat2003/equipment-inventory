// public/js/dashboard.js
async function loadDashboardFull() {
  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 30000);
    const res = await fetch('/api/dashboard-full', { signal: controller.signal });
    clearTimeout(timeoutId);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const d = await res.json();

    document.getElementById('totalItems').textContent = d.totalItems || 0;
    document.getElementById('totalOffice').textContent = d.totalOffice || 0;
    document.getElementById('totalSite').textContent = d.totalSite || 0;
    document.getElementById('todayBorrow').textContent = d.todayBorrow || 0;
    document.getElementById('todayReturn').textContent = d.todayReturn || 0;
    document.getElementById('totalFarms').textContent = d.totalFarms || 0;
    document.getElementById('totalFarmAssets').textContent = d.totalFarmAssets || 0;
    document.getElementById('totalStockAssets').textContent = d.totalStockAssets || 0;
    if (d.topFarms && d.topFarms.length > 0) {
      document.getElementById('topFarmName').textContent = `${d.topFarms[0].name} (${d.topFarms[0].count})`;
    }
    if (d.chartData) {
      allChartData = d.chartData;
      renderChart(currentChartRange);
    }
    if (d.farmSites) farmSites = d.farmSites;
    if (d.partCatalog) partCatalog = d.partCatalog;
    if (d.assets) assetData = d.assets.map(a => ({ ...a }));
    console.log('✅ Dashboard loaded with 1 API call');
  } catch (e) {
    console.error('Dashboard Full error:', e);
    if (e.name === 'AbortError') {
      console.warn('⏱️ Dashboard Full timeout, using fallback APIs');
    }
    try { await loadDashboard(); } catch (_) {}
    try { await loadDashboardExtended(); } catch (_) {}
    try { await loadHistoryChart(); } catch (_) {}
  }
}

async function loadDashboard() {
  try {
    const res = await fetch("/api/dashboard");
    const d = await res.json();
    console.log('📊 Dashboard data:', d);
    ["totalItems", "totalOffice", "totalSite", "todayBorrow", "todayReturn"].forEach(k => {
      document.getElementById(k).textContent = d[k] ?? 0;
    });
  } catch (e) {
    console.error('loadDashboard error:', e);
  }
}

async function loadDashboardExtended() {
  try {
    const res = await fetch('/api/dashboard-asset');
    const d = await res.json();
    console.log('📊 Dashboard asset data:', d);
    document.getElementById('totalFarms').textContent = d.totalFarms || 0;
    document.getElementById('totalFarmAssets').textContent = d.totalFarm || 0;
    document.getElementById('totalStockAssets').textContent = d.totalStock || 0;
    if (d.topFarms && d.topFarms.length > 0) {
      document.getElementById('topFarmName').textContent = `${d.topFarms[0].name} (${d.topFarms[0].count})`;
    } else {
      document.getElementById('topFarmName').textContent = '-';
    }
  } catch (e) {
    console.error('Dashboard extended error:', e);
  }
}

async function loadHistoryChart() {
  try {
    const res = await fetch("/api/dashboard-stats");
    allChartData = await res.json();
    console.log('📊 Chart data:', Object.keys(allChartData).length, 'days');
    renderChart(currentChartRange);
  } catch (e) {
    console.error('loadHistoryChart error:', e);
  }
}

function setChartRange(days, btn) {
  currentChartRange = days;
  document.querySelectorAll(".cr-btn").forEach(b => b.classList.remove("active"));
  btn.classList.add("active");
  renderChart(days);
}

function renderChart(days) {
  if (!allChartData || !Object.keys(allChartData).length) return;
  const all = Object.keys(allChartData).sort();
  const labels = all.slice(-days);
  const bD = labels.map(d => allChartData[d]?.borrow || 0);
  const rD = labels.map(d => allChartData[d]?.return || 0);
  const ctx = document.getElementById("historyChart").getContext("2d");
  if (historyChart) historyChart.destroy();
  const bG = ctx.createLinearGradient(0, 0, 0, 280);
  bG.addColorStop(0, "rgba(37,99,235,.3)");
  bG.addColorStop(1, "rgba(37,99,235,0)");
  const rG = ctx.createLinearGradient(0, 0, 0, 280);
  rG.addColorStop(0, "rgba(13,148,136,.3)");
  rG.addColorStop(1, "rgba(13,148,136,0)");
  historyChart = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [
        { label: "เบิก", data: bD, borderColor: "#2563eb", backgroundColor: bG, fill: true, tension: .35, pointRadius: 2, pointHoverRadius: 5, borderWidth: 2 },
        { label: "คืน", data: rD, borderColor: "#0d9488", backgroundColor: rG, fill: true, tension: .35, pointRadius: 2, pointHoverRadius: 5, borderWidth: 2 }
      ]
    },
    options: {
      responsive: true,
      interaction: { intersect: false, mode: "index" },
      plugins: {
        legend: { position: "top", labels: { usePointStyle: true, padding: 20, font: { size: 12, weight: "600" } } },
        tooltip: { backgroundColor: "#1e293b", titleColor: "#f1f5f9", bodyColor: "#94a3b8", cornerRadius: 10, padding: 12 }
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 10 }, color: "#94a3b8", maxTicksLimit: 10 } },
        y: { beginAtZero: true, grid: { color: "rgba(0,0,0,.04)" }, ticks: { precision: 0, font: { size: 10 }, color: "#94a3b8" } }
      }
    }
  });
}

function refreshDashboard() {
  loadDashboardFull();
}