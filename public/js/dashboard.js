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
  const canvas = document.getElementById("historyChart");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (historyChart) historyChart.destroy();

  // Premium design-system gradients (blue #1B6CA8 / emerald #00C896)
  const cH = canvas.parentElement ? canvas.parentElement.clientHeight || 280 : 280;
  const bG = ctx.createLinearGradient(0, 0, 0, cH);
  bG.addColorStop(0, "rgba(27,108,168,.28)");
  bG.addColorStop(.55, "rgba(27,108,168,.10)");
  bG.addColorStop(1, "rgba(27,108,168,0)");
  const rG = ctx.createLinearGradient(0, 0, 0, cH);
  rG.addColorStop(0, "rgba(0,200,150,.24)");
  rG.addColorStop(.55, "rgba(0,200,150,.08)");
  rG.addColorStop(1, "rgba(0,200,150,0)");

  historyChart = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "เบิก",
          data: bD,
          borderColor: "#1B6CA8",
          backgroundColor: bG,
          fill: true,
          tension: .45,
          pointRadius: 3,
          pointHoverRadius: 6,
          pointBackgroundColor: "#fff",
          pointBorderColor: "#1B6CA8",
          pointBorderWidth: 2,
          pointHoverBorderColor: "#1B6CA8",
          pointHoverBorderWidth: 3,
          pointHoverBackgroundColor: "#1B6CA8",
          borderWidth: 2.5,
          borderCapStyle: "round",
          borderJoinStyle: "round"
        },
        {
          label: "คืน",
          data: rD,
          borderColor: "#00C896",
          backgroundColor: rG,
          fill: true,
          tension: .45,
          pointRadius: 3,
          pointHoverRadius: 6,
          pointBackgroundColor: "#fff",
          pointBorderColor: "#00C896",
          pointBorderWidth: 2,
          pointHoverBorderColor: "#00C896",
          pointHoverBorderWidth: 3,
          pointHoverBackgroundColor: "#00C896",
          borderWidth: 2.5,
          borderCapStyle: "round",
          borderJoinStyle: "round"
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { intersect: false, mode: "index" },
      animation: { duration: 900, easing: "easeOutQuart" },
      plugins: {
        legend: {
          position: "top",
          align: "end",
          labels: {
            usePointStyle: true,
            pointStyle: "circle",
            padding: 22,
            boxWidth: 8,
            boxHeight: 8,
            color: "#475569",
            font: { size: 12, weight: "600", family: "Inter, sans-serif" }
          }
        },
        tooltip: {
          backgroundColor: "#0F172A",
          titleColor: "#F8FAFC",
          bodyColor: "#E2E8F0",
          borderColor: "rgba(255,255,255,.08)",
          borderWidth: 1,
          cornerRadius: 12,
          padding: 14,
          boxPadding: 6,
          displayColors: true,
          usePointStyle: true,
          titleFont: { size: 12, weight: "700" },
          bodyFont: { size: 13, weight: "600" },
          caretSize: 6,
          titleMarginBottom: 8
        }
      },
      scales: {
        x: {
          grid: { display: false },
          border: { display: false },
          ticks: { font: { size: 10, family: "Inter, sans-serif" }, color: "#94A3B8", maxTicksLimit: 12, maxRotation: 0 }
        },
        y: {
          beginAtZero: true,
          border: { display: false },
          grid: { color: "rgba(100,116,139,.10)", drawTicks: false },
          ticks: { precision: 0, font: { size: 10, family: "Inter, sans-serif" }, color: "#94A3B8", padding: 8 }
        }
      }
    }
  });
}

function refreshDashboard() {
  loadDashboardFull();
}