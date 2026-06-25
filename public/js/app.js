// public/js/app.js
function initApp() {
  loadCurrentUser();
  loadDashboardFull();
  loadHistory();
  loadAssets();
  const last = localStorage.getItem("activeTab") || "dashboard";
  openTab(last);
  setInterval(refreshDashboard, 120000);
}

function openTab(id) {
  document.querySelectorAll(".tab").forEach(t => { t.classList.remove("active");
    t.style.display = "none"; });
  document.querySelectorAll(".nav-btn").forEach(b => b.classList.remove("active"));
  const t = document.getElementById(id);
  if (!t) { openTab("dashboard"); return; }
  t.classList.add("active");
  t.style.display = "block";
  const nb = document.getElementById("nav-" + id);
  if (nb) nb.classList.add("active");
  localStorage.setItem("activeTab", id);
  if (id === "stock" && !_stockLoaded) {
    loadStock();
    _stockLoaded = true;
  }
  if (id === "farmMonitorTab") renderMonitorView();
  if (id === "asset") renderAssetSidebar();
}

// ---- Keyboard Shortcuts ----
document.addEventListener("keydown", e => {
  if (e.key === "Enter" && document.getElementById("loginPage").style.display !== "none") login();
});

// ---- DOM Ready & Check Auth ----
document.addEventListener("DOMContentLoaded", () => {
  // checkScanParam (ถ้ามี)
  const p = new URLSearchParams(window.location.search);
  const serial = p.get("serial");
  if (serial) {
    openTab("asset");
    const inp = document.getElementById("partSearch");
    if (inp) { inp.value = serial; }
  }
});

window.onload = checkAuth;