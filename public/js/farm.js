// public/js/farm.js
async function loadFarmSites() {
  if (farmSites.length) return farmSites;
  try {
    const r = await fetch("/api/farm-sites");
    farmSites = await r.json();
  } catch (e) { farmSites = []; }
  return farmSites;
}

// ---- Bundle data for Farm Monitor ----
// Bundle ไม่ใช่ฟาร์ม และไม่นับเป็นอุปกรณ์เดี่ยวๆ ของฟาร์ม — ใช้แค่เพื่อโชว์ว่า
// ฟาร์มไหนมี Bundle ไหน deploy อยู่ (แสดงเป็น 1 แถวต่อ 1 Bundle เท่านั้น)
let monitorBundles = [];

async function loadMonitorBundles() {
  try {
    const r = await fetch("/api/bundles");
    monitorBundles = await r.json();
  } catch (e) { monitorBundles = []; }
  return monitorBundles;
}

function _bundlesForFarm(farm) {
  const deployed = monitorBundles.filter(b => b.status === "Deployed" && b.location);
  if (farm === "ALL") return deployed;
  return deployed.filter(b => (b.location || "").trim() === farm);
}

function _famEsc(s) {
  return String(s == null ? "" : s).replace(/'/g, "\\'");
}

// เปิดหน้า Bundle แล้วพาไปดูรายละเอียดของ Bundle นั้นทันที (ใช้ flow ย้าย/ดูรายละเอียดเดิมของ Bundle System)
function openBundleFromMonitor(bundleId) {
  openTab("bundle");
  if (window.BDL) {
    BDL.refresh().then(() => BDL.openDetail(bundleId)).catch(() => BDL.openDetail(bundleId));
  }
}

// ---- Group assets by real farm site (ไม่รวมอุปกรณ์ที่อยู่ใน Bundle — จะโชว์ผ่าน Bundle แทน) ----
function _buildFarmMap() {
  const farmMap = {};
  const farmTypes = {};
  assetData.forEach(a => {
    if (a.bundleId) return; // อยู่ใน Bundle อยู่แล้ว ไม่นับเป็นอุปกรณ์เดี่ยวของฟาร์มนี้ตรงๆ
    const f = (a.siteName || "ไม่ระบุไซต์").trim();
    if (!farmMap[f]) farmMap[f] = [];
    farmMap[f].push(a);
    if (a.farmType && a.farmType !== "-") farmTypes[f] = a.farmType;
  });
  // ฟาร์มที่มีแค่ Bundle deploy อยู่ (ไม่มีอุปกรณ์เดี่ยวๆ เลย) ก็ต้องยังโผล่ในรายการฟาร์ม
  monitorBundles.forEach(b => {
    if (b.status === "Deployed" && b.location) {
      const f = b.location.trim();
      if (f && !farmMap[f]) farmMap[f] = [];
    }
  });
  return { farmMap, farmTypes };
}

function renderMonitorView() {
  const assetsReady = assetData.length ? Promise.resolve() : loadAssets();
  Promise.all([assetsReady, loadMonitorBundles()]).then(() => _buildMonitorGrouped());
}

function _buildMonitorGrouped() {
  const { farmMap, farmTypes } = _buildFarmMap();
  allFarms = Object.keys(farmMap).sort();
  const grouped = {};
  allFarms.forEach(f => {
    const type = farmTypes[f] || "อื่นๆ";
    if (!grouped[type]) grouped[type] = [];
    grouped[type].push(f);
  });
  _renderFarmListGrouped(grouped, farmMap, document.getElementById("farmSearch").value || "");
  _applyMonitorFarm(currentMonitorFarm, farmMap);
}

function _renderFarmListGrouped(grouped, farmMap, searchTerm) {
  const list = document.getElementById("monFarmList");
  const typeIcons = { "สัตว์ปีก": "🐔", "สัตว์บก": "🐄", "สุกร": "🐷", "อื่นๆ": "⚙️", "ไม่ระบุ": "🏭" };
  const totalCnt = assetData.filter(a => !a.bundleId).length; // ไม่นับอุปกรณ์ใน Bundle
  let html = `<div class="mon-farm-item ${currentMonitorFarm === "ALL" ? "active" : ""}" onclick="selectMonitorFarm('ALL')">
      <div class="mon-farm-icon">🌐</div>
      <div class="mon-farm-name">ทุกฟาร์ม</div>
      <div class="mon-farm-cnt">${totalCnt}</div>
    </div>`;
  Object.keys(grouped).sort().forEach(type => {
    const farms = grouped[type].filter(f => f.toLowerCase().includes(searchTerm.toLowerCase()));
    if (farms.length === 0) return;
    const icon = typeIcons[type] || "🌱";
    html += `<div class="mon-group-label">${icon} ${type}</div>`;
    farms.forEach(f => {
      const cnt = farmMap[f]?.length || 0;
      const isAct = currentMonitorFarm === f;
      const hasBundle = _bundlesForFarm(f).length > 0;
      html += `<div class="mon-farm-item ${isAct ? "active" : ""}" onclick="selectMonitorFarm('${f.replace(/'/g, "\\'")}')">
        <div class="mon-farm-icon">${icon}</div>
        <div class="mon-farm-name">${f}${hasBundle ? " 📦" : ""}</div>
        <div class="mon-farm-cnt">${cnt}</div>
      </div>`;
    });
  });
  list.innerHTML = html;
}

function filterFarmList(val) {
  const { farmMap, farmTypes } = _buildFarmMap();
  const grouped = {};
  Object.keys(farmMap).sort().forEach(f => {
    const type = farmTypes[f] || "อื่นๆ";
    if (!grouped[type]) grouped[type] = [];
    grouped[type].push(f);
  });
  _renderFarmListGrouped(grouped, farmMap, val);
}

function selectMonitorFarm(farm) {
  currentMonitorFarm = farm;
  const { farmMap, farmTypes } = _buildFarmMap();
  const grouped = {};
  Object.keys(farmMap).sort().forEach(f => {
    const type = farmTypes[f] || "อื่นๆ";
    if (!grouped[type]) grouped[type] = [];
    grouped[type].push(f);
  });
  _renderFarmListGrouped(grouped, farmMap, document.getElementById("farmSearch").value || "");
  document.getElementById("monitorAssetSearch").value = "";
  _applyMonitorFarm(farm, farmMap);
  collapseMonSidebarMobile("farmMonSidebar");
}

function _applyMonitorFarm(farm, farmMap) {
  document.getElementById("monFarmLabel").textContent = farm === "ALL" ? "🌐 ทุกฟาร์ม" : "🌱 " + farm;
  const farmSbCurrent = document.getElementById("farmSbCurrent");
  if (farmSbCurrent) farmSbCurrent.textContent = farm === "ALL" ? "ทุกฟาร์ม" : farm;

  const list = farm === "ALL" ? assetData.filter(a => !a.bundleId) : (farmMap[farm] || []);
  const bundles = _bundlesForFarm(farm);
  const ok = list.filter(a => a.status && a.status.includes("ใช้งานได้")).length;
  const rep = list.filter(a => a.status && a.status.includes("ซ่อม")).length;
  document.getElementById("monStatsRow").innerHTML = `
    <span class="mon-stat-chip msc-total">รวม ${list.length} ชิ้น</span>
    <span class="mon-stat-chip msc-ok">✅ ใช้งาน ${ok}</span>
    ${rep ? `<span class="mon-stat-chip msc-rep">🔧 ซ่อม ${rep}</span>` : ""}
    ${bundles.length ? `<span class="mon-stat-chip msc-total">📦 Bundle ${bundles.length} ชุด</span>` : ""}
  `;
  _renderMonitorTable(list, bundles);
}

function filterMonitorAsset() {
  const kw = (document.getElementById("monitorAssetSearch").value || "").toLowerCase().trim();
  const { farmMap } = _buildFarmMap();
  let base = currentMonitorFarm === "ALL" ? assetData.filter(a => !a.bundleId) : (farmMap[currentMonitorFarm] || []);
  let bundles = _bundlesForFarm(currentMonitorFarm);
  if (kw) {
    base = base.filter(a => (a.assetId || "").toLowerCase().includes(kw) || (a.code || "").toLowerCase().includes(kw) || (a.name || "").toLowerCase().includes(kw) || (a.serialNumber || "").toLowerCase().includes(kw));
    bundles = bundles.filter(b => (b.bundleId || "").toLowerCase().includes(kw) || (b.bundleName || "").toLowerCase().includes(kw));
  }
  _renderMonitorTable(base, bundles);
}

function _renderMonitorTable(data, bundles) {
  bundles = bundles || [];
  const tb = document.getElementById("monitorAssetBody");
  if (!data.length && !bundles.length) { tb.innerHTML = `<tr><td colspan="11" style="text-align:center;padding:48px;color:var(--tmuted)"><div style="font-size:32px;margin-bottom:10px">🏭</div>ไม่พบอุปกรณ์ในไซต์งานนี้</td></tr>`; return; }

  // Bundle จะโชว์เป็น 1 แถวต่อ 1 ชุด (ไม่แตกเป็นอุปกรณ์ย่อย) — กดเพื่อดูรายละเอียดในนั้น
  const bundleRows = bundles.map(b => {
    const cnt = b.assetIds ? b.assetIds.length : 0;
    const id = _famEsc(b.bundleId);
    return `<tr class="mon-bundle-row" style="cursor:pointer;background:var(--blue-l)" onclick="openBundleFromMonitor('${id}')">
      <td><span style="font-family:monospace;font-size:12px">${b.bundleId || "-"}</span></td>
      <td>-</td>
      <td style="font-weight:700">📦 ${b.bundleName || "-"}</td>
      <td><span style="font-size:12px;color:var(--blue)">${cnt} ชิ้นในชุด</span></td>
      <td><span class="badge bg-green">${b.status || "-"}</span></td>
      <td>${b.location || "-"}</td>
      <td>${b.location || "-"}</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td><button class="btn btn-blue btn-sm" onclick="event.stopPropagation();openBundleFromMonitor('${id}')">📦 ดูชุด</button></td>
    </tr>`;
  }).join("");

  const assetRows = data.map(a => {
    const s = a.serialNumber || "",
      enc = encodeURIComponent(s);
    return `<tr>
      <td><span style="font-family:monospace;font-size:12px">${a.assetId || "-"}</span></td>
      <td><span style="font-family:monospace;font-size:12px">${a.code || "-"}</span></td>
      <td style="font-weight:500">${a.name || "-"}</td>
      <td><span style="font-family:monospace;font-size:12px;color:var(--blue)">${s || "-"}</span></td>
      <td>${getStatusBadge(a.status || "")}</td>
      <td>${a.location || "-"}</td>
      <td>${a.siteName || "-"}</td>
      <td>${a.user || "-"}</td>
      <td><a href="/trace.html?serial=${enc}&from=internal" target="_blank" class="btn btn-out btn-sm">📜</a></td>
      <td><a href="/qr.html?serial=${enc}" target="_blank" class="btn btn-teal btn-sm">📷</a></td>
      <td><button class="btn btn-blue btn-sm" onclick="openTransferModal('${s}','${(a.status || "").replace(/'/g, "\\'")}','${(a.location || "").replace(/'/g, "\\'")}','${(a.siteName || "").replace(/'/g, "\\'")}','${(a.user || "").replace(/'/g, "\\'")}')">🚚</button></td>
    </tr>`;
  }).join("");

  tb.innerHTML = bundleRows + assetRows;
}

// ---- Farm Site & House Modals ----
async function addFarmSite() {
  const siteId = document.getElementById("farmSiteId").value.trim();
  const siteName = document.getElementById("farmSiteName").value.trim();
  if (!siteId || !siteName) { alert("กรุณากรอกรหัสและชื่อฟาร์ม"); return; }
  const body = {
    siteId,
    siteName,
    farmType: document.getElementById("farmSiteType").value,
    province: document.getElementById("farmSiteProvince").value.trim(),
    manager: document.getElementById("farmSiteManager").value.trim(),
    note: document.getElementById("farmSiteNote").value.trim()
  };
  try {
    const r = await fetch("/api/add-farm-site", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
    const d = await r.json();
    if (d.success) {
      alert("✅ เพิ่มฟาร์มสำเร็จ");
      closeModal("addFarmModal");
      ["farmSiteId", "farmSiteName", "farmSiteProvince", "farmSiteManager", "farmSiteNote"].forEach(id => document.getElementById(id).value = "");
      loadFarmSites();
    } else {
      alert("❌ " + (d.error || "เกิดข้อผิดพลาด"));
    }
  } catch (e) { alert("ไม่สามารถเชื่อมต่อได้"); }
}

async function addFarmHouse() {
  const houseId = document.getElementById("houseId").value.trim();
  const siteId = document.getElementById("houseSiteId").value.trim();
  const houseName = document.getElementById("houseName").value.trim();
  if (!houseId || !siteId || !houseName) { alert("กรุณากรอกข้อมูลให้ครบ"); return; }
  const body = {
    houseId,
    siteId,
    houseName,
    houseType: document.getElementById("houseType").value,
    capacity: document.getElementById("houseCapacity").value.trim(),
    note: document.getElementById("houseNote").value.trim()
  };
  try {
    const r = await fetch("/api/add-farm-house", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
    const d = await r.json();
    if (d.success) {
      alert("✅ เพิ่มโรงเรือนสำเร็จ");
      closeModal("addHouseModal");
      ["houseId", "houseSiteId", "houseName", "houseCapacity", "houseNote"].forEach(id => document.getElementById(id).value = "");
    } else {
      alert("❌ " + (d.error || "เกิดข้อผิดพลาด"));
    }
  } catch (e) { alert("ไม่สามารถเชื่อมต่อได้"); }
}
