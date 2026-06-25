// public/js/farm.js
async function loadFarmSites() {
  if (farmSites.length) return farmSites;
  try {
    const r = await fetch("/api/farm-sites");
    farmSites = await r.json();
  } catch (e) { farmSites = []; }
  return farmSites;
}

function renderMonitorView() {
  if (!assetData.length) { loadAssets().then(() => _buildMonitorGrouped()); return; }
  _buildMonitorGrouped();
}

function _buildMonitorGrouped() {
  const farmMap = {};
  const farmTypes = {};
  assetData.forEach(a => {
    const f = (a.siteName || "ไม่ระบุไซต์").trim();
    if (!farmMap[f]) farmMap[f] = [];
    farmMap[f].push(a);
    if (a.farmType && a.farmType !== "-") farmTypes[f] = a.farmType;
  });
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
  let html = `<div class="mon-farm-item ${currentMonitorFarm === "ALL" ? "active" : ""}" onclick="selectMonitorFarm('ALL')">
      <div class="mon-farm-icon">🌐</div>
      <div class="mon-farm-name">ทุกฟาร์ม</div>
      <div class="mon-farm-cnt">${assetData.length}</div>
    </div>`;
  Object.keys(grouped).sort().forEach(type => {
    const farms = grouped[type].filter(f => f.toLowerCase().includes(searchTerm.toLowerCase()));
    if (farms.length === 0) return;
    const icon = typeIcons[type] || "🌱";
    html += `<div class="mon-group-label">${icon} ${type}</div>`;
    farms.forEach(f => {
      const cnt = farmMap[f]?.length || 0;
      const isAct = currentMonitorFarm === f;
      html += `<div class="mon-farm-item ${isAct ? "active" : ""}" onclick="selectMonitorFarm('${f.replace(/'/g, "\\'")}')">
        <div class="mon-farm-icon">${icon}</div>
        <div class="mon-farm-name">${f}</div>
        <div class="mon-farm-cnt">${cnt}</div>
      </div>`;
    });
  });
  list.innerHTML = html;
}

function filterFarmList(val) {
  const farmMap = {};
  assetData.forEach(a => {
    const f = (a.siteName || "ไม่ระบุไซต์").trim();
    if (!farmMap[f]) farmMap[f] = [];
    farmMap[f].push(a);
  });
  const farmTypes = {};
  assetData.forEach(a => { if (a.farmType && a.farmType !== "-") farmTypes[a.siteName] = a.farmType; });
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
  const farmMap = {};
  const farmTypes = {};
  assetData.forEach(a => {
    const f = (a.siteName || "ไม่ระบุไซต์").trim();
    if (!farmMap[f]) farmMap[f] = [];
    farmMap[f].push(a);
    if (a.farmType && a.farmType !== "-") farmTypes[f] = a.farmType;
  });
  const grouped = {};
  Object.keys(farmMap).sort().forEach(f => {
    const type = farmTypes[f] || "อื่นๆ";
    if (!grouped[type]) grouped[type] = [];
    grouped[type].push(f);
  });
  _renderFarmListGrouped(grouped, farmMap, document.getElementById("farmSearch").value || "");
  document.getElementById("monitorAssetSearch").value = "";
  _applyMonitorFarm(farm, farmMap);
}

function _applyMonitorFarm(farm, farmMap) {
  document.getElementById("monFarmLabel").textContent = farm === "ALL" ? "🌐 ทุกฟาร์ม" : "🌱 " + farm;
  const list = farm === "ALL" ? assetData : (farmMap[farm] || []);
  const ok = list.filter(a => a.status && a.status.includes("ใช้งานได้")).length;
  const rep = list.filter(a => a.status && a.status.includes("ซ่อม")).length;
  document.getElementById("monStatsRow").innerHTML = `
    <span class="mon-stat-chip msc-total">รวม ${list.length} ชิ้น</span>
    <span class="mon-stat-chip msc-ok">✅ ใช้งาน ${ok}</span>
    ${rep ? `<span class="mon-stat-chip msc-rep">🔧 ซ่อม ${rep}</span>` : ""}
  `;
  _renderMonitorTable(list);
}

function filterMonitorAsset() {
  const kw = (document.getElementById("monitorAssetSearch").value || "").toLowerCase().trim();
  const farmMap = {};
  assetData.forEach(a => {
    const f = (a.siteName || "ไม่ระบุไซต์").trim();
    if (!farmMap[f]) farmMap[f] = [];
    farmMap[f].push(a);
  });
  let base = currentMonitorFarm === "ALL" ? assetData : (farmMap[currentMonitorFarm] || []);
  if (kw) base = base.filter(a => (a.assetId || "").toLowerCase().includes(kw) || (a.code || "").toLowerCase().includes(kw) || (a.name || "").toLowerCase().includes(kw) || (a.serialNumber || "").toLowerCase().includes(kw));
  _renderMonitorTable(base);
}

function _renderMonitorTable(data) {
  const tb = document.getElementById("monitorAssetBody");
  if (!data.length) { tb.innerHTML = `<tr><td colspan="11" style="text-align:center;padding:48px;color:var(--tmuted)"><div style="font-size:32px;margin-bottom:10px">🏭</div>ไม่พบอุปกรณ์ในไซต์งานนี้</td></tr>`; return; }
  tb.innerHTML = data.map(a => {
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