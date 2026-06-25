// public/js/asset.js
// ---- Asset Sidebar & Table ----
async function loadAssets() {
  try {
    const r = await fetch("/api/assets");
    if (!r.ok) throw new Error(r.status);
    assetData = await r.json();
    renderAssetSidebar();
  } catch (e) {}
}

function renderAssetSidebar() {
  if (!assetData.length) { loadAssets().then(() => _buildAssetSidebar()); return; }
  _buildAssetSidebar();
}

function _buildAssetSidebar() {
  const partMap = {};
  assetData.forEach(a => {
    const pn = a.partNumber || "ไม่ระบุ";
    if (!partMap[pn]) partMap[pn] = { partNumber: pn, name: a.name, items: [] };
    partMap[pn].items.push(a);
  });
  const allParts = Object.keys(partMap).sort();
  if (currentPart === "ALL" && allParts.length > 0) currentPart = allParts[0];
  _renderPartList(allParts, partMap, document.getElementById("partSearch").value || "");
  _applyPart(currentPart, partMap);
}

function _renderPartList(parts, partMap, searchTerm) {
  const list = document.getElementById("partList");
  const icons = ["📦", "🔧", "⚙️", "🔩", "🗃️", "💻", "🖥️", "⌨️", "🖨️", "📡"];
  let html = `<div class="mon-farm-item ${currentPart === "ALL" ? "active" : ""}" onclick="selectPart('ALL')">
      <div class="mon-farm-icon">📋</div>
      <div class="mon-farm-name">ทุก Part</div>
      <div class="mon-farm-cnt">${assetData.length}</div>
    </div>`;
  parts.filter(p => p.toLowerCase().includes(searchTerm.toLowerCase())).forEach((p, i) => {
    const cnt = partMap[p]?.items.length || 0;
    const ic = icons[i % icons.length];
    const isAct = currentPart === p;
    html += `<div class="mon-farm-item ${isAct ? "active" : ""}" onclick="selectPart('${p.replace(/'/g, "\\'")}')">
      <div class="mon-farm-icon">${ic}</div>
      <div class="mon-farm-name">${p}</div>
      <div class="mon-farm-cnt">${cnt}</div>
    </div>`;
  });
  list.innerHTML = html;
}

function filterPartList(val) {
  const partMap = {};
  assetData.forEach(a => {
    const pn = a.partNumber || "ไม่ระบุ";
    if (!partMap[pn]) partMap[pn] = { partNumber: pn, name: a.name, items: [] };
    partMap[pn].items.push(a);
  });
  _renderPartList(Object.keys(partMap).sort(), partMap, val);
}

function selectPart(part) {
  currentPart = part;
  const partMap = {};
  assetData.forEach(a => {
    const pn = a.partNumber || "ไม่ระบุ";
    if (!partMap[pn]) partMap[pn] = { partNumber: pn, name: a.name, items: [] };
    partMap[pn].items.push(a);
  });
  _renderPartList(Object.keys(partMap).sort(), partMap, document.getElementById("partSearch").value || "");
  document.getElementById("partAssetSearch").value = "";
  _applyPart(part, partMap);
}

function _applyPart(part, partMap) {
  const list = part === "ALL" ? assetData : (partMap[part]?.items || []);
  const name = part === "ALL" ? "ทุก Part" : part;
  document.getElementById("partLabel").textContent = `📦 ${name}`;
  const ok = list.filter(a => a.status && a.status.includes("ใช้งานได้")).length;
  const rep = list.filter(a => a.status && a.status.includes("ซ่อม")).length;
  document.getElementById("partStatsRow").innerHTML = `
    <span class="mon-stat-chip msc-total">รวม ${list.length} ชิ้น</span>
    <span class="mon-stat-chip msc-ok">✅ ใช้งาน ${ok}</span>
    ${rep ? `<span class="mon-stat-chip msc-rep">🔧 ซ่อม ${rep}</span>` : ""}
  `;
  _renderPartTable(list);
}

function filterPartAssets() {
  const kw = (document.getElementById("partAssetSearch").value || "").toLowerCase().trim();
  const partMap = {};
  assetData.forEach(a => {
    const pn = a.partNumber || "ไม่ระบุ";
    if (!partMap[pn]) partMap[pn] = { partNumber: pn, name: a.name, items: [] };
    partMap[pn].items.push(a);
  });
  let base = currentPart === "ALL" ? assetData : (partMap[currentPart]?.items || []);
  if (kw) base = base.filter(a => (a.assetId || "").toLowerCase().includes(kw) || (a.code || "").toLowerCase().includes(kw) || (a.name || "").toLowerCase().includes(kw) || (a.serialNumber || "").toLowerCase().includes(kw));
  _renderPartTable(base);
}

function _renderPartTable(data) {
  const tb = document.getElementById("partAssetBody");
  if (!data.length) { tb.innerHTML = `<tr><td colspan="11" style="text-align:center;padding:48px;color:var(--tmuted)"><div style="font-size:32px;margin-bottom:10px">📦</div>ไม่พบอุปกรณ์ใน Part นี้</td></tr>`; return; }
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

// ---- Transfer Modal ----
function openTransferModal(serial, status, location, site, user) {
  document.getElementById("transferSerialTitle").textContent = serial;
  document.getElementById("transferStatus").value = status;
  document.getElementById("transferLocation").value = location;
  document.getElementById("transferSite").value = site;
  document.getElementById("transferUser").value = user;
  document.getElementById("oldSiteName").value = site;
  document.getElementById("oldLocation").value = location;
  document.getElementById("transferRemark").value = "";
  document.getElementById("trHouseName").value = "";
  document.getElementById("trHouseId").value = "";
  document.getElementById("trHouseCount").value = "";
  if (document.getElementById("farmSection")) document.getElementById("farmSection").style.display = "block";
  loadFarmSites().then(sites => {
    const sel = document.getElementById("trSiteSelect");
    sel.innerHTML = '<option value="">-- เลือกฟาร์ม --</option>';
    sites.forEach(s => {
      const o = document.createElement("option");
      o.value = s.siteId;
      o.textContent = `${s.siteName} (${s.farmType})`;
      sel.appendChild(o);
    });
  });
  openModal("transferModal");
}

function onActionChange() {
  const action = document.getElementById("transferAction").value;
  const fs = document.getElementById("farmSection");
  if (action.includes("ซ่อม") || action.includes("คืนคลัง")) {
    if (fs) fs.style.display = "none";
    document.getElementById("transferStatus").value = action.includes("ซ่อม") ? "ส่งซ่อม" : "สำรอง";
  } else {
    if (fs) fs.style.display = "block";
  }
}

function onFarmTypeChange() {
  const ft = document.getElementById("trFarmType").value;
  const at = document.getElementById("trAnimalType");
  at.innerHTML = '<option value="">-- ไม่ระบุ --</option>';
  if (ft === "สัตว์ปีก") {
    ["ไก่เนื้อ", "ไก่ไข่", "เป็ด", "ห่าน", "อื่นๆ"].forEach(v => { const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      at.appendChild(o); });
  } else if (ft === "สัตว์บก") {
    ["วัว", "ควาย", "แพะ", "แกะ", "ม้า", "อื่นๆ"].forEach(v => { const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      at.appendChild(o); });
  } else if (ft === "สุกร") {
    ["สุกรพ่อพันธุ์", "สุกรแม่พันธุ์", "ลูกสุกร", "สุกรขุน"].forEach(v => { const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      at.appendChild(o); });
  }
}

async function onSiteSelect(siteId) {
  if (!siteId) {
    document.getElementById("transferSite").value = "";
    document.getElementById("trHouseSelect").innerHTML = '<option value="">-- เลือกไซต์ก่อน --</option>';
    return;
  }
  const site = farmSites.find(s => s.siteId === siteId);
  if (site) {
    document.getElementById("transferSite").value = site.siteName;
    try {
      const r = await fetch(`/api/farm-houses/${encodeURIComponent(siteId)}`);
      const houses = await r.json();
      const sel = document.getElementById("trHouseSelect");
      sel.innerHTML = '<option value="">-- เลือกโรงเรือน --</option>';
      houses.forEach(h => {
        const o = document.createElement("option");
        o.value = h.houseId;
        o.dataset.name = h.houseName;
        o.textContent = `${h.houseName} (${h.houseType})`;
        sel.appendChild(o);
      });
      if (site.farmType) document.getElementById("trFarmType").value = site.farmType;
    } catch (e) {}
  }
}

function onHouseSelect(houseId) {
  const sel = document.getElementById("trHouseSelect");
  const opt = sel.options[sel.selectedIndex];
  if (opt && opt.dataset.name) { document.getElementById("trHouseName").value = opt.dataset.name;
    document.getElementById("trHouseId").value = houseId; }
}

async function submitTransferAction() {
  const serial = document.getElementById("transferSerialTitle").textContent;
  const action = document.getElementById("transferAction").value;
  const status = document.getElementById("transferStatus").value;
  const location = document.getElementById("transferLocation").value.trim();
  const siteName = document.getElementById("transferSite").value.trim();
  const user = document.getElementById("transferUser").value.trim();
  const remark = document.getElementById("transferRemark").value.trim();
  const from = `${document.getElementById("oldSiteName").value} (${document.getElementById("oldLocation").value})`;
  const farmType = document.getElementById("trFarmType")?.value || "-";
  const animalType = document.getElementById("trAnimalType")?.value || "-";
  const houseId = document.getElementById("trHouseId")?.value || "-";
  const houseName = document.getElementById("trHouseName")?.value.trim() || "-";
  if (!siteName) { alert("กรุณาระบุไซต์งานปลายทาง"); return; }
  try {
    const r = await fetch("/api/transfer-asset", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ serialNumber: serial, action, status, location, siteName, user, remark, fromLocation: from, farmType, animalType, houseId, houseName })
    });
    const d = await r.json();
    if (d.success) { closeModal("transferModal");
      await loadAssets(); if (currentMonitorFarm) renderMonitorView(); } else alert("เกิดข้อผิดพลาด: " + d.error);
  } catch (e) { alert("ไม่สามารถเชื่อมต่อได้"); }
}

// ---- History Modal ----
function viewAssetHistory(serial) {
  if (!serial || serial === "ไม่มี Serial") { alert("ไม่พบ Serial Number"); return; }
  document.getElementById("modalSerialTitle").textContent = serial;
  document.getElementById("hmCounter").textContent = "";
  document.getElementById("hmExtLink").href = `/trace.html?serial=${encodeURIComponent(serial)}&from=internal`;
  document.getElementById("historyModalBody").innerHTML = '<div style="text-align:center;padding:40px 0"><div class="spin"></div><p style="color:var(--tmuted);margin-top:12px;font-size:13px">กำลังโหลด...</p></div>';
  openModal("historyModal");
  try {
    fetch(`/api/asset-history/${encodeURIComponent(serial)}`)
      .then(res => res.json())
      .then(data => {
        document.getElementById("hmCounter").textContent = `${data.length} รายการ`;
        if (!data.length) { document.getElementById("historyModalBody").innerHTML = '<div style="text-align:center;padding:40px;color:var(--tmuted);font-size:14px">📭 ยังไม่มีประวัติในระบบ</div>'; return; }
        let html = '<div class="hmt">';
        data.forEach((e, i) => {
          const f = i === 0;
          const { cls, icon } = f ? { cls: "hmd-cur", icon: "📍" } : _dotClass(e.action);
          const init = getInitials(e.user || "");
          const rH = (e.from && e.from !== "-") ? `<div class="hm-route"><span class="hm-chip">${e.from}</span><span class="hm-arr">→</span><span class="hm-chip hm-chip-d">${e.to || "-"}</span></div>` : (e.to && e.to !== "-") ? `<div class="hm-route"><span class="hm-chip hm-chip-d">📍 ${e.to}</span></div>` : "";
          const rem = (e.remark && e.remark !== "-") ? `<div class="hm-rem"><span class="hm-rem-lbl">หมายเหตุ</span>${e.remark}</div>` : "";
          html += `<div class="hme"><div class="hmd ${cls}">${icon}</div><div class="hmc ${f ? "hmc-a" : ""}"><div class="hmc-top"><div class="hm-action">${e.action || "-"}${f ? '<span class="hm-tag">ล่าสุด</span>' : ""}</div><div class="hm-date">📅 ${e.date || "-"}</div></div>${rH}<div class="hm-ur"><div class="hm-av">${init}</div><span style="color:var(--tmuted);font-size:12px">${e.user || "System"}</span></div>${rem}</div></div>`;
        });
        html += "</div>";
        document.getElementById("historyModalBody").innerHTML = html;
      });
  } catch (e) { document.getElementById("historyModalBody").innerHTML = '<div style="text-align:center;padding:30px;color:var(--red)">⚠️ ไม่สามารถโหลดประวัติได้</div>'; }
}

function _dotClass(a = "") {
  a = a.toLowerCase();
  if (a.includes("ลงทะเบียน") || a.includes("เพิ่ม")) return { cls: "hmd-reg", icon: "✦" };
  if (a.includes("ซ่อม")) return { cls: "hmd-rep", icon: "🔧" };
  if (a.includes("คืน")) return { cls: "hmd-ret", icon: "↩" };
  return { cls: "hmd-move", icon: "→" };
}

// ---- Part Catalog & Bulk Add ----
async function loadPartCatalogForSelect() {
  console.log('🔄 loadPartCatalogForSelect เริ่มทำงาน');
  
  if (!assetData || !assetData.length) {
    await loadAssets();
  }

  // ลองอ่านจาก Cache
  const cached = sessionStorage.getItem('partCatalogList');
  if (cached) {
    try {
      const cacheTime = sessionStorage.getItem('partCatalogListTime');
      if (cacheTime && (Date.now() - parseInt(cacheTime)) < 300000) {
        const data = JSON.parse(cached);
        partCatalogList = data;
        partCatalog = data;
        console.log('📦 โหลดจาก Cache:', data.length, 'รายการ');
        populatePartSelect();
        generateAssetId();
        generateSerial();
        await loadSiteDropdown();
        return;
      }
    } catch (e) {
      console.warn('Cache error:', e);
    }
  }

  try {
    console.log('🔄 กำลังโหลดจาก API...');
    const res = await fetch('/api/part-catalog');
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    console.log('✅ โหลดจาก API สำเร็จ:', data.length, 'รายการ');
    
    partCatalogList = data;
    partCatalog = data;
    sessionStorage.setItem('partCatalogList', JSON.stringify(data));
    sessionStorage.setItem('partCatalogListTime', String(Date.now()));
    
    populatePartSelect();
    generateAssetId();
    generateSerial();
    await loadSiteDropdown();
  } catch (e) {
    console.error('❌ Load part catalog error:', e);
    partCatalogList = [];
    partCatalog = [];
    populatePartSelect(); // เรียกเพื่อแสดงข้อความ error
  }
}

function generateAssetId() {
  if (!assetData || !assetData.length) { document.getElementById('newAssetIdAuto').value = '0001'; return; }
  const maxId = assetData.map(a => parseInt(a.assetId) || 0).reduce((max, curr) => curr > max ? curr : max, 0);
  document.getElementById('newAssetIdAuto').value = String(maxId + 1).padStart(4, '0');
}

function generateSerial() {
  const partNumber = document.getElementById('newPartNumber').value.trim().toUpperCase();
  if (!partNumber) { document.getElementById('newAssetSerialAuto').value = ''; return; }
  const now = new Date();
  const ds = String(now.getDate()).padStart(2, '0') + String(now.getMonth() + 1).padStart(2, '0') + now.getFullYear();
  const prefix = `SN-${partNumber}-${ds}`;
  let maxNum = 0;
  (assetData || []).forEach(a => {
    const s = a.serialNumber || '';
    const parts = s.split('-');
    if (parts.length >= 3) {
      if (parts[parts.length - 2] === ds) {
        const num = parseInt(parts[parts.length - 1]);
        if (!isNaN(num) && num > maxNum) maxNum = num;
      }
    }
  });
  document.getElementById('newAssetSerialAuto').value = `${prefix}-${String(maxNum + 1).padStart(4, '0')}`;
}

async function loadSiteDropdown() {
  try {
    const res = await fetch('/api/farm-sites');
    if (!res.ok) { console.warn('Farm sites API error:', res.status); return; }
    const data = await res.json();
    const sites = Array.isArray(data) ? data : [];
    const sel = document.getElementById('newAssetSite');
    if (!sel) return;
    sel.innerHTML = '<option value="Intranin">🏢 บริษัท Intranin (คลังกลาง)</option>';
    sites.forEach(s => {
      if (s.siteName && s.siteName !== 'Intranin') {
        const opt = document.createElement('option');
        opt.value = s.siteName;
        opt.textContent = s.siteName;
        sel.appendChild(opt);
      }
    });
    sel.value = 'Intranin';
  } catch (e) { console.error('Load sites error:', e); }
}

async function addAssetNewPart() {
  const partNumber = document.getElementById('newPartNumber').value.trim().toUpperCase();
  const partName = document.getElementById('newPartName').value.trim();
  const category = document.getElementById('newPartCategory').value.trim();
  const unit = document.getElementById('newPartUnit').value;
  const description = document.getElementById('newPartDescription').value.trim();
  let assetId = document.getElementById('newAssetIdAuto').value;
  let serialNumber = document.getElementById('newAssetSerialAuto').value.trim();
  if (!partNumber || !partName) { alert('กรุณากรอก Part Number และชื่ออุปกรณ์'); return; }
  if (!serialNumber) {
    generateSerial();
    serialNumber = document.getElementById('newAssetSerialAuto').value.trim();
  }
  if (!serialNumber) { alert('❌ ไม่สามารถสร้าง Serial Number ได้ กรุณาตรวจสอบ Part Number'); return; }
  const exists = partCatalogList.some(p => p.partNumber === partNumber);
  if (exists) { alert(`⚠️ Part Number "${partNumber}" มีอยู่แล้วในระบบ\nกรุณาใช้ "📦 เพิ่มหลายชิ้น" เพื่อเพิ่ม Asset ภายใต้ Part นี้`); return; }
  const partRes = await fetch('/api/add-part', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ partNumber, partName, category: category || '', description: description || '', unit: unit || 'ชิ้น' })
  });
  const partData = await partRes.json();
  if (!partData.success) { alert('❌ ไม่สามารถเพิ่ม Part ได้: ' + (partData.error || '')); return; }
  const body = {
    assetId: assetId,
    name: partName,
    code: partNumber,
    partNumber: partNumber,
    serialNumber: serialNumber,
    status: document.getElementById('newAssetStatus').value,
    location: document.getElementById('newAssetLocation').value.trim() || '-',
    siteName: document.getElementById('newAssetSite').value,
    user: document.getElementById('newAssetUser').value.trim() || ''
  };
  const assetRes = await fetch('/api/add-asset', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  const assetDataResult = await assetRes.json();
  if (assetDataResult.success) {
    alert(`✅ เพิ่ม Part + Asset สำเร็จ\nPart: ${partNumber}\nAsset ID: ${assetId}\nSerial: ${serialNumber}`);
    closeModal('addAssetModal');
    ['newPartNumber', 'newPartName', 'newPartCategory', 'newPartDescription', 'newAssetLocation', 'newAssetSite', 'newAssetUser'].forEach(id => document.getElementById(id).value = '');
    document.getElementById('newAssetSerialAuto').value = '';
    await loadAssets();
    await loadPartCatalog();
    await loadPartCatalogForSelect();
  } else {
    alert('❌ ไม่สามารถเพิ่ม Asset ได้: ' + (assetDataResult.error || ''));
  }
}

async function loadPartCatalog() {
  // ✅ ตรวจสอบ Cache ใน sessionStorage ก่อน
  const cached = sessionStorage.getItem('partCatalogList');
  if (cached) {
    try {
      const cacheTime = sessionStorage.getItem('partCatalogListTime');
      if (cacheTime && (Date.now() - parseInt(cacheTime)) < 300000) { // 5 นาที
        const data = JSON.parse(cached);
        // ✅ ตั้งค่าทั้งสองตัวแปร
        partCatalog = data;
        partCatalogList = data;
        populatePartSelect();
        console.log('✅ Part catalog loaded from browser cache (loadPartCatalog)');
        return;
      }
    } catch (e) {
      console.warn('Browser cache read error:', e);
    }
  }

  // ถ้าไม่มี Cache หรือหมดอายุ ให้โหลดจาก API
  try {
    const r = await fetch("/api/part-catalog");
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const data = await r.json();
    // ✅ ตั้งค่าทั้งสองตัวแปร
    partCatalog = data;
    partCatalogList = data;
    // ✅ บันทึก Cache
    sessionStorage.setItem('partCatalogList', JSON.stringify(data));
    sessionStorage.setItem('partCatalogListTime', String(Date.now()));
    populatePartSelect();
    console.log('✅ Part catalog loaded from API');
  } catch (e) {
    console.error("Load part catalog error:", e);
    partCatalog = [];
    partCatalogList = [];
    toast('⚠️ ไม่สามารถโหลดข้อมูล Part Catalog ได้', 'err');
  }
}

function clearPartCatalogCache() {
  sessionStorage.removeItem('partCatalogList');
  sessionStorage.removeItem('partCatalogListTime');
  console.log('🗑️ Part catalog cache cleared');
}

function populatePartSelect() {
  console.log('🔍 populatePartSelect ถูกเรียก, partCatalog.length =', partCatalog.length);
  
  const sel = document.getElementById('bulkPartSelect');
  if (!sel) {
    console.warn('❌ ไม่พบ element bulkPartSelect');
    return;
  }
  
  sel.innerHTML = '<option value="">-- เลือก Part --</option>';
  
  if (partCatalog && partCatalog.length > 0) {
    partCatalog.forEach(p => {
      const opt = document.createElement('option');
      opt.value = p.partNumber;
      opt.textContent = `${p.partNumber} - ${p.partName} (${p.totalQty || 0} ชิ้น)`;
      sel.appendChild(opt);
    });
    console.log('✅ เติม dropdown สำเร็จ:', partCatalog.length, 'รายการ');
  } else {
    console.warn('⚠️ partCatalog ว่าง');
    sel.innerHTML += '<option value="" disabled style="color:red;">⚠️ ไม่มีข้อมูล Part</option>';
  }
}

function onBulkPartChange(pn) {
  if (!pn) { document.getElementById("bulkPartNumber").value = "";
    document.getElementById("bulkPartName").value = ""; return; }
  const part = partCatalog.find(p => p.partNumber === pn);
  if (part) { document.getElementById("bulkPartNumber").value = part.partNumber;
    document.getElementById("bulkPartName").value = part.partName; }
}

function previewBulkSerial() {
  const pn = (document.getElementById("bulkPartNumber").value || "").trim().toUpperCase();
  const qty = parseInt(document.getElementById("bulkQty").value) || 0;
  if (!pn || !qty) { alert("กรอก Part Number และจำนวนก่อน"); return; }
  const now = new Date();
  const ds = String(now.getDate()).padStart(2, "0") + String(now.getMonth() + 1).padStart(2, "0") + now.getFullYear();
  const prefix = `SN-${pn}-${ds}`;
  let maxNum = 0;
  (assetData || []).forEach(a => {
    const s = a.serialNumber || '';
    const parts = s.split('-');
    if (parts.length >= 3) {
      if (parts[parts.length - 2] === ds) {
        const num = parseInt(parts[parts.length - 1]);
        if (!isNaN(num) && num > maxNum) maxNum = num;
      }
    }
  });
  const lines = [`Serial ที่จะสร้าง (${qty} ชิ้น):`];
  for (let i = 1; i <= Math.min(qty, 5); i++) {
    const runNum = String(maxNum + i).padStart(4, "0");
    lines.push(`${prefix}-${runNum}`);
  }
  if (qty > 5) lines.push(`... และอีก ${qty - 5} ชิ้น`);
  const prev = document.getElementById("bulkPreview");
  prev.style.display = "block";
  prev.innerHTML = lines.join("<br>");
}

// ฟังก์ชันเปิด Bulk Add Modal (โหลดข้อมูลก่อน แล้วค่อยเปิด)
async function openBulkAddModal() {
  console.log('🔄 กำลังโหลด Part Catalog...');
  await loadPartCatalogForSelect(); // รอให้โหลดเสร็จ
  console.log('✅ โหลดเสร็จ เปิด Modal');
  openModal('bulkAddModal');
}

async function confirmBulkAdd() {
  const pn = (document.getElementById("bulkPartNumber").value || "").trim().toUpperCase();
  const pname = document.getElementById("bulkPartName").value.trim();
  const qty = parseInt(document.getElementById("bulkQty").value);
  if (!pn || !pname || !qty) { alert("กรอกข้อมูลให้ครบ"); return; }
  const body = {
    partNumber: pn,
    partName: pname,
    qty,
    status: document.getElementById("bulkStatus").value,
    siteName: document.getElementById("bulkSite").value.trim(),
    location: document.getElementById("bulkLocation").value.trim(),
    user: document.getElementById("bulkUser").value.trim()
  };
  try {
    const r = await fetch("/api/bulk-add-asset", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
    const d = await r.json();
    if (d.success) {
      alert(`✅ เพิ่ม ${d.added} ชิ้นสำเร็จ\nSerial: ${d.firstSerial} ~ ${d.lastSerial}`);
      closeModal("bulkAddModal");
      await loadAssets();
      await loadPartCatalog();
    } else alert("เกิดข้อผิดพลาด: " + (d.error || ""));
  } catch (e) { alert("ไม่สามารถเชื่อมต่อได้"); }
}