/**
 * public/js/bundle.js — INTRANIN EMS Bundle System (Frontend)
 * ─────────────────────────────────────────────────────────────
 * เรียก REST API ที่ routes/bundle.js จัดการ
 * วางไฟล์นี้ที่: public/js/bundle.js
 */

(function () {
  'use strict';

  /* ─────────────────────────────────────────
     STATE
  ───────────────────────────────────────── */
  var _all       = [];   // bundle list ทั้งหมดจาก API
  var _currentId = null; // bundle ที่กำลังดู detail

  /* ─────────────────────────────────────────
     API HELPERS
  ───────────────────────────────────────── */
  function _get(url) {
    return fetch(url, { credentials: 'same-origin' }).then(function (r) {
      if (!r.ok) return r.json().then(function (e) { throw new Error(e.error || r.statusText); });
      return r.json();
    });
  }

  function _post(url, body) {
    return fetch(url, {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body || {})
    }).then(function (r) {
      return r.json().then(function (d) {
        if (!d.success && !r.ok) throw new Error(d.error || r.statusText);
        return d;
      });
    });
  }

  function _patch(url, body) {
    return fetch(url, {
      method: 'PATCH',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body || {})
    }).then(function (r) { return r.json(); });
  }

  function _delete(url) {
    return fetch(url, {
      method: 'DELETE',
      credentials: 'same-origin'
    }).then(function (r) { return r.json(); });
  }

  function _toast(msg, type) {
    if (window.EMS && EMS.toast) { EMS.toast(msg, type || 'info'); return; }
    if (type === 'err' || type === 'error') { alert('❌ ' + msg); return; }
    alert('✅ ' + msg);
  }

  function _esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function _val(id) {
    var el = document.getElementById(id);
    return el ? el.value.trim() : '';
  }

  function _setText(id, v) {
    var el = document.getElementById(id);
    if (el) el.textContent = v;
  }

  function _shake(id) {
    var el = document.getElementById(id);
    if (!el) return;
    el.style.border = '1.5px solid var(--red)';
    el.style.animation = 'shake .3s ease';
    setTimeout(function () { el.style.border = ''; el.style.animation = ''; }, 700);
    el.focus();
  }

  /* ─────────────────────────────────────────
     LOAD BUNDLES
  ───────────────────────────────────────── */
  function loadBundles() {
    _setListLoading(true);
    return _get('/api/bundles')
      .then(function (data) {
        _all = data || [];
        _updateStats();
        _loadFarmOptions();
        filterBundles();
      })
      .catch(function (err) {
        _toast('โหลด Bundle ล้มเหลว: ' + err.message, 'err');
        console.error(err);
      })
      .finally(function () { _setListLoading(false); });
  }

  /* ─────────────────────────────────────────
     STATS
  ───────────────────────────────────────── */
  function _updateStats() {
    var total    = _all.length;
    var inStock  = _all.filter(function (b) { return b.status === 'In Stock'; }).length;
    var deployed = _all.filter(function (b) { return b.status === 'Deployed'; }).length;
    var assets   = _all.reduce(function (n, b) { return n + (b.assetIds ? b.assetIds.length : 0); }, 0);
    _setText('bdl-stat-total',    total);
    _setText('bdl-stat-stock',    inStock);
    _setText('bdl-stat-deployed', deployed);
    _setText('bdl-stat-assets',   assets);
  }

  /* ─────────────────────────────────────────
     FARM DROPDOWN (ดึงจาก API ฟาร์มที่มีอยู่แล้ว)
  ───────────────────────────────────────── */
  function _loadFarmOptions() {
    // ลองดึงจาก /api/farms ก่อน ถ้าไม่มีก็สร้างจาก bundle data
    _get('/api/farms')
      .then(function (farms) { _populateFarmSelects(farms); })
      .catch(function () {
        // fallback: สร้างจาก deployed bundles
        var seen = {};
        var list = [];
        _all.forEach(function (b) {
          if (b.farmId && !seen[b.farmId]) {
            seen[b.farmId] = true;
            list.push({ farmId: b.farmId, farmName: b.location || b.farmId });
          }
        });
        _populateFarmSelects(list);
      });
  }

  function _populateFarmSelects(farms) {
    // farms อาจเป็น array of object {farmId, farmName} หรือ array of array จาก sheet
    var opts = farms.map(function (f) {
      var id   = f.farmId   || f[0] || '';
      var name = f.farmName || f[1] || id;
      return '<option value="' + _esc(id) + '">' + _esc(name) + '</option>';
    }).join('');

    var filterSel = document.getElementById('bdl-filter-farm');
    var deploySel = document.getElementById('bdl-deploy-farm');
    if (filterSel) filterSel.innerHTML = '<option value="">ทุกฟาร์ม</option>' + opts;
    if (deploySel) deploySel.innerHTML = '<option value="">— เลือกฟาร์ม —</option>' + opts;
  }

  /* ─────────────────────────────────────────
     FILTER + RENDER LIST
  ───────────────────────────────────────── */
  function filterBundles() {
    var q  = (_val('bdl-search') || '').toLowerCase();
    var st = _val('bdl-filter-status') || '';
    var fm = _val('bdl-filter-farm')   || '';

    var filtered = _all.filter(function (b) {
      var matchQ  = !q  || (b.bundleId || '').toLowerCase().includes(q)
                        || (b.bundleName || '').toLowerCase().includes(q)
                        || (b.description || '').toLowerCase().includes(q);
      var matchSt = !st || b.status === st;
      var matchFm = !fm || b.farmId === fm;
      return matchQ && matchSt && matchFm;
    });

    _renderGrid(filtered);
  }

  function _renderGrid(list) {
    var grid  = document.getElementById('bdl-grid');
    var empty = document.getElementById('bdl-empty');
    if (!grid) return;

    if (!list.length) {
      grid.style.display = 'none';
      if (empty) empty.style.display = 'flex';
      return;
    }
    if (empty) empty.style.display = 'none';
    grid.style.display = 'grid';
    grid.innerHTML = list.map(_cardHTML).join('');
  }

  function _cardHTML(b) {
    var cnt = b.assetIds ? b.assetIds.length : 0;
    var sc  = { 'In Stock': 'bdl-s-stock', 'Deployed': 'bdl-s-dep', 'Maintenance': 'bdl-s-maint' }[b.status] || '';
    var loc = b.status === 'Deployed' ? (b.location || b.farmId || '—') : 'Stock';

    var actionBtn = b.status === 'In Stock'
      ? '<button class="btn btn-blue btn-sm" onclick="BDL.openDeploy(\'' + _esc(b.bundleId) + '\')" style="flex:1">'
        + '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><polygon points="1 6 1 22 8 18 16 22 23 18 23 2 16 6 8 2 1 6"/></svg> ส่งฟาร์ม</button>'
      : '<button class="btn btn-out btn-sm" onclick="BDL.recall(\'' + _esc(b.bundleId) + '\')" style="flex:1">'
        + '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg> คืน Stock</button>';

    return '<div class="bdl-card">'
      + '<div class="bdl-card-head">'
      +   '<div class="bdl-card-id">' + _esc(b.bundleId) + '</div>'
      +   '<span class="bdl-status ' + sc + '">' + _esc(b.status) + '</span>'
      + '</div>'
      + '<div class="bdl-card-name">' + _esc(b.bundleName) + '</div>'
      + (b.description ? '<div class="bdl-card-desc">' + _esc(b.description) + '</div>' : '')
      + '<div class="bdl-card-meta">'
      +   '<span>📦 ' + cnt + ' อุปกรณ์</span>'
      +   '<span>📍 ' + _esc(loc) + '</span>'
      + '</div>'
      + '<div class="bdl-card-actions">'
      +   '<button class="btn btn-out btn-sm" onclick="BDL.openDetail(\'' + _esc(b.bundleId) + '\')" style="flex:1">🔍 รายละเอียด</button>'
      +   actionBtn
      +   '<button class="btn btn-out btn-sm" onclick="BDL.openEdit(\'' + _esc(b.bundleId) + '\')" title="แก้ไข" style="padding:0 10px">'
      +     '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>'
      +   '</button>'
      + '</div>'
      + '</div>';
  }

  /* ─────────────────────────────────────────
     DETAIL VIEW
  ───────────────────────────────────────── */
  function openDetail(bundleId) {
    var b = _find(bundleId);
    if (!b) return;
    _currentId = bundleId;

    document.getElementById('bdl-list-wrap').style.display = 'none';
    document.getElementById('bdl-stats').style.display     = 'none';
    document.getElementById('bdl-detail').style.display    = 'block';

    // Render skeleton ก่อน
    _renderDetailShell(b, null);

    // โหลด asset details
    var ids = b.assetIds || [];
    if (!ids.length) {
      _renderDetailShell(b, []);
      return;
    }

    // ดึงรายละเอียด asset ทั้งชุดพร้อมกัน (ไม่กรอง "อยู่ใน Bundle แล้ว")
    _get('/api/bundles/asset-info?ids=' + ids.map(encodeURIComponent).join(','))
      .then(function (assets) { _renderDetailShell(b, assets); })
      .catch(function () { _renderDetailShell(b, []); });
  }

  function _renderDetailShell(b, assets) {
    var loc = b.status === 'Deployed' ? (b.location || b.farmId) : 'Stock';
    var sc  = { 'In Stock': 'bdl-s-stock', 'Deployed': 'bdl-s-dep', 'Maintenance': 'bdl-s-maint' }[b.status] || '';

    var actionBtn = b.status === 'In Stock'
      ? '<button class="btn btn-blue" onclick="BDL.openDeploy(\'' + _esc(b.bundleId) + '\')">📤 ส่งไปฟาร์ม</button>'
      : '<button class="btn btn-out" onclick="BDL.recall(\'' + _esc(b.bundleId) + '\')">↩ คืน Stock</button>';

    var assetHTML = '';
    if (assets === null) {
      // skeleton
      assetHTML = [1, 2, 3].map(function () {
        return '<div style="display:flex;gap:10px;padding:14px 16px;border-bottom:1px solid var(--g100)">'
          + '<div class="ems-sk" style="width:90px;height:18px;border-radius:6px"></div>'
          + '<div class="ems-sk" style="flex:1;height:14px;border-radius:6px"></div>'
          + '<div class="ems-sk" style="width:64px;height:22px;border-radius:20px"></div>'
          + '</div>';
      }).join('');
    } else if (!assets.length) {
      assetHTML = '<div class="ems-empty" style="padding:28px 16px">'
        + '<div class="ems-empty-ico"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><polyline points="22 12 16 12 14 15 10 15 8 12 2 12"/><path d="M5.45 5.11L2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/></svg></div>'
        + '<h3>ยังไม่มีอุปกรณ์ในชุดนี้</h3><p>กด "เพิ่มอุปกรณ์" เพื่อเพิ่มเข้าชุด</p></div>';
    } else {
      assetHTML = assets.map(function (a) {
        var stColor = a.status === 'ใช้งานได้' ? 'var(--emerald-d)'
                    : a.status === 'ซ่อมแซม'  ? 'var(--amber)' : 'var(--g500)';
        var ser = a.serial || a.assetId;
        var enc = encodeURIComponent(ser);
        var actBtns =
          '<a href="/trace.html?serial=' + enc + '&from=bundle" target="_blank" title="ดูประวัติ" class="btn btn-out btn-sm" style="padding:2px 8px">📜</a>'
          + '<a href="/qr.html?serial=' + enc + '" target="_blank" title="QR Code" class="btn btn-teal btn-sm" style="padding:2px 8px">📷</a>'
          + '<button onclick="openTransferModal(\'' + _esc(ser).replace(/'/g, "\\'") + '\',\'' + _esc(a.status || '').replace(/'/g, "\\'") + '\',\'' + _esc(a.location || '').replace(/'/g, "\\'") + '\',\'' + _esc(a.site || '').replace(/'/g, "\\'") + '\',\'' + _esc(a.user || '').replace(/'/g, "\\'") + '\')" title="จัดการ/โอนย้าย" class="btn btn-blue btn-sm" style="padding:2px 8px">🚚</button>';
        return '<div class="bdl-asset-row">'
          + '<div class="bdl-asset-serial">' + _esc(a.serial || a.assetId) + '</div>'
          + '<div class="bdl-asset-info">'
          +   '<div class="bdl-asset-name">' + _esc(a.name) + '</div>'
          +   '<div class="bdl-asset-code">' + _esc(a.assetId) + (a.code && a.code !== a.assetId ? ' · ' + _esc(a.code) : '') + '</div>'
          + '</div>'
          + '<span class="bdl-asset-status" style="color:' + stColor + ';background:' + stColor + '18;border-color:' + stColor + '30">'
          +   _esc(a.status || '—')
          + '</span>'
          + '<div style="display:flex;gap:6px;flex-shrink:0">'
          +   actBtns
          +   '<button onclick="BDL.removeAsset(\'' + _esc(b.bundleId) + '\',\'' + _esc(a.assetId) + '\')" '
          +     'title="ลบออกจากชุด" style="background:none;border:none;cursor:pointer;color:var(--red);padding:4px 8px;border-radius:6px">'
          +     '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>'
          +   '</button>'
          + '</div>'
          + '</div>';
      }).join('');
    }

    var el = document.getElementById('bdl-detail-content');
    if (!el) return;
    el.innerHTML =
      // ── Header Card ──
      '<div class="panel" style="margin-bottom:14px"><div style="padding:20px 22px">'
      + '<div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px;flex-wrap:wrap">'
      +   '<div>'
      +     '<div style="font-size:10px;font-weight:800;letter-spacing:1px;color:var(--tmuted);text-transform:uppercase;margin-bottom:6px;font-family:monospace">' + _esc(b.bundleId) + '</div>'
      +     '<div style="font-size:21px;font-weight:800;color:var(--text);line-height:1.2;margin-bottom:5px">' + _esc(b.bundleName) + '</div>'
      +     (b.description ? '<div style="font-size:13px;color:var(--tsub)">' + _esc(b.description) + '</div>' : '')
      +   '</div>'
      +   '<span class="bdl-status ' + sc + '" style="font-size:12px;padding:5px 12px">' + _esc(b.status) + '</span>'
      + '</div>'
      + '<div style="display:flex;gap:18px;margin-top:14px;padding-top:14px;border-top:1px solid var(--g100);flex-wrap:wrap">'
      +   '<div style="font-size:12px"><b style="color:var(--g600);display:block;margin-bottom:2px">ตำแหน่ง</b>' + _esc(loc) + '</div>'
      +   '<div style="font-size:12px"><b style="color:var(--g600);display:block;margin-bottom:2px">อุปกรณ์</b>' + (b.assetIds ? b.assetIds.length : 0) + ' ชิ้น</div>'
      +   '<div style="font-size:12px"><b style="color:var(--g600);display:block;margin-bottom:2px">สร้างโดย</b>' + _esc(b.createdBy || '—') + '</div>'
      +   '<div style="font-size:12px"><b style="color:var(--g600);display:block;margin-bottom:2px">อัพเดท</b>' + _esc(b.updatedDate || '—') + '</div>'
      + '</div>'
      + '<div style="display:flex;gap:8px;margin-top:14px;flex-wrap:wrap">'
      +   '<button class="btn btn-out btn-sm" onclick="BDL.openEdit(\'' + _esc(b.bundleId) + '\')">✏️ แก้ไข</button>'
      +   actionBtn
      + '</div>'
      + '</div></div>'
      // ── Asset List Card ──
      + '<div class="panel"><div class="ph" style="justify-content:space-between">'
      +   '<div style="font-size:14px;font-weight:700">📦 รายการอุปกรณ์ในชุด</div>'
      +   '<button class="btn btn-blue btn-sm" onclick="BDL.openAddAsset(\'' + _esc(b.bundleId) + '\')">+ เพิ่มอุปกรณ์</button>'
      + '</div>'
      + '<div id="bdl-asset-list">' + assetHTML + '</div>'
      + '</div>';
  }

  function backToList() {
    _currentId = null;
    document.getElementById('bdl-list-wrap').style.display = '';
    document.getElementById('bdl-stats').style.display     = '';
    document.getElementById('bdl-detail').style.display    = 'none';
  }

  /* ─────────────────────────────────────────
     CREATE / EDIT
  ───────────────────────────────────────── */
  function openCreate() {
    document.getElementById('bdlModalTitle').textContent = 'สร้าง Bundle ใหม่';
    document.getElementById('bdl-edit-id').value         = '';
    document.getElementById('bdl-inp-id').value          = '';
    document.getElementById('bdl-inp-id').disabled       = false;
    document.getElementById('bdl-inp-name').value        = '';
    document.getElementById('bdl-inp-desc').value        = '';
    document.getElementById('bdl-inp-status').value      = 'In Stock';
    if (typeof openModal === 'function') openModal('bdlCreateModal');
  }

  function openEdit(bundleId) {
    var b = _find(bundleId);
    if (!b) return;
    document.getElementById('bdlModalTitle').textContent = 'แก้ไข Bundle';
    document.getElementById('bdl-edit-id').value         = b.bundleId;
    document.getElementById('bdl-inp-id').value          = b.bundleId;
    document.getElementById('bdl-inp-id').disabled       = true;
    document.getElementById('bdl-inp-name').value        = b.bundleName;
    document.getElementById('bdl-inp-desc').value        = b.description || '';
    document.getElementById('bdl-inp-status').value      = b.status;
    if (typeof openModal === 'function') openModal('bdlCreateModal');
  }

  function saveBundle() {
    var editId = _val('bdl-edit-id');
    var id     = _val('bdl-inp-id').toUpperCase();
    var name   = _val('bdl-inp-name');
    var desc   = _val('bdl-inp-desc');
    var status = _val('bdl-inp-status');

    if (!id)   { _shake('bdl-inp-id');   return; }
    if (!name) { _shake('bdl-inp-name'); return; }

    var p = editId
      ? _patch('/api/bundles/' + encodeURIComponent(editId), { bundleName: name, description: desc, status: status })
      : _post('/api/bundles', { bundleId: id, bundleName: name, description: desc });

    if (typeof closeModal === 'function') closeModal('bdlCreateModal');

    p.then(function (res) {
      if (res.success === false) { _toast(res.error || 'บันทึกล้มเหลว', 'err'); return; }
      _toast(editId ? 'อัพเดท Bundle สำเร็จ' : 'สร้าง Bundle สำเร็จ', 'ok');
      loadBundles();
    }).catch(function (err) { _toast(err.message, 'err'); });
  }

  /* ─────────────────────────────────────────
     DEPLOY
  ───────────────────────────────────────── */
  function openDeploy(bundleId) {
    var b = _find(bundleId);
    if (!b) return;
    _currentId = bundleId;
    _setText('bdl-deploy-name',  b.bundleName + ' (' + b.bundleId + ')');
    _setText('bdl-deploy-count', (b.assetIds ? b.assetIds.length : 0) + ' อุปกรณ์จะถูกย้ายพร้อมกัน');
    document.getElementById('bdl-deploy-farm').value  = '';
    document.getElementById('bdl-deploy-note').value  = '';
    if (typeof openModal === 'function') openModal('bdlDeployModal');
  }

  function confirmDeploy() {
    var sel      = document.getElementById('bdl-deploy-farm');
    var farmId   = sel ? sel.value : '';
    var farmName = sel && sel.selectedIndex >= 0 ? sel.options[sel.selectedIndex].text : farmId;
    var note     = _val('bdl-deploy-note');

    if (!farmId) { _shake('bdl-deploy-farm'); return; }

    if (typeof closeModal === 'function') closeModal('bdlDeployModal');

    _post('/api/bundles/' + encodeURIComponent(_currentId) + '/deploy', {
      farmId: farmId, farmName: farmName, note: note
    }).then(function (res) {
      if (res.success === false) { _toast(res.error || 'ย้ายล้มเหลว', 'err'); return; }
      _toast(res.message || 'Deploy สำเร็จ', 'ok');
      loadBundles().then(function () { if (_currentId) openDetail(_currentId); });
    }).catch(function (err) { _toast(err.message, 'err'); });
  }

  /* ─────────────────────────────────────────
     RECALL
  ───────────────────────────────────────── */
  function recallBundle(bundleId) {
    var b = _find(bundleId);
    if (!b) return;
    if (!confirm('คืน "' + b.bundleName + '" ทั้งชุดกลับ Stock?\n(' + (b.assetIds ? b.assetIds.length : 0) + ' อุปกรณ์)')) return;

    _post('/api/bundles/' + encodeURIComponent(bundleId) + '/recall')
      .then(function (res) {
        if (res.success === false) { _toast(res.error || 'คืนล้มเหลว', 'err'); return; }
        _toast(res.message || 'คืน Stock สำเร็จ', 'ok');
        loadBundles().then(function () { if (_currentId === bundleId) openDetail(bundleId); });
      }).catch(function (err) { _toast(err.message, 'err'); });
  }

  /* ─────────────────────────────────────────
     ADD / REMOVE ASSET
  ───────────────────────────────────────── */
  function openAddAsset(bundleId) {
  _currentId     = bundleId;
  _pendingAssets = []; // reset ทุกครั้งที่เปิด modal
  var inp = document.getElementById('bdl-add-asset-search');
  var res = document.getElementById('bdl-asset-results');
  var pending = document.getElementById('bdl-pending-list');
  if (inp) inp.value = '';
  if (res) res.innerHTML = '<div style="padding:20px;text-align:center;color:var(--tmuted);font-size:13px">พิมพ์เพื่อค้นหา Asset</div>';
  if (pending) pending.innerHTML = '';
  _updatePendingBar();
  if (typeof openModal === 'function') openModal('bdlAddAssetModal');
}

  // debounce search
  var _searchTimer = null;
  var _pendingAssets = []; // รายการที่รอเพิ่ม
  function searchAssets(q) {
    clearTimeout(_searchTimer);
    _searchTimer = setTimeout(function () { _doSearch(q); }, 300);
  }

  function _doSearch(q) {
    var res = document.getElementById('bdl-asset-results');
    if (!res) return;
    if (!q || q.length < 2) {
      res.innerHTML = '<div style="padding:20px;text-align:center;color:var(--tmuted);font-size:13px">พิมพ์อย่างน้อย 2 ตัวอักษร</div>';
      return;
    }
    res.innerHTML = '<div style="padding:16px;text-align:center;color:var(--tmuted);font-size:13px">กำลังค้นหา...</div>';

    _get('/api/bundles/search-assets?q=' + encodeURIComponent(q))
      .then(function (items) {
        if (!items.length) {
          res.innerHTML = '<div style="padding:20px;text-align:center;color:var(--tmuted);font-size:13px">ไม่พบ Asset ที่ตรงกัน</div>';
          return;
        }
        // กรองออก asset ที่อยู่ในชุดนี้แล้ว
        var b   = _find(_currentId);
        var cur = b && b.assetIds ? b.assetIds : [];
        res.innerHTML = items.map(function (a) {
          var inBundle = cur.indexOf(a.assetId) !== -1;
          return '<div class="bdl-search-row" '
            + (inBundle ? 'style="opacity:.45;cursor:not-allowed"'
                        : 'onclick="BDL.addAsset(\'' + _esc(_currentId) + '\',\'' + _esc(a.assetId) + '\')"')
            + '>'
            + '<div style="font-size:13px;font-weight:600">' + _esc(a.assetId) + ' — ' + _esc(a.name) + (inBundle ? ' <span style="font-size:10px;color:var(--blue)">(อยู่ในชุดแล้ว)</span>' : '') + '</div>'
            + '<div style="font-size:11px;color:var(--tmuted)">S/N: ' + _esc(a.serial || '—') + ' · ' + _esc(a.status || '—') + ' · ' + _esc(a.location || '—') + '</div>'
            + '</div>';
        }).join('');
      })
      .catch(function () {
        res.innerHTML = '<div style="padding:16px;text-align:center;color:var(--red);font-size:13px">ค้นหาล้มเหลว</div>';
      });
  }

  function addAsset(bundleId, assetId, assetName) {
  if (_pendingAssets.indexOf(assetId) !== -1) {
    _toast(assetId + ' เลือกไว้แล้ว', 'warn'); return;
  }
  var b   = _find(bundleId);
  var cur = b && b.assetIds ? b.assetIds : [];
  if (cur.indexOf(assetId) !== -1) {
    _toast(assetId + ' อยู่ในชุดอยู่แล้ว', 'warn'); return;
  }
  _pendingAssets.push(assetId);
  _updatePendingBar();
  // re-render search results เพื่อแสดงว่าเลือกแล้ว
  var inp = document.getElementById('bdl-add-asset-search');
  if (inp && inp.value.length >= 2) _doSearch(inp.value);
}

function _updatePendingBar() {
  var bar = document.getElementById('bdl-pending-bar');
  var cnt = document.getElementById('bdl-pending-count');
  var list = document.getElementById('bdl-pending-list');
  if (!bar) return;
  if (!_pendingAssets.length) {
    bar.style.display = 'none'; return;
  }
  bar.style.display = 'block';
  if (cnt) cnt.textContent = _pendingAssets.length;
  if (list) list.innerHTML = _pendingAssets.map(function(id) {
    return '<span style="display:inline-flex;align-items:center;gap:4px;background:var(--blue-l);color:var(--blue);border:1px solid var(--blue-b);border-radius:20px;padding:3px 10px;font-size:11px;font-weight:600">'
      + _esc(id)
      + '<button onclick="BDL.removePending(\'' + _esc(id) + '\')" style="background:none;border:none;cursor:pointer;color:var(--blue);padding:0;margin-left:2px;font-size:12px;line-height:1">×</button>'
      + '</span>';
  }).join('');
}

function removePending(assetId) {
  _pendingAssets = _pendingAssets.filter(function(id) { return id !== assetId; });
  _updatePendingBar();
  var inp = document.getElementById('bdl-add-asset-search');
  if (inp && inp.value.length >= 2) _doSearch(inp.value);
}

function commitAddAssets() {
  if (!_pendingAssets.length) return;
  _post('/api/bundles/' + encodeURIComponent(_currentId) + '/assets/bulk', {
    assetIds: _pendingAssets
  }).then(function(res) {
    if (res.success === false) { _toast(res.error || 'เพิ่มล้มเหลว', 'err'); return; }
    _toast(res.message || 'เพิ่มอุปกรณ์สำเร็จ', 'ok');
    if (typeof closeModal === 'function') closeModal('bdlAddAssetModal');
    loadBundles().then(function() { openDetail(_currentId); });
  }).catch(function(err) { _toast(err.message, 'err'); });
}

  function removeAsset(bundleId, assetId) {
    if (!confirm('ลบ ' + assetId + ' ออกจากชุดนี้?')) return;
    _delete('/api/bundles/' + encodeURIComponent(bundleId) + '/assets/' + encodeURIComponent(assetId))
      .then(function (res) {
        if (res.success === false) { _toast(res.error || 'ลบล้มเหลว', 'err'); return; }
        _toast('ลบ ' + assetId + ' ออกจากชุดสำเร็จ', 'ok');
        loadBundles().then(function () { openDetail(bundleId); });
      }).catch(function (err) { _toast(err.message, 'err'); });
  }

  /* ─────────────────────────────────────────
     LOADING STATE
  ───────────────────────────────────────── */
  function _setListLoading(on) {
    var grid  = document.getElementById('bdl-grid');
    var empty = document.getElementById('bdl-empty');
    if (!grid) return;
    if (on) {
      if (empty) empty.style.display = 'none';
      grid.style.display = 'grid';
      grid.innerHTML = [1, 2, 3].map(function () {
        return '<div class="bdl-card" style="pointer-events:none">'
          + '<div class="ems-sk" style="height:14px;width:70px;border-radius:6px;margin-bottom:10px"></div>'
          + '<div class="ems-sk" style="height:19px;width:75%;border-radius:6px;margin-bottom:8px"></div>'
          + '<div class="ems-sk" style="height:12px;width:90%;border-radius:6px"></div>'
          + '</div>';
      }).join('');
    }
  }

  /* ─────────────────────────────────────────
     UTILS
  ───────────────────────────────────────── */
  function _find(id) {
    return _all.find(function (b) { return b.bundleId === id; }) || null;
  }

  /* ─────────────────────────────────────────
     PUBLIC API  →  window.BDL
  ───────────────────────────────────────── */
  window.BDL = {
    refresh:     loadBundles,
    getAll:      function () { return _all; },
    ensureLoaded: function () { return _all.length ? Promise.resolve(_all) : loadBundles(); },
    filter:      filterBundles,
    openCreate:  openCreate,
    openEdit:    openEdit,
    saveBundle:  saveBundle,
    openDetail:  openDetail,
    backToList:  backToList,
    openDeploy:  openDeploy,
    confirmDeploy: confirmDeploy,
    recall:      recallBundle,
    openAddAsset: openAddAsset,
    searchAssets: searchAssets,
    addAsset:    addAsset,
    removeAsset: removeAsset,
    removePending:   removePending,
  commitAddAssets: commitAddAssets,
  };

  /* ─────────────────────────────────────────
     AUTO-INIT: โหลด Bundle เมื่อเปิด tab
  ───────────────────────────────────────── */
  document.addEventListener('DOMContentLoaded', function () {
    var _loaded = false;
    var _orig   = window.openTab;
    if (typeof _orig === 'function') {
      window.openTab = function (tabId) {
        _orig(tabId);
        if (tabId === 'bundle' && !_loaded) {
          _loaded = true;
          loadBundles();
        }
      };
    }
  });

})();
