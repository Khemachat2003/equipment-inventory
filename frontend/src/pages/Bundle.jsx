import { useEffect, useState, useMemo, useCallback, useRef } from 'react';
import axios from 'axios';
import { useSearchParams } from 'react-router-dom';
import Icon from '../components/ui/Icon.jsx';
import StatusBadge from '../components/ui/StatusBadge.jsx';
import TransferModal from '../components/TransferModal.jsx';
import { AssetHistoryModal } from '../components/AssetHistory.jsx';
import { useBusy, BusyOverlay } from '../components/ui/Busy.jsx';
import { showToast, ToastHost } from '../components/ui/Toast.jsx';
import { showConfirm, ConfirmHost } from '../components/ui/Confirm.jsx';
import { buildLocation, buildBundleLocation } from '../utils/location.js';
import { farmKeyOf, STOCK_KEY } from '../utils/bundleGroups.js';
import { FarmInlineAdd, HouseInlineAdd } from '../components/InlineFarmAdd.jsx';
import { readDispatchQueue, writeDispatchQueue, removeDispatchedAssets } from '../utils/dispatchQueue.js';

const BTONES = {
  'In Stock': 'blue',
  Deployed: 'green',
  Maintenance: 'amber',
};

function isUnbundledStockAsset(asset) {
  if (!asset || asset.bundleId) return false;
  const location = String(asset.location || '').trim().toLowerCase();
  const site = String(asset.siteName || '').trim().toLowerCase();
  return location === 'stock' || location === 'intranin' || site === 'stock' || site === 'intranin'
    || String(asset.status || '').trim().toLowerCase() === 'in stock';
}

function suggestBundleId(existingIds = []) {
  const used = new Set(existingIds.map((id) => String(id || '').toUpperCase()));
  let next = 1;
  while (used.has(`BDL-${String(next).padStart(3, '0')}`)) next += 1;
  return `BDL-${String(next).padStart(3, '0')}`;
}

// ── ก้อน ⑥ (Issue A): จัดกลุ่มอัตโนมัติ + ตัวกรองฟาร์ม ──────────────────────────
// เกณฑ์ "ชุดนี้อยู่คลังหรือฟาร์มไหน" = farmKeyOf / STOCK_KEY จาก utils/bundleGroups.js (import ด้านบน)

export default function Bundle() {
  const [bundles, setBundles] = useState([]);
  const [farms, setFarms] = useState([]);
  const [search, setSearch] = useState('');
  const [filterStatus, setFilterStatus] = useState('');
  const [filterFarm, setFilterFarm] = useState('');
  const [detail, setDetail] = useState(null);  // bundleId ที่เปิด detail
  const [detailAssets, setDetailAssets] = useState([]);
  const [createOpen, setCreateOpen] = useState(false);
  const [deployOpen, setDeployOpen] = useState(null);
  const [addAssetOpen, setAddAssetOpen] = useState(null);
  const [pendingAssets, setPendingAssets] = useState([]);
  const [searchResults, setSearchResults] = useState([]);
  const [transfer, setTransfer] = useState(null);
  const [historySerial, setHistorySerial] = useState('');
  const [loading, setLoading] = useState(true);
  const [searchParams, setSearchParams] = useSearchParams();
  const [dispatchQueue, setDispatchQueue] = useState(() => readDispatchQueue());
  const [dispatchOpen, setDispatchOpen] = useState(false);
  const [collapsedFarms, setCollapsedFarms] = useState({}); // ก้อน ⑥: ฟาร์มไหนถูกพับอยู่ (default กางหมด)
  const busy = useBusy();

  const load = useCallback(async () => {
    try {
      const [b, f] = await Promise.all([axios.get('/api/bundles'), axios.get('/api/farms')]);
      setBundles((b.data || []));
      setFarms(f.data || []);
    } catch (e) {
      console.error('load bundles', e);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { load(); }, [load]);

  useEffect(() => {
    if (searchParams.get('dispatch') === '1') {
      setDispatchQueue(readDispatchQueue());
      setDispatchOpen(true);
      setSearchParams({}, { replace: true });
    }
  }, [searchParams, setSearchParams]);

  // เพิ่มฟาร์มใหม่จากฟอร์มย่อใน DeployModal → refresh รายการฟาร์ม (dropdown อัปเดตทันที)
  async function refreshFarms() {
    try {
      const { data } = await axios.get('/api/farms');
      setFarms(data || []);
    } catch (e) { /* ignore */ }
  }

  // Auto-open จาก Farm Monitor (?open=bundleId) + เส้นทางย้ายทั้งชุดจากหน้าแรก (?deploy=bundleId — B3 รอบ 3)
  useEffect(() => {
    if (!bundles.length) return;
    const deployId = searchParams.get('deploy');
    const openId = searchParams.get('open');
    if (deployId) {
      const b = bundles.find((x) => x.bundleId === deployId);
      if (b && b.status === 'In Stock') setDeployOpen(deployId); // ยังอยู่คลัง → เปิดฟอร์มย้ายทั้งชุด 2 ขั้นเลย
      else openDetail(deployId); // ติดตั้งที่ฟาร์มแล้ว → เข้าหน้าชุด (ย้ายรายชิ้น/คืนเข้าคลัง)
      setSearchParams({}, { replace: true });
      return;
    }
    if (openId) {
      openDetail(openId);
      setSearchParams({}, { replace: true });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [searchParams, bundles]);

  // Stats
  const stats = useMemo(() => ({
    all: bundles.length,
    inStock: bundles.filter((b) => b.status === 'In Stock').length,
    deployed: bundles.filter((b) => b.status === 'Deployed').length,
    devices: bundles.reduce((s, b) => s + (b.assetIds || []).length, 0),
  }), [bundles]);

  const filtered = useMemo(() => {
    const q = search.toLowerCase();
    return bundles.filter((b) => {
      if (q && !(b.bundleId.toLowerCase().includes(q) || b.bundleName.toLowerCase().includes(q) || (b.description || '').toLowerCase().includes(q))) return false;
      if (filterStatus && b.status !== filterStatus) return false;
      if (filterFarm) {
        const k = farmKeyOf(b);
        if (filterFarm === STOCK_KEY ? k !== '' : k !== filterFarm) return false;
      }
      return true;
    });
  }, [bundles, search, filterStatus, filterFarm]);

  // จำนวน Bundle ต่อฟาร์ม คำนวณจากคำค้นหาและสถานะที่เลือกอยู่
  const farmCounts = useMemo(() => {
    const q = search.toLowerCase();
    const base = bundles.filter((b) => {
      if (q && !(b.bundleId.toLowerCase().includes(q) || b.bundleName.toLowerCase().includes(q) || (b.description || '').toLowerCase().includes(q))) return false;
      if (filterStatus && b.status !== filterStatus) return false;
      return true;
    });
    const counts = {};
    const nameHint = {};
    let stock = 0;
    base.forEach((b) => {
      const k = farmKeyOf(b);
      if (!k) { stock += 1; return; }
      counts[k] = (counts[k] || 0) + 1;
      if (b.farmName && !nameHint[k]) nameHint[k] = b.farmName;
    });
    const th = (x, y) => x.localeCompare(y, 'th');
    const list = farms
      .filter((farm) => farm.farmId)
      .map((farm) => ({ key: farm.farmId, name: farm.farmName || farm.farmId, count: counts[farm.farmId] || 0 }));
    const knownKeys = new Set(list.map((farm) => farm.key));
    Object.keys(counts).forEach((key) => {
      if (!knownKeys.has(key)) list.push({ key, name: nameHint[key] || key, count: counts[key] });
    });
    list.sort((a, b2) => th(a.name, b2.name));
    return { list, stock, total: base.length };
  }, [bundles, farms, search, filterStatus]);

  // ── ก้อน ⑥ (opt1): จัดกลุ่ม 2 โซน — "อยู่ในคลัง" / "ติดตั้งที่ฟาร์มแล้ว" (แบ่งย่อยตามฟาร์ม เรียงชื่อก-ฮ)
  const grouped = useMemo(() => {
    const stock = [];
    const byFarm = {};
    const nameHint = {};
    filtered.forEach((b) => {
      const k = farmKeyOf(b);
      if (!k) { stock.push(b); return; }
      if (!byFarm[k]) byFarm[k] = [];
      byFarm[k].push(b);
      if (b.farmName && !nameHint[k]) nameHint[k] = b.farmName;
    });
    const th = (x, y) => x.localeCompare(y, 'th');
    const farmSections = Object.keys(byFarm)
      .map((k) => ({
        key: k,
        name: farms.find((f) => f.farmId === k)?.farmName || nameHint[k] || k,
        type: farms.find((f) => f.farmId === k)?.farmType || '',
        bundles: byFarm[k].sort((a, b2) => th(a.bundleName || '', b2.bundleName || '')),
      }))
      .sort((a, b2) => th(a.name, b2.name));
    stock.sort((a, b2) => th(a.bundleName || '', b2.bundleName || ''));
    return { stock, farmSections, deployedCount: filtered.length - stock.length };
  }, [filtered, farms]);

  const toggleFarmGroup = (k) => setCollapsedFarms((p) => ({ ...p, [k]: !p[k] }));

  // Detail
  const detailBundle = useMemo(() => bundles.find((b) => b.bundleId === detail), [bundles, detail]);

  async function openDetail(id) {
    setDetail(id);
    setDetailAssets([]);
    try {
      const bundle = bundles.find((b) => b.bundleId === id);
      const ids = (bundle?.assetIds || []).join(',');
      if (!ids) return;
      const { data } = await axios.get(`/api/bundles/asset-info?ids=${encodeURIComponent(ids)}`);
      setDetailAssets(data || []);
    } catch (e) { /* ignore */ }
  }

  async function createBundle(data) {
    await busy.run('กำลังสร้างชุดใหม่...', async () => {
      try {
        await axios.post('/api/bundles', data);
        await load();
        setCreateOpen(false);
        showToast('สร้างชุดใหม่สำเร็จ', { actionLabel: 'เปิดดูชุดนี้', onAction: () => openDetail(data.bundleId) });
      } catch (e) {
        showToast(e.response?.data?.error || 'เกิดข้อผิดพลาด', { type: 'err' });
      }
    });
  }

  async function assembleDispatchBundle(data, assetIds) {
    if (!assetIds.length) return;
    await busy.run('กำลังประกอบชุดจากคิวจัดส่ง...', async () => {
      try {
        const latest = await axios.get('/api/assets');
        const currentById = new Map((latest.data || []).map((asset) => [asset.assetId, asset]));
        const stale = assetIds.filter((assetId) => {
          const asset = currentById.get(assetId);
          return !asset || !isUnbundledStockAsset(asset) || (dispatchQueue.poNumber && asset.poNumber !== dispatchQueue.poNumber);
        });
        if (stale.length) {
          showToast(`มี Serial ที่ถูกจัดชุดหรือเปลี่ยน PO ไปแล้ว: ${stale.join(', ')} · รีเฟรชคิวก่อนสร้าง`, { type: 'warn' });
          return;
        }
        await axios.post('/api/bundles', data);
        await axios.post(`/api/bundles/${encodeURIComponent(data.bundleId)}/assets/bulk`, { assetIds });
        const nextQueue = removeDispatchedAssets(dispatchQueue, assetIds);
        writeDispatchQueue(nextQueue);
        setDispatchQueue(nextQueue);
        await load();
        if (!nextQueue) setDispatchOpen(false);
        showToast(nextQueue ? `สร้าง ${data.bundleName} แล้ว · เหลือ ${nextQueue.items.length} Serial ในคิว` : `สร้าง ${data.bundleName} สำเร็จ`, {
          actionLabel: 'กำหนดปลายทาง',
          onAction: () => {
            setDispatchOpen(false);
            setDeployOpen(data.bundleId);
          },
        });
      } catch (error) {
        showToast(error.response?.data?.error || 'สร้างชุดจากคิวไม่สำเร็จ', { type: 'err' });
      }
    });
  }

  async function removeAsset(bundleId, assetId) {
    // ยืนยันด้วยกล่องในระบบ (แทน window.confirm) — ยกเลิก/Esc/คลิกพื้นหลัง = ไม่ทำอะไร
    if (!(await showConfirm({
      title: `นำ ${assetId} ออกจากชุด?`,
      message: 'อุปกรณ์จะหลุดออกจากชุดนี้ — ข้อมูลชิ้นงานยังอยู่ครบในระบบตามเดิม',
      confirmLabel: 'นำออกจากชุด',
      danger: true,
    }))) return;
    await busy.run('กำลังนำอุปกรณ์ออกจากชุด...', async () => {
      try {
        await axios.delete(`/api/bundles/${bundleId}/assets/${assetId}`);
        await load();
        await openDetail(bundleId);
      } catch (e) {
        showToast(e.response?.data?.error || 'เกิดข้อผิดพลาด', { type: 'err' });
      }
    });
  }

  async function deploy(bundleId, farmId, farmName, note, houseId, houseName) {
    await busy.run('กำลังย้ายชุดไปฟาร์ม...', async () => {
      try {
        const { data } = await axios.post(`/api/bundles/${bundleId}/deploy`, {
          farmId, farmName, note, houseId, houseName,
        });
        await load();
        setDeployOpen(null);
        if (detail) await openDetail(bundleId);
        showToast(data.message || 'ย้ายชุดไปฟาร์มแล้ว', { actionLabel: 'เปิดดูชุดนี้', onAction: () => openDetail(bundleId) });
      } catch (e) {
        showToast(e.response?.data?.error || 'เกิดข้อผิดพลาด', { type: 'err' });
      }
    });
  }

  async function recall(bundleId) {
    // ยืนยันด้วยกล่องในระบบ (แทน window.confirm) — ยกเลิก/Esc/คลิกพื้นหลัง = ไม่ทำอะไร
    if (!(await showConfirm({
      title: 'คืนชุดนี้กลับเข้าคลัง?',
      message: 'อุปกรณ์ทุกชิ้นในชุดจะถูกคืนเข้าคลังกลางด้วย',
      confirmLabel: 'คืนเข้าคลัง',
    }))) return;
    await busy.run('กำลังคืนชุดกลับเข้าคลัง...', async () => {
      try {
        const { data } = await axios.post(`/api/bundles/${bundleId}/recall`);
        await load();
        if (detail) await openDetail(bundleId);
        showToast(data.message || 'คืนชุดเข้าคลังแล้ว', { actionLabel: 'เปิดดูชุดนี้', onAction: () => openDetail(bundleId) });
      } catch (e) {
        showToast(e.response?.data?.error || 'เกิดข้อผิดพลาด', { type: 'err' });
      }
    });
  }

  async function commitAddAssets(bundleId) {
    await busy.run('กำลังเพิ่มอุปกรณ์เข้าชุด...', async () => {
      try {
        const { data } = await axios.post(`/api/bundles/${bundleId}/assets/bulk`, { assetIds: pendingAssets });
        await load();
        await openDetail(bundleId);
        setAddAssetOpen(null);
        setPendingAssets([]);
        showToast(data.message || 'เพิ่มอุปกรณ์เข้าชุดแล้ว', { actionLabel: 'เปิดดูชุดนี้', onAction: () => openDetail(bundleId) });
      } catch (e) {
        showToast(e.response?.data?.error || 'เกิดข้อผิดพลาด', { type: 'err' });
      }
    });
  }

  async function searchAssets(q) {
    if (!q || q.length < 2) { setSearchResults([]); return; }
    try {
      const { data } = await axios.get(`/api/bundles/search-assets?q=${encodeURIComponent(q)}`);
      setSearchResults(data || []);
    } catch (e) {
      setSearchResults([]);
    }
  }

  function exportCSV() {
    const rows = [['BundleID', 'ชื่อชุด', 'คำอธิบาย', 'สถานะ', 'ตำแหน่ง', 'FarmID', 'อุปกรณ์ (AssetIDs)']];
    bundles.forEach((b) => rows.push([b.bundleId, b.bundleName, b.description, b.status, b.location, b.farmId, (b.assetIds || []).join(' | ')]));
    dlCSV(`bundle_export_${stamp()}.csv`, rows);
  }

  // ก้อน ⑥ — grid การ์ดชุด (ใช้ซ้ำทั้งโซนคลังและกลุ่มย่อยฟาร์ม)
  const renderCards = (list) => (
    <div className="grid w-full min-w-0 grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
      {list.map((b) => (
        <BundleCard
          key={b.bundleId}
          bundle={b}
          onDetail={() => openDetail(b.bundleId)}
          onDeploy={() => setDeployOpen(b.bundleId)}
          onRecall={() => recall(b.bundleId)}
        />
      ))}
    </div>
  );

  return (
    <div className="w-full min-w-0 space-y-4">
      <BusyOverlay label={busy.busyLabel} />
      <ToastHost />
      <ConfirmHost />
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div className="flex flex-wrap gap-2 w-full sm:w-auto">
          <button onClick={exportCSV} className="flex flex-1 sm:flex-none items-center justify-center gap-1.5 min-h-11 sm:min-h-0 px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[14px] sm:text-[13px] hover:bg-[var(--surface2)] whitespace-nowrap">
            <Icon name="download" size="sm" /> CSV
          </button>
          <button onClick={() => setCreateOpen(true)} className="flex flex-1 sm:flex-none items-center justify-center gap-1.5 min-h-11 sm:min-h-0 px-3 py-1.5 rounded-lg bg-[var(--blue)] text-white text-[14px] sm:text-[13px] font-semibold hover:bg-[var(--blue-d)] whitespace-nowrap">
            <Icon name="add" size="sm" /> สร้างชุดใหม่ (Bundle)
          </button>
        </div>
      </div>

      {dispatchQueue?.items?.length > 0 && (
        <section className="flex flex-col gap-3 rounded-xl border border-[var(--blue-b)] bg-[var(--blue-l)]/50 p-3 sm:flex-row sm:items-center sm:justify-between sm:p-4">
          <div className="min-w-0"><div className="text-[13px] font-bold text-[var(--blue)]">คิวจัดชุดจาก PO {dispatchQueue.poNumber}</div><div className="text-[12px] text-[var(--tsub)]">เหลือ {dispatchQueue.items.length} Serial · แบ่งลง Bundle ได้หลายชุด</div></div>
          <div className="flex gap-2"><button onClick={() => { writeDispatchQueue(null); setDispatchQueue(null); setDispatchOpen(false); }} className="min-h-11 rounded-lg border border-[var(--g300)] bg-[var(--surface)] px-3 text-[12px] font-semibold text-[var(--tsub)]">ล้างคิว</button><button onClick={() => setDispatchOpen(true)} className="min-h-11 rounded-lg bg-[var(--blue)] px-4 text-[12px] font-semibold text-white">จัดชุดต่อ</button></div>
        </section>
      )}

      {/* Stats */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-2 sm:gap-3">
        <StatCard label="ชุดทั้งหมด" value={stats.all} icon="folder_open" tone="blue" />
        <StatCard label="อยู่ในคลัง (In Stock)" value={stats.inStock} icon="inventory_2" tone="green" />
        <StatCard label={<>ติดตั้งที่ฟาร์มแล้ว <span className="hidden sm:inline">(Deployed)</span></>} value={stats.deployed} icon="local_shipping" tone="amber" />
        <StatCard label="อุปกรณ์ในชุด" value={stats.devices} icon="devices_other" tone="red" />
      </div>

      {/* Filters — มือถือ: ช่องค้นหา/ตัวกรองกางเต็มกว้างเรียงลงเป็นชั้น (sm ขึ้นไป: เรียงแถวเดียวพับได้) */}
      <div className="grid grid-cols-1 gap-2 sm:flex sm:flex-wrap sm:items-center sm:gap-3">
        <div className="relative sm:max-w-xs sm:flex-1">
          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[var(--tmuted)]"><Icon name="search" size="sm" /></span>
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="ค้นหาชุดอุปกรณ์..."
            className="w-full h-11 sm:h-9 pl-9 pr-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[14px] sm:text-[13px]" />
        </div>
        <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)} className="w-full sm:w-auto h-11 sm:h-9 px-3 rounded-lg border border-[var(--g200)] text-[14px] sm:text-[13px]">
          <option value="">ทุกสถานะ</option>
          <option value="In Stock">อยู่ในคลัง (In Stock)</option>
          <option value="Deployed">ติดตั้งที่ฟาร์มแล้ว (Deployed)</option>
          <option value="Maintenance">ซ่อมบำรุง (Maintenance)</option>
        </select>
        </div>

      <FarmFilterCombobox value={filterFarm} onChange={setFilterFarm} counts={farmCounts} />

      {/* Detail view or list */}
      {detailBundle ? (
        <BundleDetail
          bundle={detailBundle}
          assets={detailAssets}
          onBack={() => { setDetail(null); setDetailAssets([]); }}
          onRefresh={() => openDetail(detail)}
          onAdd={() => { setPendingAssets([]); setAddAssetOpen(detail); }}
          onRemove={(assetId) => removeAsset(detail, assetId)}
          onDeploy={() => setDeployOpen(detail)}
          onRecall={() => recall(detail)}
          onTransfer={(a) => setTransfer({
            serial: a.serial || a.assetId,
            current: { status: a.status, location: a.location, siteName: a.site, user: a.user },
          })}
          onHistory={(a) => { if (a.serial) setHistorySerial(a.serial); }}
        />
      ) : (
        <div className="space-y-5">
          {loading && <div className="text-center py-10 text-[var(--tmuted)]">กำลังโหลด...</div>}
          {!loading && filtered.length === 0 && (
            <div className="text-center py-10 text-[var(--tmuted)]">ไม่พบ Bundle</div>
          )}

          {/* 📦 โซนคลัง — ชุดที่ยังไม่ได้ติดตั้ง (In Stock / ไม่ผูกฟาร์ม) */}
          {!loading && grouped.stock.length > 0 && (
            <section className="space-y-3">
              <ZoneHeader icon="inventory_2" tone="blue" title="อยู่ในคลัง" count={grouped.stock.length} />
              {renderCards(grouped.stock)}
            </section>
          )}

          {/* 📍 โซนฟาร์ม — แบ่งย่อยตามฟาร์ม (กดหัวกลุ่มเพื่อพับ/กาง) */}
          {!loading && grouped.farmSections.length > 0 && (
            <section className="space-y-3">
              <ZoneHeader icon="place" tone="green" title="ติดตั้งที่ฟาร์มแล้ว" count={grouped.deployedCount} />
              {grouped.farmSections.map((sec) => (
                <div key={sec.key} className="space-y-2.5">
                  <button
                    onClick={() => toggleFarmGroup(sec.key)}
                    className="flex w-full items-center gap-1.5 text-left"
                    title={collapsedFarms[sec.key] ? 'กางดูชุดของฟาร์มนี้' : 'พับชุดของฟาร์มนี้'}
                  >
                    <Icon name={collapsedFarms[sec.key] ? 'expand_more' : 'expand_less'} size="sm" className="text-[var(--tmuted)]" />
                    <Icon name="agriculture" size="sm" className="text-[var(--emerald-d)]" />
                    <span className="text-[13px] font-bold text-[var(--text)]">{sec.name}{sec.type ? ` (${sec.type})` : ''}</span>
                    <span className="text-[12px] font-semibold text-[var(--tmuted)]">— {sec.bundles.length} ชุด</span>
                  </button>
                  {/* กำลังกรองอยู่ที่ฟาร์มนี้ → บังคับกางเสมอ (กันหน้าจอว่างทั้งหน้า) */}
                  {filterFarm !== sec.key && collapsedFarms[sec.key] ? null : renderCards(sec.bundles)}
                </div>
              ))}
            </section>
          )}
        </div>
      )}

      {/* Create modal */}
      {createOpen && <CreateModal onClose={() => setCreateOpen(false)} onSubmit={createBundle} existingIds={bundles.map((b) => b.bundleId)} />}
      {dispatchOpen && dispatchQueue?.items?.length > 0 && <DispatchAssemblyModal key={dispatchQueue.items.map((item) => item.assetId).join('|')} queue={dispatchQueue} existingIds={bundles.map((b) => b.bundleId)} onClose={() => setDispatchOpen(false)} onSubmit={assembleDispatchBundle} />}

      {/* Deploy modal */}
      {deployOpen && (
        <DeployModal
          bundle={bundles.find((b) => b.bundleId === deployOpen)}
          farms={farms}
          onClose={() => setDeployOpen(null)}
          onDeploy={(farmId, farmName, note, houseId, houseName) => deploy(deployOpen, farmId, farmName, note, houseId, houseName)}
          onFarmAdded={refreshFarms}
        />
      )}

      {/* Add asset modal */}
      {addAssetOpen && (
        <AddAssetModal
          bundleId={addAssetOpen}
          pending={pendingAssets}
          setPending={setPendingAssets}
          results={searchResults}
          onSearch={searchAssets}
          onClose={() => { setAddAssetOpen(null); setPendingAssets([]); }}
          onCommit={() => commitAddAssets(addAssetOpen)}
        />
      )}

      {/* Transfer */}
      {transfer && (
        <TransferModal
          open
          onClose={() => setTransfer(null)}
          onSuccess={() => { load(); if (detail) openDetail(detail); }}
          serial={transfer.serial}
          current={transfer.current}
        />
      )}
      {historySerial && <AssetHistoryModal key={historySerial} serial={historySerial} onClose={() => setHistorySerial('')} />}
    </div>
  );
}

function FarmFilterCombobox({ value, onChange, counts }) {
  const rootRef = useRef(null);
  const triggerRef = useRef(null);
  const inputRef = useRef(null);
  const optionRefs = useRef([]);
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState('');
  const [activeIndex, setActiveIndex] = useState(0);

  const options = useMemo(() => [
    { key: '', name: 'ทุกฟาร์ม', count: counts.total, icon: 'grid_on' },
    ...counts.list.map((farm) => ({ ...farm, icon: 'agriculture' })),
    { key: STOCK_KEY, name: 'คลัง', count: counts.stock, icon: 'inventory_2' },
  ], [counts]);
  const normalizedQuery = query.trim().toLocaleLowerCase();
  const visibleOptions = options.filter((option) => (
    !normalizedQuery
    || option.name.toLocaleLowerCase().includes(normalizedQuery)
    || option.key.toLocaleLowerCase().includes(normalizedQuery)
  ));
  const currentActiveIndex = Math.min(activeIndex, visibleOptions.length - 1);
  const selected = options.find((option) => option.key === value) || options[0];
  const listboxId = 'bundle-farm-listbox';

  useEffect(() => {
    if (!open) return undefined;
    inputRef.current?.focus();
    const onPointerDown = (event) => {
      if (!rootRef.current?.contains(event.target)) setOpen(false);
    };
    document.addEventListener('pointerdown', onPointerDown);
    return () => document.removeEventListener('pointerdown', onPointerDown);
  }, [open]);

  useEffect(() => {
    if (currentActiveIndex >= 0) optionRefs.current[currentActiveIndex]?.scrollIntoView?.({ block: 'nearest' });
  }, [currentActiveIndex]);

  function showOptions() {
    setQuery('');
    setActiveIndex(Math.max(0, options.findIndex((option) => option.key === value)));
    setOpen(true);
  }

  function selectOption(option) {
    onChange(option.key);
    setOpen(false);
    setQuery('');
    triggerRef.current?.focus();
  }

  function handleSearchKeyDown(event) {
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      setActiveIndex(Math.min(currentActiveIndex + 1, visibleOptions.length - 1));
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      setActiveIndex(Math.max(currentActiveIndex - 1, 0));
    } else if (event.key === 'Enter') {
      event.preventDefault();
      if (visibleOptions[currentActiveIndex]) selectOption(visibleOptions[currentActiveIndex]);
    } else if (event.key === 'Escape') {
      event.preventDefault();
      setOpen(false);
      triggerRef.current?.focus();
    } else if (event.key === 'Tab') {
      setOpen(false);
    }
  }

  return (
    <div ref={rootRef} className="relative w-full sm:w-72">
      <div className="flex gap-2">
        <button
          ref={triggerRef}
          type="button"
          aria-label={`เลือกฟาร์ม, ค่าปัจจุบัน ${selected.name}`}
          aria-haspopup="listbox"
          aria-expanded={open}
          aria-controls={listboxId}
          onClick={() => (open ? setOpen(false) : showOptions())}
          className="flex min-w-0 flex-1 items-center gap-2 rounded-lg border border-[var(--g200)] bg-[var(--surface)] px-3 min-h-11 text-left text-[14px] text-[var(--text)] hover:border-[var(--blue)] focus:outline-none focus:ring-2 focus:ring-[var(--blue)]/20"
        >
          <Icon name={selected.icon} size="sm" className="shrink-0 text-[var(--blue)]" />
          <span className="min-w-0 flex-1 truncate">{selected.name}</span>
          <span className="shrink-0 text-[12px] text-[var(--tmuted)]">{selected.count} ชุด</span>
          <Icon name={open ? 'expand_less' : 'expand_more'} size="sm" className="shrink-0 text-[var(--tmuted)]" />
        </button>
        {value && (
          <button
            type="button"
            onClick={() => { onChange(''); setOpen(false); }}
            aria-label="ล้างตัวกรองฟาร์ม"
            title="ล้างตัวกรองฟาร์ม"
            className="flex h-11 w-11 shrink-0 items-center justify-center rounded-lg border border-[var(--g200)] bg-[var(--surface)] text-[var(--tsub)] hover:border-[var(--blue)] hover:text-[var(--blue)] focus:outline-none focus:ring-2 focus:ring-[var(--blue)]/20"
          >
            <Icon name="close" size="sm" />
          </button>
        )}
      </div>

      {open && (
        <div className="absolute left-0 right-0 top-full z-40 mt-2 overflow-hidden rounded-xl border border-[var(--g200)] bg-[var(--surface)] shadow-[var(--sh-md)]">
          <label htmlFor="bundle-farm-search" className="sr-only">ค้นหาชื่อฟาร์มหรือ Farm ID</label>
          <div className="relative border-b border-[var(--g100)] p-2">
            <Icon name="search" size="sm" className="absolute left-5 top-1/2 -translate-y-1/2 text-[var(--tmuted)]" />
            <input
              ref={inputRef}
              id="bundle-farm-search"
              role="combobox"
              aria-autocomplete="list"
              aria-expanded={open}
              aria-controls={listboxId}
              aria-activedescendant={currentActiveIndex >= 0 && visibleOptions[currentActiveIndex] ? `${listboxId}-option-${currentActiveIndex}` : undefined}
              value={query}
              onChange={(event) => { setQuery(event.target.value); setActiveIndex(0); }}
              onKeyDown={handleSearchKeyDown}
              placeholder="ค้นหาชื่อฟาร์มหรือ Farm ID..."
              autoComplete="off"
              className="h-11 w-full rounded-lg border border-[var(--g200)] bg-[var(--surface2)] pl-10 pr-3 text-[14px] text-[var(--text)] outline-none focus:border-[var(--blue)]"
            />
          </div>
          <div id={listboxId} role="listbox" aria-label="ฟาร์ม" className="max-h-[55vh] overflow-y-auto p-1.5 sm:max-h-80">
            {visibleOptions.map((option, index) => {
              const isSelected = option.key === value;
              const isActive = index === currentActiveIndex;
              return (
                <div
                  key={option.key || 'all-farms'}
                  ref={(node) => { optionRefs.current[index] = node; }}
                  id={`${listboxId}-option-${index}`}
                  role="option"
                  aria-selected={isSelected}
                  onMouseEnter={() => setActiveIndex(index)}
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => selectOption(option)}
                  className={`flex min-h-11 cursor-pointer items-center gap-2 rounded-lg px-3 py-2 text-[14px] ${
                    isActive ? 'bg-[var(--blue-l)]' : 'hover:bg-[var(--surface2)]'
                  } ${isSelected ? 'font-semibold text-[var(--blue)]' : 'text-[var(--text)]'}`}
                >
                  <Icon name={option.icon} size="sm" className="shrink-0 text-[var(--tsub)]" />
                  <span className="min-w-0 flex-1 truncate">{option.name}</span>
                  <span className="shrink-0 text-[12px] text-[var(--tmuted)]">{option.count} ชุด</span>
                  {isSelected && <Icon name="check" size="sm" className="shrink-0" />}
                </div>
              );
            })}
            {visibleOptions.length === 0 && (
              <div className="px-3 py-6 text-center text-[13px] text-[var(--tmuted)]">ไม่พบฟาร์มที่ค้นหา</div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

function StatCard({ label, value, icon, tone }) {
  const tones = {
    blue: 'bg-[var(--blue-l)] text-[var(--blue)]',
    green: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]',
    amber: 'bg-[var(--amber-l)] text-[var(--amber-d)]',
    red: 'bg-[var(--red-l)] text-[var(--red)]',
  };
  return (
    <div className="rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] p-3 sm:p-4 flex items-center gap-2.5 sm:gap-3 min-w-0">
      <span className={`flex items-center justify-center w-9 h-9 sm:w-10 sm:h-10 rounded-xl flex-shrink-0 ${tones[tone]}`}><Icon name={icon} size="md" /></span>
      <div className="min-w-0">
        <div className="text-[12px] text-[var(--tmuted)] leading-tight">{label}</div>
        <div className="text-lg sm:text-xl font-bold text-[var(--text)]">{value}</div>
      </div>
    </div>
  );
}

// ก้อน ⑥ — หัวโซนใหญ่ (คลัง / ติดตั้งที่ฟาร์มแล้ว)
function ZoneHeader({ icon, tone, title, count }) {
  const tones = {
    blue: 'bg-[var(--blue-l)] text-[var(--blue)]',
    green: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]',
  };
  return (
    <div className="flex items-center gap-2.5 pt-1">
      <span className={`flex items-center justify-center w-8 h-8 rounded-lg ${tones[tone]}`}>
        <Icon name={icon} size="sm" />
      </span>
      <span className="text-[14px] font-bold text-[var(--text)]">{title}</span>
      <span className="text-[13px] font-semibold text-[var(--tmuted)]">— {count} ชุด</span>
    </div>
  );
}

function BundleCard({ bundle, onDetail, onDeploy, onRecall }) {
  const b = bundle;
  const inStock = b.status === 'In Stock';
  // Step 2: โชว์ชื่อฟาร์มที่อ่านง่ายก่อน (farmName) + โรงเรือนถ้ามี — ตรงกับหน้า Asset/Scan ที่ใช้ buildLocation
  const locParts = inStock
    ? []
    : [
        b.farmName || b.location || b.farmId,
        b.houseName ? [b.houseId, b.houseName].filter(Boolean).join(' ') : '',
      ].filter(Boolean);
  const loc = inStock ? 'คลังกลาง' : locParts.join(' › ');
  return (
    <div className="rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] p-4 flex flex-col gap-3">
      <div className="flex items-start justify-between">
        <div>
          <div className="font-mono text-[13px] font-bold text-[var(--text)]">{b.bundleId}</div>
          <div className="text-[14px] sm:text-[13px] font-semibold text-[var(--text)]">{b.bundleName}</div>
          {b.description && <div className="text-[11px] text-[var(--tmuted)] mt-0.5 line-clamp-2">{b.description}</div>}
        </div>
        <StatusBadge status={b.status} tone={BTONES[b.status]} />
      </div>
      <div className="flex items-center gap-3 text-[12px] text-[var(--tsub)]">
        <span className="flex items-center gap-1 flex-shrink-0"><Icon name="devices" size="xs" /> {b.assetIds?.length || 0} อุปกรณ์</span>
        <span className="flex items-center gap-1 min-w-0">
          <Icon name="place" size="xs" className="flex-shrink-0" />
          <span className="truncate" title={loc}>
            {loc}
          </span>
        </span>
        {!inStock && (
          <button onClick={onDetail} className="ml-auto flex-shrink-0 text-[11px] font-semibold text-[var(--blue)] hover:underline whitespace-nowrap">
            ดูตำแหน่งเต็ม
          </button>
        )}
      </div>
      <div className="flex flex-col sm:flex-row gap-2 pt-2 border-t border-[var(--g100)]">
        <button onClick={onDetail} className="flex items-center justify-center gap-1 w-full sm:flex-1 sm:w-auto min-h-11 sm:min-h-0 px-3 py-2 sm:py-1.5 rounded-lg border border-[var(--g300)] text-[13px] sm:text-[12px] text-[var(--tsub)] hover:bg-[var(--surface2)] whitespace-nowrap"><Icon name="search" size="xs" /> รายละเอียด</button>
        {b.status === 'In Stock'
          ? <button onClick={onDeploy} className="w-full sm:w-auto min-h-11 sm:min-h-0 px-3 py-2 sm:py-1.5 rounded-lg bg-[var(--blue)] text-white text-[13px] sm:text-[12px] font-semibold whitespace-nowrap">ย้ายไปฟาร์ม (Deploy)</button>
          : <button onClick={onRecall} className="w-full sm:w-auto min-h-11 sm:min-h-0 px-3 py-2 sm:py-1.5 rounded-lg bg-[var(--amber)] text-white text-[13px] sm:text-[12px] font-semibold whitespace-nowrap">คืนเข้าคลัง (Recall)</button>}
      </div>
    </div>
  );
}

function BundleDetail({ bundle: b, assets, onBack, onRefresh, onAdd, onRemove, onDeploy, onRecall, onTransfer, onHistory }) {
  const inStock = b.status === 'In Stock';
  // ตำแหน่งของชุด = ตำแหน่งของสมาชิก (ย้ายทั้งชุดพร้อมกัน จึงใช้ค่าจากสมาชิกได้เลย)
  // ถ้าสมาชิกอยู่คนละจุด แปลว่ามีการย้ายเฉพาะรายชิ้นหลังจากนั้น — ต้องเตือนให้เห็นชัด
  const loc = useMemo(
    () => (inStock ? buildLocation({ location: 'Stock' }) : buildBundleLocation(assets).main),
    [assets, inStock],
  );
  const divergence = useMemo(
    () => (inStock || assets.length < 2 ? null : buildBundleLocation(assets)),
    [assets, inStock],
  );
  const divergent = divergence?.divergent;
  return (
    <div className="rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
      <div className="p-5 border-b border-[var(--g100)] bg-gradient-to-r from-[var(--blue-l)] to-transparent">
        <button onClick={onBack} className="flex items-center gap-1 text-[12px] text-[var(--blue)] font-semibold mb-3"><Icon name="arrow_back" size="sm" /> กลับ</button>
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <div>
              <div className="flex items-center gap-2">
                <span className="font-mono text-[16px] font-bold text-[var(--text)]">{b.bundleId}</span>
                <StatusBadge status={b.status} tone={BTONES[b.status]} />
              </div>
              <div className="text-[14px] font-semibold text-[var(--text)] mt-1">{b.bundleName}</div>
              {b.description && <div className="text-[12px] text-[var(--tmuted)]">{b.description}</div>}
            </div>
          </div>
          <div className="flex w-full sm:w-auto flex-wrap items-center gap-2">
            {b.status === 'In Stock'
              ? <button onClick={onDeploy} className="flex-1 sm:flex-none min-h-11 sm:min-h-0 inline-flex items-center justify-center gap-1.5 whitespace-nowrap px-3.5 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] sm:text-[12px] font-semibold"><Icon name="local_shipping" size="sm" /> ย้ายไปฟาร์ม (Deploy)</button>
              : <>
                <button onClick={onDeploy} className="flex-1 sm:flex-none min-h-11 sm:min-h-0 inline-flex items-center justify-center gap-1.5 whitespace-nowrap px-3.5 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] sm:text-[12px] font-semibold"><Icon name="local_shipping" size="sm" /> ย้ายทั้งชุดไปฟาร์ม</button>
                <button onClick={onRecall} className="flex-1 sm:flex-none min-h-11 sm:min-h-0 inline-flex items-center justify-center gap-1.5 whitespace-nowrap px-3.5 py-2 rounded-lg bg-[var(--amber)] text-white text-[13px] sm:text-[12px] font-semibold">คืนเข้าคลัง (Recall)</button>
              </>}
            <button onClick={onAdd} className="flex-1 sm:flex-none min-h-11 sm:min-h-0 inline-flex items-center justify-center gap-1.5 whitespace-nowrap px-3.5 py-2 rounded-lg bg-[var(--emerald)] text-white text-[13px] sm:text-[12px] font-semibold"><Icon name="add" size="sm" /> เพิ่มอุปกรณ์</button>
          </div>
        </div>
        <div className="flex flex-wrap gap-4 mt-3 text-[12px] text-[var(--tsub)]">
          <span className="flex items-center gap-1.5 min-w-0">
            <Icon name="place" size="xs" /> ตำแหน่ง:
            <strong className="font-semibold text-[var(--text)] truncate">{loc.full}</strong>
          </span>
          {loc.chips.map((c) => (
            <span key={c.key} className="flex items-center gap-1">
              <span className="rounded bg-[var(--surface2)] px-1.5 py-px text-[10px] font-medium text-[var(--tsub)] border border-[var(--g100)]">{c.label}</span>
              <strong className="font-semibold text-[var(--text)]">{c.value}</strong>
            </span>
          ))}
          <span className="flex items-center gap-1"><Icon name="folder_open" size="xs" /> อุปกรณ์: <strong>{b.assetIds?.length || 0} ชิ้น</strong></span>
          <span className="flex items-center gap-1"><Icon name="person" size="xs" /> สร้างโดย: <strong>{b.createdBy || '-'}</strong></span>
          <span className="flex items-center gap-1"><Icon name="sync" size="xs" /> อัพเดท: <strong>{b.updatedDate || '-'}</strong></span>
        </div>
        {divergent && (
          <div className="mt-3 flex items-start gap-2 px-3 py-2.5 rounded-xl bg-[var(--amber-l)] border border-[var(--amber-b)] text-[11px] text-[var(--amber-d)]">
            <Icon name="warning" size="sm" />
            <div>
              <div className="font-bold">อุปกรณ์ในชุดนี้ไม่ได้อยู่ที่เดียวกัน</div>
              <div className="mt-1 space-y-0.5">
                {divergence.paths.map((p) => <div key={p}>• {p}</div>)}
              </div>
              <div className="mt-1.5">ย้ายทั้งชุดอีกครั้งเพื่อให้อยู่ตำแหน่งเดียวกัน</div>
            </div>
          </div>
        )}
      </div>

      <div className="p-5">
        <div className="flex items-center justify-between mb-3">
          <div className="text-[13px] font-bold text-[var(--text)]">รายการอุปกรณ์ในชุด ({assets.length})</div>
          <Icon name="refresh" size="sm" className="cursor-pointer text-[var(--tmuted)] hover:text-[var(--text)]" onClick={onRefresh} />
        </div>
        {assets.length === 0 && (
          <div className="text-center py-8 text-[12px] text-[var(--tmuted)]">ยังไม่มีอุปกรณ์ในชุดนี้ — กด 'เพิ่มอุปกรณ์' เพื่อเพิ่มเข้าชุด</div>
        )}
        <div className="space-y-2">
          {assets.map((a) => (
            <div key={a.assetId} className="flex flex-wrap items-center gap-x-3 gap-y-2 px-3 py-3 sm:py-2.5 rounded-xl border border-[var(--g100)] hover:bg-[var(--surface2)]">
              <div className="flex-1 min-w-0 sm:min-w-[150px]">
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="font-mono text-[12px] font-semibold text-[var(--text)]">{a.serial || a.assetId}</span>
                  <StatusBadge status={a.status} />
                </div>
                <div className="text-[12px] text-[var(--tsub)] truncate">{a.name} <span className="text-[var(--tmuted)]">· {a.assetId} · {a.code}</span></div>
                <div className="mt-1 text-[11px] text-[var(--tsub)]">{a.batchId ? `ล็อต ${a.batchId} · รับเข้า ${formatInboundDate(a.receivedAt)}` : 'ไม่ระบุล็อต'}</div>
              </div>
              <div className="flex w-full sm:w-auto items-center justify-end sm:justify-start gap-1">
                <button onClick={() => onHistory(a)} disabled={!a.serial} title="ประวัติ" className="h-11 w-11 sm:h-9 sm:w-9 flex items-center justify-center rounded-lg text-[var(--blue)] hover:bg-[var(--blue-l)] disabled:opacity-40"><Icon name="description" size="sm" /></button>
                <a href={`/qr?serial=${encodeURIComponent(a.serial)}`} target="_blank" title="QR" className="h-11 w-11 sm:h-9 sm:w-9 flex items-center justify-center rounded-lg text-[var(--tsub)] hover:bg-[var(--surface2)]"><Icon name="qr_code" size="sm" /></a>
                <button onClick={() => onTransfer(a)} title="โอนย้าย" className="h-11 w-11 sm:h-9 sm:w-9 flex items-center justify-center rounded-lg text-[var(--blue)] hover:bg-[var(--blue-l)]"><Icon name="local_shipping" size="sm" /></button>
                <span className="w-px h-6 bg-[var(--g200)] mx-1" aria-hidden="true" />
                <button
                  onClick={() => onRemove(a.assetId)}
                  title="ถอดออกจากชุด"
                  className="h-11 sm:h-9 px-2.5 flex items-center gap-1 rounded-lg bg-[var(--red-l)] text-[13px] sm:text-[12px] text-[var(--red)] font-semibold hover:bg-[var(--red)] hover:text-white"
                >
                  <Icon name="close" size="xs" /> ถอดออก
                </button>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function DispatchAssemblyModal({ queue, existingIds, onClose, onSubmit }) {
  const [bundleId, setBundleId] = useState(() => suggestBundleId(existingIds));
  const [bundleName, setBundleName] = useState('');
  const [description, setDescription] = useState(`PO ${queue.poNumber || 'ไม่ระบุ PO'}`);
  const [selectedIds, setSelectedIds] = useState(() => queue.items.map((item) => item.assetId));
  const duplicate = existingIds.some((id) => String(id).trim().toUpperCase() === bundleId.trim().toUpperCase());
  const selectedSet = new Set(selectedIds);
  function toggle(id) { setSelectedIds((current) => current.includes(id) ? current.filter((value) => value !== id) : [...current, id]); }
  return (
    <Modal onClose={onClose} title="ประกอบ Bundle จากคิว PO">
      <div className="rounded-lg border border-[var(--blue-b)] bg-[var(--blue-l)] px-3 py-2 text-[12px] text-[var(--blue)]">PO {queue.poNumber || 'ไม่ระบุ PO'} · เลือก Serial สำหรับ Bundle นี้ ({selectedIds.length}/{queue.items.length})</div>
      <Field label="รหัส Bundle *"><input value={bundleId} onChange={(event) => setBundleId(event.target.value.toUpperCase())} className={inp} /></Field>
      {duplicate && <div className="text-[12px] font-semibold text-[var(--red)]">รหัส Bundle นี้มีอยู่แล้ว</div>}
      <Field label="ชื่อ Bundle *"><input value={bundleName} onChange={(event) => setBundleName(event.target.value)} placeholder="เช่น ตู้ควบคุมพัดลม · โรงเรือน 1" className={inp} /></Field>
      <Field label="คำอธิบาย"><input value={description} onChange={(event) => setDescription(event.target.value)} className={inp} /></Field>
      <div className="flex items-center justify-between gap-2"><span className="text-[12px] font-semibold text-[var(--tsub)]">Serial ในคิว</span><button onClick={() => setSelectedIds(selectedIds.length === queue.items.length ? [] : queue.items.map((item) => item.assetId))} className="min-h-11 px-2 text-[12px] font-semibold text-[var(--blue)]">{selectedIds.length === queue.items.length ? 'ล้างที่เลือก' : 'เลือกทั้งหมด'}</button></div>
      <div className="max-h-56 space-y-1 overflow-y-auto rounded-lg border border-[var(--g200)] p-1">{queue.items.map((item) => <label key={item.assetId} className="flex min-h-11 cursor-pointer items-center gap-3 rounded-md px-2 hover:bg-[var(--surface2)]"><input type="checkbox" className="h-5 w-5 shrink-0" checked={selectedSet.has(item.assetId)} onChange={() => toggle(item.assetId)} /><span className="min-w-0 flex-1"><span className="block truncate text-[12px] font-semibold">{item.name || item.assetId}</span><span className="block truncate font-mono text-[11px] text-[var(--blue)]">{item.serialNumber} · {item.batchId}</span></span></label>)}</div>
      <p className="text-[11px] text-[var(--tmuted)]">หลังสร้างชุด สามารถกด “กำหนดปลายทาง” เพื่อเลือกฟาร์มและโรงเรือนให้ Bundle นี้แยกจากชุดอื่น</p>
      <ModalFooter onClose={onClose} onSubmit={() => onSubmit({ bundleId: bundleId.trim(), bundleName: bundleName.trim(), description, status: 'In Stock' }, selectedIds)} submitLabel={`สร้าง Bundle (${selectedIds.length} Serial)`} submitDisabled={!bundleId.trim() || !bundleName.trim() || duplicate || selectedIds.length === 0} />
    </Modal>
  );
}

function CreateModal({ onClose, onSubmit, existingIds = [] }) {
  // ก้อน ⑤ (B2) — เดารหัสชุดถัดไป BDL-XXX ให้อัตโนมัติ (แก้ได้) + กันกรอกรหัสซ้ำกับที่มีอยู่
  const [bundleId, setBundleId] = useState(() => suggestBundleId(existingIds));
  const [bundleName, setBundleName] = useState('');
  const [description, setDescription] = useState('');
  const [status, setStatus] = useState('In Stock');
  const dup = bundleId.trim()
    ? existingIds.some((x) => (x || '').trim().toUpperCase() === bundleId.trim().toUpperCase())
    : false;
  return (
    <Modal onClose={onClose} title="สร้างชุดใหม่ (Bundle)">
      <Field label="รหัสชุด (Bundle ID) *">
        <input value={bundleId} onChange={(e) => setBundleId(e.target.value.toUpperCase())} placeholder="เช่น BDL-001" className={inp} />
        <div className="mt-1 flex items-center gap-1 text-[11px] flex-wrap">
          {dup
            ? <span className="flex items-center gap-1 font-semibold text-[var(--red)]"><Icon name="warning" size="sm" /> รหัสนี้ถูกใช้แล้ว — ต้องไม่ซ้ำกับชุดเดิม</span>
            : <span className="flex items-center gap-1 text-[var(--tmuted)]"><Icon name="info" size="sm" /> ระบบเดารหัสถัดไป (BDL-…) ให้แล้ว — แก้ได้ตามต้องการ</span>}
        </div>
      </Field>
      <Field label="ชื่อชุดอุปกรณ์ *"><input value={bundleName} onChange={(e) => setBundleName(e.target.value)} placeholder="เช่น ชุดตู้ควบคุมฟาร์ม 1" className={inp} /></Field>
      <Field label="คำอธิบาย"><input value={description} onChange={(e) => setDescription(e.target.value)} className={inp} /></Field>
      <Field label="สถานะเริ่มต้น">
        <select value={status} onChange={(e) => setStatus(e.target.value)} className={inp}>
          <option>In Stock</option>
          <option>Maintenance</option>
        </select>
      </Field>
      <ModalFooter
        onClose={onClose}
        onSubmit={() => onSubmit({ bundleId: bundleId.trim(), bundleName, description, status })}
        submitLabel="บันทึก"
        submitDisabled={!bundleId.trim() || !bundleName.trim() || dup}
      />
    </Modal>
  );
}

function DeployModal({ bundle, farms, onClose, onDeploy, onFarmAdded }) {
  const [farmId, setFarmId] = useState('');
  const [houseId, setHouseId] = useState('');
  const [houses, setHouses] = useState([]);
  const [loadingHouses, setLoadingHouses] = useState(false);
  const [note, setNote] = useState('');
  const [houseAddOpen, setHouseAddOpen] = useState(false);
  const [farmAddOpen, setFarmAddOpen] = useState(false);
  // ฟาร์มที่เพิ่งสร้างจากฟอร์มย่อ — เติมใน dropdown ทันที ก่อน parent refresh มาสมบูรณ์
  const [extraFarms, setExtraFarms] = useState([]);
  const allFarms = [
    ...(farms || []),
    ...extraFarms.filter((f) => !(farms || []).some((x) => x.farmId === f.farmId)),
  ];
  const farmName = allFarms.find((f) => f.farmId === farmId)?.farmName || '';

  // เพิ่มฟาร์มใหม่จากฟอร์มย่อ → เติมเข้า dropdown + เลือกให้เลย + แจ้ง parent refresh
  function handleFarmAdded(site) {
    const f = { farmId: site.siteId, farmName: site.siteName, farmType: site.farmType || '' };
    setExtraFarms((prev) => (prev.some((x) => x.farmId === f.farmId) ? prev : [...prev, f]));
    setFarmId(site.siteId);
    onFarmAdded && onFarmAdded();
  }

  // เพิ่มโรงเรือนใหม่จากฟอร์มย่อ → เลือกให้เลย + refresh รายการโรงเรือนของฟาร์มนี้
  async function handleHouseAdded(house) {
    if (house?.houseId) setHouseId(house.houseId);
    setHouseAddOpen(false);
    try {
      const { data } = await axios.get(`/api/farm-houses/${encodeURIComponent(farmId)}`);
      setHouses(data || []);
    } catch (e) { /* ignore */ }
  }

  // โหลดโรงเรือนของฟาร์มที่เลือก — เปลี่ยนฟาร์มแล้วต้องล้างโรงเรือนเดิมเสมอ
  useEffect(() => {
    setHouseId('');
    setHouseAddOpen(false); // เปลี่ยนฟาร์ม → ปิดฟอร์มย่อเพิ่มโรงเรือน (รหัส auto ของฟาร์มเก่าไม่ valid แล้ว) — เหมือน TransferModal
    if (!farmId) { setHouses([]); return; }
    setLoadingHouses(true);
    axios
      .get(`/api/farm-houses/${encodeURIComponent(farmId)}`)
      .then(({ data }) => setHouses(data || []))
      .catch(() => setHouses([]))
      .finally(() => setLoadingHouses(false));
  }, [farmId]);

  const houseName = houses.find((h) => h.houseId === houseId)?.houseName || '';

  // ก้อน ⑤ (B2) — ฟอร์มย้ายชุด 2 ขั้น: 'form' = เลือกปลายทาง → 'confirm' = สรุปก่อนย้าย (หมายเหตุพับไว้)
  // (DeployModal ถูก mount ใหม่ทุกครั้งที่กด "ย้าย" — step/noteOpen จึงกลับค่าเริ่มต้นเอง ไม่ต้อง reset ผ่าน effect)
  const [step, setStep] = useState('form');
  const [noteOpen, setNoteOpen] = useState(false);

  return (
    <Modal onClose={onClose} title={step === 'form' ? 'ย้ายชุดอุปกรณ์ไปฟาร์ม' : 'สรุปก่อนย้ายชุด'}>
      {bundle && (
        <div className="px-4 py-3 rounded-xl bg-[var(--blue-l)] border border-[var(--blue-b)] mb-4">
          <div className="text-[12px] font-bold text-[var(--blue)]">Bundle ที่เลือก</div>
          <div className="text-[14px] font-bold text-[var(--text)]">{bundle.bundleName}</div>
          <div className="text-[12px] text-[var(--tsub)]">{bundle.assetIds?.length || 0} อุปกรณ์จะถูกย้ายพร้อมกัน</div>
        </div>
      )}
      {step === 'form' ? (
        <>
      <Field label="เลือกฟาร์มปลายทาง *">
        <select
          value={farmId}
          onChange={(e) => setFarmId(e.target.value)}
          className={inp}
        >
          <option value="">— เลือกฟาร์ม —</option>
          {allFarms.map((f) => <option key={f.farmId} value={f.farmId}>{f.farmName} ({f.farmType})</option>)}
        </select>
        {/* เพิ่มฟาร์มใหม่ได้ทันที ไม่ต้องออกจาก modal — เพิ่มแล้ว dropdown refresh + เลือกให้เลย */}
        <FarmInlineAdd mode="trigger" open={farmAddOpen} onOpenChange={setFarmAddOpen} />
      </Field>
      <Field label={`โรงเรือน${farmId ? ' (ตามฟาร์มที่เลือก)' : ''}`}>
        <select
          value={houseId}
          onChange={(e) => setHouseId(e.target.value)}
          className={inp}
          disabled={!farmId || loadingHouses}
        >
          <option value="">{loadingHouses ? '— กำลังโหลดโรงเรือน… —' : '— ไม่ระบุ —'}</option>
          {houses.map((h) => (
            <option key={h.houseId} value={h.houseId}>
              {h.houseId} · {h.houseName}{h.houseType ? ` (${h.houseType})` : ''}
            </option>
          ))}
        </select>
        <HouseInlineAdd
          mode="trigger"
          open={houseAddOpen}
          onOpenChange={setHouseAddOpen}
          siteId={farmId}
          houses={houses}
        />
        {farmId && !loadingHouses && houses.length === 0 && (
          <div className="mt-1.5 flex items-start gap-1.5 text-[11px] text-[var(--tmuted)]">
            <Icon name="info" size="sm" />
            <span>ฟาร์มนี้ยังไม่มีโรงเรือนลงทะเบียน</span>
            <button onClick={() => setHouseAddOpen(true)} className="font-semibold text-[var(--blue)] hover:underline whitespace-nowrap">
              + เพิ่มโรงเรือนทันที
            </button>
          </div>
        )}
      </Field>
      <HouseInlineAdd
        mode="panel"
        open={houseAddOpen}
        onOpenChange={setHouseAddOpen}
        siteId={farmId}
        houses={houses}
        onAdded={handleHouseAdded}
      />
          <ModalFooter
            onClose={onClose}
            onSubmit={() => { setFarmAddOpen(false); setHouseAddOpen(false); setStep('confirm'); }}
            submitLabel="ถัดไป — สรุปก่อนย้าย"
            submitDisabled={!farmId}
          />
        </>
      ) : (
        <>
          {/* ก้อน ⑤ (B2) — การ์ดสรุปก่อนย้าย: ชุด → ปลายทาง (ฟาร์ม › โรงเรือน) + จำนวนอุปกรณ์ */}
          <div className="px-4 py-3 rounded-xl border border-[var(--blue-b)] bg-[var(--blue-l)] space-y-2">
            <div className="flex items-center gap-2 text-[13px] font-bold text-[var(--text)]">
              <Icon name="folder_open" size="sm" /> {bundle?.bundleName || '-'}
              <span className="font-mono text-[11px] font-normal text-[var(--tmuted)]">{bundle?.bundleId}</span>
            </div>
            <div className="flex items-center gap-1.5 text-[13px] font-semibold text-[var(--text)] flex-wrap">
              <Icon name="place" size="sm" className="text-[var(--emerald-d)]" />
              {farmName || '-'}{houseName ? ` › ${houseId} ${houseName}` : ''}
            </div>
            <div className="text-[12px] text-[var(--tsub)]">{bundle?.assetIds?.length || 0} อุปกรณ์จะถูกย้ายตำแหน่งพร้อมกัน</div>
          </div>
          {/* หมายเหตุการย้าย — พับไว้ default (เปิดได้เมื่อต้องการใส่) */}
          <div className="rounded-xl border border-[var(--g200)]">
            <button onClick={() => setNoteOpen(!noteOpen)} className="w-full flex items-center justify-between px-3.5 py-2.5 text-[12px] font-semibold text-[var(--tsub)]">
              <span className="flex items-center gap-1.5"><Icon name="description" size="sm" /> หมายเหตุการย้าย (ถ้ามี)</span>
              <Icon name={noteOpen ? 'expand_less' : 'expand_more'} size="sm" />
            </button>
            {noteOpen && (
              <div className="px-3.5 pb-3"><input value={note} onChange={(e) => setNote(e.target.value)} className={inp} /></div>
            )}
          </div>
      <div className="flex gap-2 px-4 py-3 rounded-xl bg-[var(--amber-l)] text-[var(--amber-d)] text-[12px]">
        <Icon name="warning" size="sm" />
        <span>
          อุปกรณ์ทุกชิ้นในชุดนี้จะถูกย้ายตำแหน่งไปฟาร์มที่เลือกพร้อมกัน
          {houseName ? ` พร้อมระบุโรงเรือน "${houseName}" ให้ทุกชิ้น` : ''}
          {' '}การดำเนินการนี้จะถูกบันทึกในประวัติ
        </span>
      </div>
      <ModalFooter
        onClose={onClose}
        onSubmit={() => onDeploy(farmId, farmName, note, houseId, houseName)}
        submitLabel="ยืนยันย้ายทั้งชุด"
      />
        </>
      )}
      <FarmInlineAdd mode="panel" open={farmAddOpen} onOpenChange={setFarmAddOpen} onAdded={handleFarmAdded} />
      <HouseInlineAdd
        mode="panel"
        open={houseAddOpen}
        onOpenChange={setHouseAddOpen}
        siteId={farmId}
        houses={houses}
        onAdded={handleHouseAdded}
      />
    </Modal>
  );
}

function formatInboundDate(value) {
  return value ? new Intl.DateTimeFormat('th-TH', { timeZone: 'Asia/Bangkok', dateStyle: 'medium' }).format(new Date(value)) : 'ไม่ระบุวัน';
}

function AddAssetModal({ pending, setPending, results, onSearch, onClose, onCommit }) {
  const [q, setQ] = useState('');
  function check(a, idx) {
    const inPending = pending.includes(a.assetId);
    setPending((prev) => inPending ? prev.filter((x) => x !== a.assetId) : [...prev, a.assetId]);
  }
  return (
    <Modal onClose={onClose} title="เพิ่มอุปกรณ์เข้าชุด">
      <Field label="ค้นหา Asset ID หรือ Serial Number">
        <input value={q} onChange={(e) => { setQ(e.target.value); onSearch(e.target.value); }} placeholder="พิมพ์เพื่อค้นหา..." className={inp} />
      </Field>
      <div className="max-h-[280px] overflow-y-auto border border-[var(--g200)] rounded-xl mt-2 divide-y divide-[var(--g100)]">
        {results.map((a, i) => (
          <button key={a.assetId} onClick={() => check(a, i)} className="w-full flex items-center justify-between px-3 py-2.5 text-left hover:bg-[var(--surface2)]">
            <div>
              <div className="text-[13px] font-medium">{a.name}</div>
              <div className="text-[11px] text-[var(--tmuted)]">S/N: {a.serial} · {a.status} · {a.location}</div>
            </div>
            <span className={`text-[12px] font-bold ${pending.includes(a.assetId) ? 'text-[var(--blue)]' : 'text-[var(--tmuted)]'}`}>
              {pending.includes(a.assetId) ? '✓ เลือกแล้ว' : '+ เพิ่ม'}
            </span>
          </button>
        ))}
        {results.length === 0 && q.length >= 2 && <div className="p-3 text-center text-[12px] text-[var(--tmuted)]">ไม่พบอุปกรณ์</div>}
      </div>
      {pending.length > 0 && (
        <div className="mt-3 pt-3 border-t border-[var(--g100)]">
          <div className="text-[12px] font-bold text-[var(--tsub)] mb-2">เลือกไว้ ({pending.length})</div>
          <div className="flex flex-wrap gap-1.5">
            {pending.map((id) => (
              <span key={id} className="flex items-center gap-1 px-2 py-1 rounded-full bg-[var(--blue-l)] text-[var(--blue)] text-[11px]">
                {id}
                <button onClick={() => setPending((p) => p.filter((x) => x !== id))} className="text-[var(--blue)]">×</button>
              </span>
            ))}
          </div>
        </div>
      )}
      <ModalFooter onClose={onClose} onSubmit={onCommit} submitLabel={`เพิ่มทั้งหมดเข้าชุด (${pending.length})`} submitDisabled={pending.length === 0} />
    </Modal>
  );
}

function Modal({ onClose, title, children }) {
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center p-3 sm:p-4 bg-black/40 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
      <div className="w-full max-w-md bg-[var(--surface)] rounded-2xl shadow-xl my-3 sm:my-8" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-4 sm:px-5 py-4 border-b border-[var(--g100)]">
          <span className="text-[15px] font-bold text-[var(--text)]">{title}</span>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)]"><Icon name="close" size="sm" /></button>
        </div>
        <div className="p-4 sm:p-5 space-y-4">{children}</div>
      </div>
    </div>
  );
}

function ModalFooter({ onClose, onSubmit, submitLabel, submitDisabled }) {
  return (
    <div className="flex flex-col-reverse sm:flex-row sm:justify-end gap-2 pt-3 border-t border-[var(--g100)]">
      <button onClick={onClose} className="w-full sm:w-auto min-h-11 sm:min-h-0 px-4 py-2 rounded-lg border border-[var(--g300)] text-[14px] sm:text-[13px] text-[var(--tsub)]">ยกเลิก</button>
      <button onClick={onSubmit} disabled={submitDisabled} className="w-full sm:w-auto min-h-11 sm:min-h-0 px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[14px] sm:text-[13px] font-semibold disabled:opacity-50">✓ {submitLabel}</button>
    </div>
  );
}

function Field({ label, children }) {
  return <div className="space-y-1"><label className="block text-[13px] sm:text-[12px] font-medium text-[var(--tsub)]">{label}</label>{children}</div>;
}

const inp = 'w-full h-11 sm:h-9 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[14px] sm:text-[13px] focus:outline-none focus:border-[var(--blue)]';

function dlCSV(filename, rows) {
  const csv = rows.map((r) => r.map((c) => {
    const v = String(c == null ? '' : c);
    return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
  }).join(',')).join('\r\n');
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

function stamp() {
  const d = new Date();
  return d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
}
