import { useEffect, useMemo, useState } from 'react';
import axios from 'axios';
import { useNavigate } from 'react-router-dom';
import Icon from './ui/Icon.jsx';
import TransferModal from './TransferModal.jsx';
import { showToast } from './ui/Toast.jsx';
import { readDispatchQueue, writeDispatchQueue, removeDispatchedAssets } from '../utils/dispatchQueue.js';
import { buildLocation } from '../utils/location.js';

function isStockAsset(asset) {
  const location = String(asset.location || '').trim().toLowerCase();
  const site = String(asset.siteName || '').trim().toLowerCase();
  return !asset.bundleId && (
    location === 'stock' || location === 'intranin'
    || site === 'intranin' || site === 'stock'
    || String(asset.status || '').trim().toLowerCase() === 'in stock'
  );
}

// Step 2: ใช้ buildLocation ตัวเดียวกับหน้า Asset/Scan/Bundle — ตำแหน่งแสดงตรงกันทุกหน้า
function locationLabel(asset) {
  const loc = buildLocation({
    siteName: asset.siteName,
    houseName: asset.houseName,
    houseId: asset.houseId,
    location: asset.location,
    bundleId: asset.bundleId,
  });
  return loc.full || 'ไม่ระบุตำแหน่ง';
}

// 2 แท็บชัดเจนตาม spec — เข้าใจง่าย ไม่ซับซ้อน
const TABS = [
  { id: 'po', icon: 'local_shipping', label: 'โอนย้ายตาม PO' },
  { id: 'general', icon: 'swap_horiz', label: 'โอนย้ายอุปกรณ์ทั่วไป (นอก PO)' },
];

// Transfer Center Modal (Clean Home Layout) — แทนการ์ด "ติดตามและจัดการอุปกรณ์ตาม PO" ที่เคยแผ่กว้างกลางหน้าแรก:
//   • ปุ่มวงกลม "ย้ายอุปกรณ์" บนหน้าแรก → เปิดป๊อปอัปกลางหน้าจอทันที (หน้าแรกไม่ถูกดันเนื้อหาอีกต่อไป)
//   • แท็บ 1 "โอนย้ายตาม PO" — เลือก PO ที่ดำเนินงานอยู่ → ติ๊กอุปกรณ์ในคลัง → "ย้ายรายตัว" (เลือกฟาร์ม/โรงเรือนปลายทางใน TransferModal)
//     หรือจัดชุด: เพิ่มเข้า Bundle เดิม / สร้าง Bundle ใหม่ (stage คิว → /bundle?dispatch=1)
//   • แท็บ 2 "โอนย้ายอุปกรณ์ทั่วไป (นอก PO)" — พิมพ์ค้นหา Serial/ชื่อ/รหัส → กดโอนย้ายได้ทันที
//   • preloadAssets / preloadBundles = ข้อมูลที่หน้าแรกโหลดไว้แล้ว (ส่งเข้ามา ไม่ดึง /api/assets ซ้ำ)
//   • onTransferSuccess = แจ้งหน้าแรกให้รีเฟรชข้อมูลหลังย้ายสำเร็จ (bump dataVersion)
//   • โอนย้ายสำเร็จ → ปิด Modal อัตโนมัติ + toast แจ้งเตือน (ผ่าน ToastHost ระดับหน้าของ Home)
export default function TransferCenterModal({ open, onClose, preloadAssets = null, preloadBundles = null, onTransferSuccess, initialTab = 'po' }) {
  const navigate = useNavigate();
  const [tab, setTab] = useState(initialTab);
  const [pos, setPOs] = useState([]);
  const [fetchedAssets, setFetchedAssets] = useState([]);
  // ข้อมูลอุปกรณ์: ถ้าผู้ใช้ (Home) ส่ง preloadAssets มา → ใช้ของเขาตรง ๆ (รีเฟรชจาก dataVersion ของ Home อัตโนมัติ)
  // ไม่งั้น (ใช้เดี่ยว ๆ เช่นเทสต์/หน้าอื่น) → ดึงเองตามเดิม
  const assets = preloadAssets || fetchedAssets;
  const [fetchedBundles, setFetchedBundles] = useState([]);
  const bundleRecords = preloadBundles || fetchedBundles;
  const [poNumber, setPONumber] = useState('');
  const [selected, setSelected] = useState([]);
  const [bundleId, setBundleId] = useState('');
  const [loading, setLoading] = useState(!preloadAssets);
  const [saving, setSaving] = useState(false);
  const [transferAsset, setTransferAsset] = useState(null);
  const [queue, setQueue] = useState(() => readDispatchQueue());
  const [query, setQuery] = useState('');
  const [generalQuery, setGeneralQuery] = useState('');
  const [message, setMessage] = useState('');

  async function refresh() {
    try {
      const poResponse = await axios.get('/api/inbound-pos');
      const allPOs = (poResponse.data || []).filter((po) => po.poNumber)
        .sort((a, b) => Number(b.status === 'active') - Number(a.status === 'active'));
      setPOs(allPOs);
      setPONumber((current) => current && allPOs.some((po) => po.poNumber === current) ? current : allPOs[0]?.poNumber || '');
      // แบบเปิดจากหน้าแรก: อุปกรณ์/ชุดเป็นของ Home (preload) — รีเฟรชผ่าน onTransferSuccess แทน
      if (!preloadAssets) {
        const assetResponse = await axios.get('/api/assets');
        setFetchedAssets(assetResponse.data || []);
      }
      setMessage('');
      if (!preloadBundles) axios.get('/api/bundles').then(({ data }) => setFetchedBundles(data || [])).catch(() => {});
    } catch {
      setMessage('โหลดรายการ PO หรืออุปกรณ์ไม่สำเร็จ ลองปิดแล้วเปิดใหม่อีกครั้ง');
    }
  }

  useEffect(() => {
    if (!open) return; // Home render แบบ conditional — mount = เปิด Modal ใช้งาน
    let alive = true;
    const poPromise = axios.get('/api/inbound-pos').then((poResponse) => {
      if (!alive) return;
      const allPOs = (poResponse.data || []).filter((po) => po.poNumber)
        .sort((a, b) => Number(b.status === 'active') - Number(a.status === 'active'));
      setPOs(allPOs);
      setPONumber(allPOs[0]?.poNumber || '');
    });
    const assetPromise = preloadAssets
      ? Promise.resolve() // มีข้อมูลจากหน้าแรกส่งมาแล้ว — ไม่ดึงซ้ำ
      : axios.get('/api/assets').then((assetResponse) => { if (alive) setFetchedAssets(assetResponse.data || []); });
    Promise.all([poPromise, assetPromise])
      .catch(() => {
        if (alive) setMessage('โหลดรายการ PO หรืออุปกรณ์ไม่สำเร็จ ลองปิดแล้วเปิดใหม่อีกครั้ง');
      })
      .finally(() => { if (alive) setLoading(false); });
    if (preloadBundles) return () => { alive = false; };
    axios.get('/api/bundles').then(({ data }) => { if (alive) setFetchedBundles(data || []); }).catch(() => {});
    return () => { alive = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps -- preloadAssets/preloadBundles อ่านครั้งเดียวตอน mount (ค่าเริ่มต้นจาก Home)
  }, []);

  const poAssets = useMemo(() => assets.filter((asset) => (
    poNumber && asset.poNumber === poNumber && asset.serialNumber
  )), [assets, poNumber]);
  const available = useMemo(() => poAssets.filter(isStockAsset), [poAssets]);
  const bundles = useMemo(() => {
    const byId = new Map(bundleRecords.filter((bundle) => bundle.bundleId).map((bundle) => [bundle.bundleId, bundle]));
    assets.filter((asset) => asset.bundleId).forEach((asset) => {
      if (!byId.has(asset.bundleId)) byId.set(asset.bundleId, {
        bundleId: asset.bundleId,
        bundleName: asset.bundleName || asset.bundleId,
        status: 'มีอุปกรณ์แล้ว',
        location: asset.siteName && asset.siteName !== 'Intranin' ? asset.siteName : asset.location,
      });
    });
    return [...byId.values()].sort((a, b) => a.bundleName.localeCompare(b.bundleName, 'th'));
  }, [assets, bundleRecords]);
  const visible = poAssets.filter((asset) => [asset.serialNumber, asset.assetId, asset.name, asset.bundleName]
    .some((value) => String(value || '').toLowerCase().includes(query.trim().toLowerCase())));
  const selectedAssets = available.filter((asset) => selected.includes(asset.assetId));
  const selectedBundle = bundles.find((bundle) => bundle.bundleId === bundleId);

  // Non-PO Transfer: ค้นหาอุปกรณ์ทั้งระบบเพื่อโอนย้าย — ตัดรายการของ PO ที่กำลังเลือกออก (แท็บ PO จัดการเอง)
  const generalVisible = useMemo(() => assets.filter((asset) => (
    asset.serialNumber && asset.poNumber !== poNumber
    && [asset.serialNumber, asset.assetId, asset.name, asset.code, asset.siteName, asset.houseName]
      .some((value) => String(value || '').toLowerCase().includes(generalQuery.trim().toLowerCase()))
  )), [assets, poNumber, generalQuery]);

  useEffect(() => { setSelected(available.map((asset) => asset.assetId)); }, [available]);

  function stageDispatch() {
    if (!selectedAssets.length) return;
    const items = selectedAssets.map(({ assetId, serialNumber, name, batchId }) => ({ assetId, serialNumber, name, batchId }));
    const nextQueue = { poNumber, items, createdAt: new Date().toISOString() };
    writeDispatchQueue(nextQueue);
    setQueue(nextQueue);
    navigate('/bundle?dispatch=1');
  }

  async function addToExistingBundle() {
    if (!selectedBundle || !selectedAssets.length) return;
    setSaving(true);
    setMessage('');
    try {
      const { data: latestAssets } = await axios.get('/api/assets');
      const latestById = new Map((latestAssets || []).map((asset) => [asset.assetId, asset]));
      const stale = selectedAssets.filter((asset) => {
        const latest = latestById.get(asset.assetId);
        return !latest || latest.bundleId || latest.poNumber !== poNumber || !isStockAsset(latest);
      });
      if (stale.length) {
        await refresh();
        setMessage(`รายการเปลี่ยนแปลงแล้ว: ${stale.map((asset) => asset.serialNumber).join(', ')} โหลดข้อมูลใหม่ก่อนเพิ่มเข้า Bundle`);
        return;
      }
      await axios.post(`/api/bundles/${encodeURIComponent(bundleId)}/assets/bulk`, {
        assetIds: selectedAssets.map((asset) => asset.assetId),
      });
      const nextQueue = removeDispatchedAssets(queue, selectedAssets.map((asset) => asset.assetId));
      writeDispatchQueue(nextQueue);
      setQueue(nextQueue);
      setSelected([]);
      await refresh();
      setMessage(`เพิ่ม ${selectedAssets.length} Serial เข้า ${selectedBundle.bundleName || selectedBundle.bundleId} แล้ว`);
    } catch (error) {
      setMessage(error.response?.data?.error || 'เพิ่มอุปกรณ์เข้า Bundle ไม่สำเร็จ');
    } finally {
      setSaving(false);
    }
  }

  // โอนย้ายสำเร็จ → ปิดฟอร์มรายตัว + ปิด Transfer Center ทันที + toast แจ้งเตือน + รีเฟรชข้อมูล
  async function handleTransferSuccess(assetId) {
    const nextQueue = removeDispatchedAssets(queue, [assetId]);
    writeDispatchQueue(nextQueue);
    setQueue(nextQueue);
    setTransferAsset(null);
    onClose?.();
    onTransferSuccess?.(); // bump dataVersion → หน้าแรกโหลดข้อมูลใหม่ (preload ใหม่ไหลเข้าอัตโนมัติ)
    await refresh();
    // ดัน toast ไป tick ถัดไป — ให้ React ถอด ToastHost ของ TransferModal ออกก่อน
    // → toast ไปโชว์ที่ ToastHost ระดับหน้า (Home) และอยู่ต่อแม้ Transfer Center ปิดแล้ว
    setTimeout(() => showToast('ย้ายอุปกรณ์สำเร็จแล้ว'), 0);
  }

  // ── แท็บ 1: โอนย้ายตาม PO — เลือก PO → ติ๊กอุปกรณ์ในคลัง → ย้ายรายตัว / จัดชุด (Bundle) ──
  function renderPOPanel() {
    return (
      <div role="tabpanel" aria-labelledby="transfer-tab-po" data-testid="panel-po" className="space-y-3">
        <p className="text-[11px] text-[var(--tmuted)]">① เลือก PO → ② ติ๊กอุปกรณ์ในคลัง → ③ กด "ย้ายรายตัว" แล้วเลือกฟาร์ม/โรงเรือนปลายทาง</p>
        {loading ? (
          <div className="py-6 text-center text-[13px] text-[var(--tmuted)]">กำลังโหลด PO...</div>
        ) : pos.length === 0 ? (
          <div className="py-6 text-center text-[13px] text-[var(--tmuted)]">ไม่พบ PO ที่มีเลข PO</div>
        ) : (
          <>
            <label className="block space-y-1"><span className="text-[12px] font-medium text-[var(--tsub)]">เลือก PO</span><select value={poNumber} onChange={(event) => { setPONumber(event.target.value); setSelected([]); setMessage(''); }} className="h-11 w-full rounded-lg border border-[var(--g200)] bg-[var(--surface)] px-3 text-[13px]">{pos.map((po) => <option key={po.key} value={po.poNumber}>{`PO ${po.poNumber}`} · {po.status === 'active' ? 'กำลังดำเนินการ' : 'ปิดแล้ว'} · คงคลัง {po.stockCount} / ทั้งหมด {po.assetCount}</option>)}</select></label>
            <div className="grid grid-cols-2 gap-2 text-center sm:grid-cols-4">
              <div className="rounded-lg bg-[var(--surface2)] px-2 py-2"><div className="text-[11px] text-[var(--tmuted)]">ทั้งหมด</div><div className="text-[14px] font-bold">{poAssets.length}</div></div>
              <div className="rounded-lg bg-[var(--blue-l)] px-2 py-2"><div className="text-[11px] text-[var(--blue)]">รอจัดการ</div><div className="text-[14px] font-bold text-[var(--blue)]">{available.length}</div></div>
              <div className="rounded-lg bg-[var(--emerald-l)] px-2 py-2"><div className="text-[11px] text-[var(--emerald-d)]">อยู่ใน Bundle</div><div className="text-[14px] font-bold text-[var(--emerald-d)]">{poAssets.filter((asset) => asset.bundleId).length}</div></div>
              <div className="rounded-lg bg-[var(--surface2)] px-2 py-2"><div className="text-[11px] text-[var(--tmuted)]">ย้ายออกจากคลังแล้ว</div><div className="text-[14px] font-bold">{poAssets.filter((asset) => !asset.bundleId && !isStockAsset(asset)).length}</div></div>
            </div>
            <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="ค้น Serial / Asset ID / Bundle" className="h-11 w-full rounded-lg border border-[var(--g200)] bg-[var(--surface2)] px-3 text-[13px]" />
            <div className="max-h-72 space-y-1 overflow-y-auto rounded-lg border border-[var(--g200)] p-1">
              {visible.map((asset) => {
                const stock = isStockAsset(asset);
                const inBundle = Boolean(asset.bundleId);
                return <div key={asset.assetId} className="flex min-h-14 items-center gap-2 rounded-md px-2 py-1 hover:bg-[var(--surface2)]">
                  {stock && <input type="checkbox" aria-label={`เลือก ${asset.serialNumber}`} className="h-5 w-5 shrink-0" checked={selected.includes(asset.assetId)} onChange={() => setSelected((current) => current.includes(asset.assetId) ? current.filter((id) => id !== asset.assetId) : [...current, asset.assetId])} />}
                  <div className="min-w-0 flex-1"><div className="truncate text-[12px] font-semibold">{asset.name || asset.code}</div><div className="truncate font-mono text-[11px] text-[var(--blue)]">{asset.serialNumber} · {asset.batchId || 'ไม่ระบุล็อต'}</div><div className={`truncate text-[11px] ${inBundle ? 'text-[var(--emerald-d)]' : stock ? 'text-[var(--blue)]' : 'text-[var(--tmuted)]'}`}>{inBundle ? `อยู่ใน Bundle: ${asset.bundleName || asset.bundleId}` : stock ? 'อยู่ในคลัง · เลือกเพื่อจัดการ' : `ย้ายแล้ว · ${locationLabel(asset)}`}</div></div>
                  {stock && <button type="button" aria-label={`Transfer ${asset.serialNumber}`} onClick={() => setTransferAsset(asset)} className="min-h-11 shrink-0 rounded-lg border border-[var(--g300)] px-2 text-[11px] font-semibold text-[var(--blue)] sm:px-3 sm:text-[12px]">ย้ายรายตัว</button>}
                </div>;
              })}
              {visible.length === 0 && <div className="py-4 text-center text-[12px] text-[var(--tmuted)]">ไม่พบอุปกรณ์ใน PO นี้</div>}
            </div>
            <div className="flex flex-wrap items-center gap-2">
              <button type="button" onClick={() => setSelected(available.map((asset) => asset.assetId))} className="min-h-11 px-2 text-[12px] font-semibold text-[var(--blue)]">เลือกทั้งหมดในคลัง ({available.length})</button>
              <div className="w-full flex-1 sm:min-w-52"><select aria-label="Existing Bundle" value={bundleId} onChange={(event) => setBundleId(event.target.value)} className="h-11 w-full rounded-lg border border-[var(--g200)] bg-[var(--surface)] px-3 text-[12px]" disabled={!bundles.length}><option value="">เลือก Bundle เดิม</option>{bundles.map((bundle) => <option key={bundle.bundleId} value={bundle.bundleId}>{bundle.bundleName} ({bundle.bundleId}) · {bundle.status === 'Deployed' ? bundle.location || bundle.status : bundle.status}</option>)}</select></div>
              <button type="button" data-testid="dispatch-add-existing" onClick={addToExistingBundle} disabled={saving || !selected.length || !selectedBundle} className="min-h-11 w-full rounded-lg border border-[var(--blue)] px-3 text-[12px] font-semibold text-[var(--blue)] disabled:opacity-50 sm:w-auto">{saving ? 'กำลังเพิ่ม...' : `เพิ่ม ${selectedAssets.length} เข้า Bundle เดิม`}</button>
              <button type="button" data-testid="dispatch-create-bundle" onClick={stageDispatch} disabled={!selectedAssets.length} className="min-h-11 w-full rounded-lg bg-[var(--blue)] px-4 text-[13px] font-semibold text-white disabled:opacity-50 sm:w-auto">สร้าง Bundle ใหม่ ({selectedAssets.length})</button>
            </div>
            {message && <div role="status" className="break-words text-[12px] text-[var(--tsub)]">{message}</div>}
          </>
        )}
      </div>
    );
  }

  // ── แท็บ 2: โอนย้ายอุปกรณ์ทั่วไป (นอก PO) — ค้นหา Serial/ชื่อ → โอนย้ายทันที ──
  function renderGeneralPanel() {
    return (
      <div role="tabpanel" aria-labelledby="transfer-tab-general" data-testid="panel-general">
        <div data-testid="general-transfer-section" className="space-y-2">
          <p className="text-[11px] text-[var(--tmuted)]">พิมพ์ค้นหา Serial / ชื่อ / รหัสอุปกรณ์ → กด "โอนย้าย" แล้วเลือกปลายทาง</p>
          <input value={generalQuery} onChange={(event) => setGeneralQuery(event.target.value)} placeholder="พิมพ์ค้นหา Serial / ชื่อ / รหัสอุปกรณ์ แล้วกดโอนย้ายได้ทันที" className="h-11 w-full rounded-lg border border-[var(--g200)] bg-[var(--surface)] px-3 text-[13px]" />
          <div className="max-h-72 space-y-1 overflow-y-auto rounded-lg border border-[var(--g200)] bg-[var(--surface)] p-1">
            {generalVisible.map((asset) => (
              <div key={asset.assetId} className="flex min-h-12 items-center gap-2 rounded-md px-2 py-1 hover:bg-[var(--surface2)]">
                <div className="min-w-0 flex-1">
                  <div className="truncate text-[12px] font-semibold">{asset.name || asset.code}</div>
                  <div className="truncate font-mono text-[11px] text-[var(--blue)]">{asset.serialNumber}</div>
                  <div className="truncate text-[11px] text-[var(--tsub)]">{locationLabel(asset)}</div>
                </div>
                <button type="button" aria-label={`Transfer ${asset.serialNumber}`} onClick={() => setTransferAsset(asset)} className="min-h-11 shrink-0 rounded-lg border border-[var(--g300)] px-2 text-[11px] font-semibold text-[var(--blue)] sm:px-3 sm:text-[12px]">โอนย้าย</button>
              </div>
            ))}
            {generalVisible.length === 0 && <div className="py-3 text-center text-[12px] text-[var(--tmuted)]">{generalQuery.trim() ? `ไม่พบอุปกรณ์ที่ตรงกับ "${generalQuery.trim()}"` : 'ยังไม่มีอุปกรณ์ที่ไม่อยู่ใน PO นี้'}</div>}
          </div>
        </div>
      </div>
    );
  }

  if (!open) return null;

  return (
    <>
      <div className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto bg-black/40 p-3 sm:p-4 backdrop-blur-sm" onClick={() => onClose?.()}>
        <section
          role="dialog"
          aria-modal="true"
          aria-labelledby="transfer-center-title"
          data-testid="transfer-center"
          className="my-3 sm:my-8 flex max-h-[calc(100vh-1.5rem)] w-full max-w-2xl flex-col overflow-hidden rounded-2xl bg-[var(--surface)] shadow-[var(--sh-lg)] sm:max-h-[calc(100vh-4rem)]"
          onClick={(e) => e.stopPropagation()}
        >
          {/* ── Header: ชื่อ + คิวจัดชุด + ปุ่มปิด ── */}
          <header className="flex shrink-0 items-center justify-between gap-3 border-b border-[var(--g100)] px-4 py-3">
            <div className="flex min-w-0 items-center gap-2.5">
              <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-[var(--blue-l)] text-[var(--blue)]">
                <Icon name="local_shipping" size="md" />
              </span>
              <div className="min-w-0">
                <h2 id="transfer-center-title" className="truncate text-[15px] font-bold text-[var(--text)]">ย้ายอุปกรณ์</h2>
                <p className="truncate text-[11px] text-[var(--tmuted)]">เลือกอุปกรณ์ → ระบุฟาร์ม/โรงเรือนปลายทาง → ยืนยัน</p>
              </div>
            </div>
            <div className="flex shrink-0 items-center gap-2">
              {queue?.items?.length > 0 && <span className="rounded-full bg-[var(--blue)] px-2 py-0.5 text-[11px] font-bold text-white">คิวจัดชุด {queue.items.length}</span>}
              <button onClick={() => onClose?.()} title="ปิด" aria-label="ปิด Transfer Center" className="flex h-9 w-9 items-center justify-center rounded-full border border-[var(--g200)] text-[var(--tsub)] hover:bg-[var(--surface2)] hover:text-[var(--text)]">
                <Icon name="close" size="sm" />
              </button>
            </div>
          </header>

          {/* ── Tabs: 2 แท็บชัดเจน — มือถือซ้อนกัน / จอกว้างเรียงข้างกัน ── */}
          <div role="tablist" aria-label="ประเภทการโอนย้าย" className="grid shrink-0 grid-cols-1 gap-1 border-b border-[var(--g100)] bg-[var(--surface2)]/60 p-2 sm:grid-cols-2">
            {TABS.map((t) => (
              <button
                key={t.id}
                type="button"
                role="tab"
                id={`transfer-tab-${t.id}`}
                aria-selected={tab === t.id}
                data-testid={`tab-${t.id}`}
                onClick={() => setTab(t.id)}
                className={`flex min-h-11 items-center justify-center gap-1.5 rounded-lg px-3 text-[12.5px] font-semibold transition-colors ${tab === t.id ? 'bg-[var(--blue)] text-white shadow-[var(--sh-sm)]' : 'text-[var(--tsub)] hover:bg-[var(--surface)] hover:text-[var(--text)]'}`}
              >
                <Icon name={t.icon} size="sm" />
                {t.label}
              </button>
            ))}
          </div>

          {/* ── Content (scroll ภายใน Modal) ── */}
          <div className="min-h-40 flex-1 overflow-y-auto p-3 sm:p-4">
            {tab === 'po' ? renderPOPanel() : renderGeneralPanel()}
          </div>
        </section>
      </div>
      {/* ฟอร์มโอนย้ายรายตัว — กางทับ Transfer Center: เลือกปลายทาง (ฟาร์ม/คลัง/ซ่อม) → ยืนยัน → ปิดทั้งคู่ */}
      {transferAsset && <TransferModal open serial={transferAsset.serialNumber} current={{ status: transferAsset.status, location: transferAsset.location, siteName: transferAsset.siteName, user: transferAsset.user }} onClose={() => setTransferAsset(null)} onSuccess={() => handleTransferSuccess(transferAsset.assetId)} />}
    </>
  );
}