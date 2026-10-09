import { useEffect, useMemo, useState } from 'react';
import axios from 'axios';
import { useNavigate } from 'react-router-dom';
import Icon from './ui/Icon.jsx';
import TransferModal from './TransferModal.jsx';
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

export default function PODispatchPicker() {
  const navigate = useNavigate();
  const [pos, setPOs] = useState([]);
  const [assets, setAssets] = useState([]);
  const [bundleRecords, setBundleRecords] = useState([]);
  const [poNumber, setPONumber] = useState('');
  const [selected, setSelected] = useState([]);
  const [bundleId, setBundleId] = useState('');
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [transferAsset, setTransferAsset] = useState(null);
  const [queue, setQueue] = useState(() => readDispatchQueue());
  const [query, setQuery] = useState('');
  const [message, setMessage] = useState('');

  async function refresh() {
    try {
      const [poResponse, assetResponse] = await Promise.all([
        axios.get('/api/inbound-pos'),
        axios.get('/api/assets'),
      ]);
      const allPOs = (poResponse.data || []).filter((po) => po.poNumber)
        .sort((a, b) => Number(b.status === 'active') - Number(a.status === 'active'));
      setPOs(allPOs);
      setAssets(assetResponse.data || []);
      setPONumber((current) => current && allPOs.some((po) => po.poNumber === current) ? current : allPOs[0]?.poNumber || '');
      setMessage('');
      axios.get('/api/bundles').then(({ data }) => setBundleRecords(data || [])).catch(() => {});
    } catch {
      setMessage('โหลดรายการ PO หรืออุปกรณ์ไม่สำเร็จ ลองโหลดหน้าใหม่อีกครั้ง');
    }
  }

  useEffect(() => {
    let alive = true;
    Promise.all([
      axios.get('/api/inbound-pos'),
      axios.get('/api/assets'),
    ]).then(([poResponse, assetResponse]) => {
      if (!alive) return;
      const allPOs = (poResponse.data || []).filter((po) => po.poNumber)
        .sort((a, b) => Number(b.status === 'active') - Number(a.status === 'active'));
      setPOs(allPOs);
      setAssets(assetResponse.data || []);
      setPONumber(allPOs[0]?.poNumber || '');
    }).catch(() => {
      if (alive) setMessage('โหลดรายการ PO หรืออุปกรณ์ไม่สำเร็จ ลองโหลดหน้าใหม่อีกครั้ง');
    }).finally(() => { if (alive) setLoading(false); });
    axios.get('/api/bundles').then(({ data }) => { if (alive) setBundleRecords(data || []); }).catch(() => {});
    return () => { alive = false; };
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

  async function onTransferSuccess(assetId) {
    const nextQueue = removeDispatchedAssets(queue, [assetId]);
    writeDispatchQueue(nextQueue);
    setQueue(nextQueue);
    setTransferAsset(null);
    await refresh();
    setMessage('ย้ายอุปกรณ์สำเร็จแล้ว');
  }

  return (
    <>
      <details className="rounded-2xl border border-[var(--blue-b)] bg-[var(--blue-l)]/40 shadow-[var(--sh-sm)]">
        <summary className="flex min-h-12 cursor-pointer list-none items-center justify-between gap-3 px-4 py-3 [&::-webkit-details-marker]:hidden">
          <span className="flex min-w-0 items-center gap-2"><Icon name="inventory_2" size="sm" className="text-[var(--blue)]" /><span data-testid="po-dispatch-heading" className="truncate text-[13px] font-semibold text-[var(--text)]">ติดตามและจัดการอุปกรณ์ตาม PO</span></span>
          {queue?.items?.length > 0 && <span className="shrink-0 rounded-full bg-[var(--blue)] px-2 py-0.5 text-[11px] font-bold text-white">คิวจัดชุด {queue.items.length}</span>}
        </summary>
        <div className="space-y-3 border-t border-[var(--blue-b)] bg-[var(--surface)] p-3 sm:p-4">
          {loading ? <div className="py-3 text-center text-[13px] text-[var(--tmuted)]">กำลังโหลด PO...</div> : pos.length === 0 ? <div className="py-3 text-center text-[13px] text-[var(--tmuted)]">ไม่พบ PO ที่มีเลข PO</div> : <>
            <label className="block space-y-1"><span className="text-[12px] font-medium text-[var(--tsub)]">เลือก PO (Active แสดงก่อน)</span><select value={poNumber} onChange={(event) => { setPONumber(event.target.value); setSelected([]); setMessage(''); }} className="h-11 w-full rounded-lg border border-[var(--g200)] bg-[var(--surface)] px-3 text-[13px]">{pos.map((po) => <option key={po.key} value={po.poNumber}>{`PO ${po.poNumber}`} · {po.status === 'active' ? 'กำลังดำเนินการ' : 'ปิดแล้ว'} · คงคลัง {po.stockCount} / ทั้งหมด {po.assetCount}</option>)}</select></label>
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
              <button type="button" onClick={() => setSelected(available.map((asset) => asset.assetId))} className="min-h-11 px-2 text-[12px] font-semibold text-[var(--blue)]">เลือกอุปกรณ์ในคลังทั้งหมด ({available.length})</button>
              <div className="w-full flex-1 sm:min-w-52"><select aria-label="Existing Bundle" value={bundleId} onChange={(event) => setBundleId(event.target.value)} className="h-11 w-full rounded-lg border border-[var(--g200)] bg-[var(--surface)] px-3 text-[12px]" disabled={!bundles.length}><option value="">เลือก Bundle เดิม</option>{bundles.map((bundle) => <option key={bundle.bundleId} value={bundle.bundleId}>{bundle.bundleName} ({bundle.bundleId}) · {bundle.status === 'Deployed' ? bundle.location || bundle.status : bundle.status}</option>)}</select></div>
              <button type="button" data-testid="dispatch-add-existing" onClick={addToExistingBundle} disabled={saving || !selected.length || !selectedBundle} className="min-h-11 w-full rounded-lg border border-[var(--blue)] px-3 text-[12px] font-semibold text-[var(--blue)] disabled:opacity-50 sm:w-auto">{saving ? 'กำลังเพิ่ม...' : `เพิ่ม ${selectedAssets.length} เข้า Bundle เดิม`}</button>
              <button type="button" data-testid="dispatch-create-bundle" onClick={stageDispatch} disabled={!selectedAssets.length} className="min-h-11 w-full rounded-lg bg-[var(--blue)] px-4 text-[13px] font-semibold text-white disabled:opacity-50 sm:w-auto">สร้าง Bundle ใหม่ ({selectedAssets.length})</button>
            </div>
            {message && <div role="status" className="break-words text-[12px] text-[var(--tsub)]">{message}</div>}
            <p className="text-[11px] text-[var(--tmuted)]">รายการที่อยู่ใน Bundle หรือย้ายออกจากคลังจะแสดงสถานะไว้เพื่อให้ทีมติดตาม แต่จะไม่มีปุ่มย้ายซ้ำ</p>
          </>}
        </div>
      </details>
      {transferAsset && <TransferModal open serial={transferAsset.serialNumber} current={{ status: transferAsset.status, location: transferAsset.location, siteName: transferAsset.siteName, user: transferAsset.user }} onClose={() => setTransferAsset(null)} onSuccess={() => onTransferSuccess(transferAsset.assetId)} />}
    </>
  );
}
