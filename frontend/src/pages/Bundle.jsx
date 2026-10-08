import { useEffect, useState, useMemo, useCallback } from 'react';
import axios from 'axios';
import { useSearchParams } from 'react-router-dom';
import Icon from '../components/ui/Icon.jsx';
import StatusBadge from '../components/ui/StatusBadge.jsx';
import TransferModal from '../components/TransferModal.jsx';
import { useBusy, BusyOverlay } from '../components/ui/Busy.jsx';
import { showToast, ToastHost } from '../components/ui/Toast.jsx';
import { buildLocation, buildBundleLocation } from '../utils/location.js';
import { FarmInlineAdd, HouseInlineAdd } from '../components/InlineFarmAdd.jsx';

const BTONES = {
  'In Stock': 'blue',
  Deployed: 'green',
  Maintenance: 'amber',
};

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
  const [loading, setLoading] = useState(true);
  const [searchParams, setSearchParams] = useSearchParams();
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

  // เพิ่มฟาร์มใหม่จากฟอร์มย่อใน DeployModal → refresh รายการฟาร์ม (dropdown อัปเดตทันที)
  async function refreshFarms() {
    try {
      const { data } = await axios.get('/api/farms');
      setFarms(data || []);
    } catch (e) { /* ignore */ }
  }

  // Auto-open จาก Farm Monitor (?open=bundleId)
  useEffect(() => {
    const openId = searchParams.get('open');
    if (openId && bundles.length) {
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
      if (filterFarm && b.farmId !== filterFarm) return false;
      return true;
    });
  }, [bundles, search, filterStatus, filterFarm]);

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

  async function removeAsset(bundleId, assetId) {
    if (!confirm(`นำ ${assetId} ออกจากชุด?`)) return;
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
    if (!confirm('คืนชุดนี้กลับเข้าคลัง?')) return;
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

  return (
    <div className="space-y-4">
      <BusyOverlay label={busy.busyLabel} />
      <ToastHost />
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div className="flex gap-2">
          <button onClick={exportCSV} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[13px] hover:bg-[var(--surface2)]">
            <Icon name="download" size="sm" /> CSV
          </button>
          <button onClick={() => setCreateOpen(true)} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)]">
            <Icon name="add" size="sm" /> สร้างชุดใหม่ (Bundle)
          </button>
        </div>
      </div>

      {/* Stats */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <StatCard label="ชุดทั้งหมด" value={stats.all} icon="folder_open" tone="blue" />
        <StatCard label="อยู่ในคลัง (In Stock)" value={stats.inStock} icon="inventory_2" tone="green" />
        <StatCard label="ติดตั้งที่ฟาร์มแล้ว (Deployed)" value={stats.deployed} icon="local_shipping" tone="amber" />
        <StatCard label="อุปกรณ์ในชุด" value={stats.devices} icon="devices_other" tone="red" />
      </div>

      {/* Filters */}
      <div className="flex flex-wrap items-center gap-3">
        <div className="relative max-w-xs flex-1">
          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[var(--tmuted)]"><Icon name="search" size="sm" /></span>
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="ค้นหาชุดอุปกรณ์..."
            className="w-full h-9 pl-9 pr-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px]" />
        </div>
        <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)} className="h-9 px-3 rounded-lg border border-[var(--g200)] text-[13px]">
          <option value="">ทุกสถานะ</option>
          <option value="In Stock">อยู่ในคลัง (In Stock)</option>
          <option value="Deployed">ติดตั้งที่ฟาร์มแล้ว (Deployed)</option>
          <option value="Maintenance">ซ่อมบำรุง (Maintenance)</option>
        </select>
        <select value={filterFarm} onChange={(e) => setFilterFarm(e.target.value)} className="h-9 px-3 rounded-lg border border-[var(--g200)] text-[13px]">
          <option value="">ทุกฟาร์ม</option>
          {farms.map((f) => <option key={f.farmId} value={f.farmId}>{f.farmName}</option>)}
        </select>
      </div>

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
        />
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
          {loading && <div className="col-span-full text-center py-10 text-[var(--tmuted)]">กำลังโหลด...</div>}
          {!loading && filtered.length === 0 && <div className="col-span-full text-center py-10 text-[var(--tmuted)]">ไม่พบ Bundle</div>}
          {filtered.map((b) => (
            <BundleCard
              key={b.bundleId}
              bundle={b}
              onDetail={() => openDetail(b.bundleId)}
              onDeploy={() => setDeployOpen(b.bundleId)}
              onRecall={() => recall(b.bundleId)}
            />
          ))}
        </div>
      )}

      {/* Create modal */}
      {createOpen && <CreateModal onClose={() => setCreateOpen(false)} onSubmit={createBundle} />}

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
    <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-4 flex items-center gap-3">
      <span className={`flex items-center justify-center w-10 h-10 rounded-xl ${tones[tone]}`}><Icon name={icon} size="md" /></span>
      <div>
        <div className="text-[12px] text-[var(--tmuted)]">{label}</div>
        <div className="text-xl font-bold text-[var(--text)]">{value}</div>
      </div>
    </div>
  );
}

function BundleCard({ bundle, onDetail, onDeploy, onRecall }) {
  const b = bundle;
  const inStock = b.status === 'In Stock';
  const loc = inStock ? 'คลังกลาง' : b.location || b.farmId;
  return (
    <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-4 flex flex-col gap-3">
      <div className="flex items-start justify-between">
        <div>
          <div className="font-mono text-[13px] font-bold text-[var(--text)]">{b.bundleId}</div>
          <div className="text-[13px] font-semibold text-[var(--text)]">{b.bundleName}</div>
          {b.description && <div className="text-[11px] text-[var(--tmuted)] mt-0.5 line-clamp-2">{b.description}</div>}
        </div>
        <StatusBadge status={b.status} tone={BTONES[b.status]} />
      </div>
      <div className="flex items-center gap-3 text-[12px] text-[var(--tsub)]">
        <span className="flex items-center gap-1 flex-shrink-0"><Icon name="devices" size="xs" /> {b.assetIds?.length || 0} อุปกรณ์</span>
        <span className="flex items-center gap-1 min-w-0">
          <Icon name="place" size="xs" className="flex-shrink-0" />
          <span className="truncate" title={loc}>
            {inStock ? 'คลังกลาง' : loc}
          </span>
        </span>
        {!inStock && (
          <button onClick={onDetail} className="ml-auto flex-shrink-0 text-[11px] font-semibold text-[var(--blue)] hover:underline whitespace-nowrap">
            ดูตำแหน่งเต็ม
          </button>
        )}
      </div>
      <div className="flex items-center gap-2 pt-2 border-t border-[var(--g100)]">
        <button onClick={onDetail} className="flex items-center justify-center gap-1 flex-1 px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)] hover:bg-[var(--surface2)]"><Icon name="search" size="xs" /> รายละเอียด</button>
        {b.status === 'In Stock'
          ? <button onClick={onDeploy} className="px-3 py-1.5 rounded-lg bg-[var(--blue)] text-white text-[12px] font-semibold">ย้ายไปฟาร์ม (Deploy)</button>
          : <button onClick={onRecall} className="px-3 py-1.5 rounded-lg bg-[var(--amber)] text-white text-[12px] font-semibold">คืนเข้าคลัง (Recall)</button>}
      </div>
    </div>
  );
}

function BundleDetail({ bundle: b, assets, onBack, onRefresh, onAdd, onRemove, onDeploy, onRecall, onTransfer }) {
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
    <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
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
          <div className="flex items-center gap-2">
            {b.status === 'In Stock'
              ? <button onClick={onDeploy} className="px-3.5 py-2 rounded-lg bg-[var(--blue)] text-white text-[12px] font-semibold"><Icon name="local_shipping" size="sm" /> ย้ายไปฟาร์ม (Deploy)</button>
              : <button onClick={onRecall} className="px-3.5 py-2 rounded-lg bg-[var(--amber)] text-white text-[12px] font-semibold">คืนเข้าคลัง (Recall)</button>}
            <button onClick={onAdd} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg bg-[var(--emerald)] text-white text-[12px] font-semibold"><Icon name="add" size="sm" /> เพิ่มอุปกรณ์</button>
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
            <div key={a.assetId} className="flex flex-wrap items-center gap-x-3 gap-y-2 px-3 py-2.5 rounded-xl border border-[var(--g100)] hover:bg-[var(--surface2)]">
              <div className="flex-1 min-w-[150px]">
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="font-mono text-[12px] font-semibold text-[var(--text)]">{a.serial || a.assetId}</span>
                  <StatusBadge status={a.status} />
                </div>
                <div className="text-[12px] text-[var(--tsub)] truncate">{a.name} <span className="text-[var(--tmuted)]">· {a.assetId} · {a.code}</span></div>
              </div>
              <div className="flex items-center gap-1">
                <a href={`/trace/${encodeURIComponent(a.serial)}`} target="_blank" title="Trace" className="h-9 w-9 flex items-center justify-center rounded-lg text-[var(--blue)] hover:bg-[var(--blue-l)]"><Icon name="description" size="sm" /></a>
                <a href={`/qr?serial=${encodeURIComponent(a.serial)}`} target="_blank" title="QR" className="h-9 w-9 flex items-center justify-center rounded-lg text-[var(--tsub)] hover:bg-[var(--surface2)]"><Icon name="qr_code" size="sm" /></a>
                <button onClick={() => onTransfer(a)} title="โอนย้าย" className="h-9 w-9 flex items-center justify-center rounded-lg text-[var(--blue)] hover:bg-[var(--blue-l)]"><Icon name="local_shipping" size="sm" /></button>
                <span className="w-px h-6 bg-[var(--g200)] mx-1" aria-hidden="true" />
                <button
                  onClick={() => onRemove(a.assetId)}
                  title="ถอดออกจากชุด"
                  className="h-9 px-2.5 flex items-center gap-1 rounded-lg bg-[var(--red-l)] text-[var(--red)] text-[12px] font-semibold hover:bg-[var(--red)] hover:text-white"
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

function CreateModal({ onClose, onSubmit }) {
  const [bundleId, setBundleId] = useState('');
  const [bundleName, setBundleName] = useState('');
  const [description, setDescription] = useState('');
  const [status, setStatus] = useState('In Stock');
  return (
    <Modal onClose={onClose} title="สร้างชุดใหม่ (Bundle)">
      <Field label="รหัสชุด (Bundle ID) *"><input value={bundleId} onChange={(e) => setBundleId(e.target.value.toUpperCase())} placeholder="เช่น BDL-001" className={inp} /></Field>
      <Field label="ชื่อชุดอุปกรณ์ *"><input value={bundleName} onChange={(e) => setBundleName(e.target.value)} placeholder="เช่น ชุดตู้ควบคุมฟาร์ม 1" className={inp} /></Field>
      <Field label="คำอธิบาย"><input value={description} onChange={(e) => setDescription(e.target.value)} className={inp} /></Field>
      <Field label="สถานะเริ่มต้น">
        <select value={status} onChange={(e) => setStatus(e.target.value)} className={inp}>
          <option>In Stock</option>
          <option>Maintenance</option>
        </select>
      </Field>
      <ModalFooter onClose={onClose} onSubmit={() => onSubmit({ bundleId, bundleName, description, status })} submitLabel="บันทึก" />
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

  return (
    <Modal onClose={onClose} title="ย้ายชุดอุปกรณ์ไปฟาร์ม">
      {bundle && (
        <div className="px-4 py-3 rounded-xl bg-[var(--blue-l)] border border-[var(--blue-b)] mb-4">
          <div className="text-[12px] font-bold text-[var(--blue)]">Bundle ที่เลือก</div>
          <div className="text-[14px] font-bold text-[var(--text)]">{bundle.bundleName}</div>
          <div className="text-[12px] text-[var(--tsub)]">{bundle.assetIds?.length || 0} อุปกรณ์จะถูกย้ายพร้อมกัน</div>
        </div>
      )}
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
      <FarmInlineAdd mode="panel" open={farmAddOpen} onOpenChange={setFarmAddOpen} onAdded={handleFarmAdded} />
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
      <Field label="หมายเหตุการย้าย"><input value={note} onChange={(e) => setNote(e.target.value)} className={inp} /></Field>
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
        onSubmit={() => farmId && onDeploy(farmId, farmName, note, houseId, houseName)}
        submitLabel="ย้ายทั้งชุด"
        submitDisabled={!farmId}
      />
    </Modal>
  );
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
    <div className="fixed inset-0 z-50 flex items-start justify-center p-4 bg-black/40 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
      <div className="w-full max-w-md bg-white rounded-2xl shadow-xl my-8" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-[var(--g100)]">
          <span className="text-[15px] font-bold text-[var(--text)]">{title}</span>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)]"><Icon name="close" size="sm" /></button>
        </div>
        <div className="p-5 space-y-4">{children}</div>
      </div>
    </div>
  );
}

function ModalFooter({ onClose, onSubmit, submitLabel, submitDisabled }) {
  return (
    <div className="flex justify-end gap-2 pt-3 border-t border-[var(--g100)]">
      <button onClick={onClose} className="px-4 py-2 rounded-lg border border-[var(--g300)] text-[13px] text-[var(--tsub)]">ยกเลิก</button>
      <button onClick={onSubmit} disabled={submitDisabled} className="px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold disabled:opacity-50">✓ {submitLabel}</button>
    </div>
  );
}

function Field({ label, children }) {
  return <div className="space-y-1"><label className="block text-[12px] font-medium text-[var(--tsub)]">{label}</label>{children}</div>;
}

const inp = 'w-full h-9 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px] focus:outline-none focus:border-[var(--blue)]';

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