import { useEffect, useState, useMemo, useCallback } from 'react';
import axios from 'axios';
import Icon from '../components/ui/Icon.jsx';
import StatusBadge from '../components/ui/StatusBadge.jsx';
import Pagination from '../components/ui/Pagination.jsx';
import TransferModal from '../components/TransferModal.jsx';
import { CATEGORY_FALLBACK, categoryIcon, categoryLabel } from '../data/categories.js';

export default function Asset() {
  const [assets, setAssets] = useState([]);
  const [parts, setParts] = useState([]);
  const [sites, setSites] = useState([]);
  const [categoryList, setCategoryList] = useState(CATEGORY_FALLBACK);
  const [categoryFilter, setCategoryFilter] = useState('');
  const [currentPart, setCurrentPart] = useState('');
  const [partSearch, setPartSearch] = useState('');
  const [assetSearch, setAssetSearch] = useState('');
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(20);
  const [transfer, setTransfer] = useState(null);
  const [history, setHistory] = useState(null);
  const [addOpen, setAddOpen] = useState(false);
  const [bulkOpen, setBulkOpen] = useState(false);
  const [loading, setLoading] = useState(true);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const [a, p, s, c] = await Promise.all([
        axios.get('/api/assets'),
        axios.get('/api/part-catalog'),
        axios.get('/api/farm-sites'),
        axios.get('/api/categories'),
      ]);
      setAssets(a.data || []);
      setParts(p.data || []);
      setSites(s.data || []);
      if (c.data && c.data.length) setCategoryList(c.data);
    } catch (e) {
      console.error('load asset', e);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { load(); }, [load]);

  // Sidebar: parts grouping
  const partGroups = useMemo(() => {
    const map = {};
    assets.forEach((a) => {
      const key = a.partNumber || 'ไม่ระบุ';
      if (!map[key]) map[key] = { count: 0, first: a };
      map[key].count++;
    });
    const q = partSearch.toLowerCase();
    return Object.entries(map)
      .map(([key, v]) => ({
        partNumber: key,
        count: v.count,
        partName: parts.find((p) => p.partNumber === key)?.partName || key,
        category: parts.find((p) => p.partNumber === key)?.category || '',
      }))
      .filter((x) => !q || x.partNumber.toLowerCase().includes(q) || x.partName.toLowerCase().includes(q))
      .sort((a, b) => a.partNumber.localeCompare(b.partNumber, 'th'));
  }, [assets, parts, partSearch]);

  // Table: filter by part + search + category
  const catOf = useCallback((a) => parts.find((p) => p.partNumber === (a.partNumber || a.code))?.category || '', [parts]);
  const filtered = useMemo(() => {
    const k = assetSearch.toLowerCase();
    return assets.filter((a) => {
      if (currentPart && a.partNumber !== currentPart) return false;
      if (categoryFilter && catOf(a) !== categoryFilter) return false;
      if (!k) return true;
      return (
        (a.assetId || '').toLowerCase().includes(k) ||
        (a.code || '').toLowerCase().includes(k) ||
        (a.name || '').toLowerCase().includes(k) ||
        (a.serialNumber || '').toLowerCase().includes(k)
      );
    });
  }, [assets, currentPart, categoryFilter, catOf, assetSearch]);

  const paged = filtered.slice((page - 1) * pageSize, page * pageSize);

  useEffect(() => { setPage(1); }, [currentPart, categoryFilter, assetSearch, pageSize]);

  const currentPartName = currentPart ? (parts.find((p) => p.partNumber === currentPart)?.partName || currentPart) : '';
  const stats = useMemo(() => {
    const usable = filtered.filter((a) => (a.status || '').includes('ใช้งานได้')).length;
    const repair = filtered.filter((a) => (a.status || '').includes('ซ่อม')).length;
    return { total: filtered.length, usable, repair };
  }, [filtered]);

  async function openHistory(serial) {
    try {
      const { data } = await axios.get(`/api/asset-history/${encodeURIComponent(serial)}`);
      setHistory({ serial, logs: data || [] });
    } catch (e) {
      alert('โหลดประวัติไม่ได้');
    }
  }

  function doTransfer(a) {
    setTransfer({
      serial: a.serialNumber,
      current: { status: a.status, location: a.location, siteName: a.siteName, user: a.user },
    });
  }

  function exportCSV() {
    const heads = ['Asset ID', 'Code', 'Name', 'Serial', 'Status', 'Location', 'Site', 'User', 'Category'];
    const rows = [heads];
    filtered.forEach((a) => rows.push([a.assetId, a.code, a.name, a.serialNumber, a.status, a.location, a.siteName, a.user, catOf(a)]));
    dlCSV(`asset_export_${stamp()}.csv`, rows);
  }

  return (
    <div className="space-y-4">
      {/* Header */}
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-2">
          <button
            onClick={() => setBulkOpen(true)}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-[var(--emerald)] text-white text-[13px] font-semibold hover:opacity-90"
          >
            <Icon name="add" size="sm" /> เพิ่มหลายชิ้น
          </button>
          <button
            onClick={() => setAddOpen(true)}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)]"
          >
            <Icon name="add" size="sm" /> เพิ่ม Asset
          </button>
          <button onClick={load} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[13px] hover:bg-[var(--surface2)]">
            <Icon name="refresh" size="sm" /> รีเฟรช
          </button>
        </div>
      </div>

      {/* Stats */}
      <div className="flex flex-wrap gap-2">
        <StatChip label="รวม" value={stats.total} tone="blue" />
        <StatChip label="ใช้งาน" value={stats.usable} tone="green" />
        {stats.repair > 0 && <StatChip label="ซ่อม" value={stats.repair} tone="red" />}
        {currentPartName && <div className="px-3 py-1.5 rounded-full bg-[var(--g100)] text-[12px] text-[var(--tsub)]"><Icon name="inventory_2" size="xs" /> {currentPartName}</div>}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-[240px_1fr] gap-4">
        {/* Part sidebar — จอใหญ่เท่านั้น */}
        <div className="hidden lg:block rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden lg:max-h-[75vh] lg:sticky lg:top-[var(--topbar-h)]">
          <div className="p-2 border-b border-[var(--g100)]">
            <input
              value={partSearch}
              onChange={(e) => setPartSearch(e.target.value)}
              placeholder="ค้นหา Part..."
              className="w-full h-9 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px]"
            />
          </div>
          <div className="overflow-y-auto lg:max-h-[60vh] p-1.5">
            <PartItem
              icon="format_list_bulleted"
              label="ทุก Part"
              count={assets.length}
              active={currentPart === ''}
              onClick={() => setCurrentPart('')}
            />
            {partGroups.map((g) => (
              <PartItem
                key={g.partNumber}
                icon={categoryIcon(g.category)}
                label={g.partNumber}
                sub={g.partName !== g.partNumber ? g.partName : ''}
                count={g.count}
                active={currentPart === g.partNumber}
                onClick={() => setCurrentPart(g.partNumber)}
              />
            ))}
          </div>
        </div>

        {/* Table */}
        <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
          {/* Mobile part selector */}
          <div className="lg:hidden flex items-center gap-2 p-2.5 border-b border-[var(--g100)]">
            <Icon name="inventory_2" size="sm" />
            <select
              value={currentPart}
              onChange={(e) => setCurrentPart(e.target.value)}
              className="flex-1 h-9 px-2.5 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px]"
            >
              <option value="">ทุก Part ({assets.length})</option>
              {partGroups.map((g) => (
                <option key={g.partNumber} value={g.partNumber}>{g.partNumber} ({g.count})</option>
              ))}
            </select>
          </div>

          <div className="flex flex-wrap items-center gap-2 p-3 sm:p-3.5 border-b border-[var(--g100)]">
            <div className="relative max-w-xs flex-1">
              <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[var(--tmuted)]"><Icon name="search" size="sm" /></span>
              <input
                value={assetSearch}
                onChange={(e) => setAssetSearch(e.target.value)}
                placeholder="ค้นหาใน Part นี้..."
                className="w-full h-9 pl-9 pr-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px]"
              />
            </div>
            <select value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)} className="h-9 px-2 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[12px] min-w-[140px]">
              <option value="">ทุกหมวดหมู่</option>
              {categoryList.map((c) => <option key={c.name} value={c.name}>{c.label}</option>)}
            </select>
            <button onClick={exportCSV} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)]">
              <Icon name="download" size="sm" /> CSV
            </button>
          </div>

          {/* Table (จอใหญ่) */}
          <div className="hidden md:block overflow-x-auto">
            <table className="w-full text-[12px]">
              <thead>
                <tr className="text-left text-[var(--tmuted)] bg-[var(--surface2)] border-b border-[var(--g100)]">
                  <th className="px-3 py-2.5 font-medium">Asset ID</th>
                  <th className="px-3 py-2.5 font-medium">Code</th>
                  <th className="px-3 py-2.5 font-medium">Name</th>
                  <th className="px-3 py-2.5 font-medium">Serial</th>
                  <th className="px-3 py-2.5 font-medium">Status</th>
                  <th className="px-3 py-2.5 font-medium">Location</th>
                  <th className="px-3 py-2.5 font-medium">Site</th>
                  <th className="px-3 py-2.5 font-medium">User</th>
                  <th className="px-3 py-2.5 font-medium text-center">Trace</th>
                  <th className="px-3 py-2.5 font-medium text-center">QR</th>
                  <th className="px-3 py-2.5 font-medium text-center">จัดการ</th>
                </tr>
              </thead>
              <tbody>
                {loading && <tr><td colSpan={11} className="text-center py-10 text-[var(--tmuted)]">กำลังโหลด...</td></tr>}
                {!loading && paged.length === 0 && (
                  <tr><td colSpan={11} className="text-center py-10 text-[var(--tmuted)]">ไม่พบอุปกรณ์ใน Part นี้</td></tr>
                )}
                {paged.map((a) => (
                  <tr key={a.serialNumber + a.assetId} className="border-b border-[var(--g100)] hover:bg-[var(--surface2)]">
                    <td className="px-3 py-2 font-mono text-[11px]">{a.assetId}</td>
                    <td className="px-3 py-2 font-mono text-[11px] text-[var(--blue)]">{a.code}</td>
                    <td className="px-3 py-2 font-medium"><span title={categoryLabel(catOf(a))} className="inline-flex items-center gap-1"><Icon name={categoryIcon(catOf(a))} size="xs" className="text-[var(--tsub)]" /> {a.name}</span></td>
                    <td className="px-3 py-2 font-mono text-[11px] text-[var(--blue)]">{a.serialNumber}</td>
                    <td className="px-3 py-2"><StatusBadge status={a.status} /></td>
                    <td className="px-3 py-2 text-[var(--tsub)]">{a.location}</td>
                    <td className="px-3 py-2 text-[var(--tsub)]">{a.siteName}</td>
                    <td className="px-3 py-2 text-[var(--tsub)]">{a.user}</td>
                    <td className="px-3 py-2 text-center">
                      <a href={`/trace.html?serial=${encodeURIComponent(a.serialNumber)}&from=internal`} target="_blank" title="ดู trace" className="text-[var(--blue)]">
                        <Icon name="description" size="sm" />
                      </a>
                    </td>
                    <td className="px-3 py-2 text-center">
                      <a href={`/qr.html?serial=${encodeURIComponent(a.serialNumber)}`} target="_blank" title="QR" className="text-[var(--tsub)]">
                        <Icon name="qr_code" size="sm" />
                      </a>
                    </td>
                    <td className="px-3 py-2 text-center">
                      <button onClick={() => doTransfer(a)} title="โอนย้าย" className="text-[var(--blue)]">
                        <Icon name="local_shipping" size="sm" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Mobile card list */}
          <div className="md:hidden divide-y divide-[var(--g100)]">
            {loading && <div className="text-center py-10 text-[var(--tmuted)] text-[13px]">กำลังโหลด...</div>}
            {!loading && paged.length === 0 && (
              <div className="text-center py-10 text-[var(--tmuted)] text-[13px]">ไม่พบอุปกรณ์ใน Part นี้</div>
            )}
            {paged.map((a) => (
              <div key={a.serialNumber + a.assetId} className="p-3.5 space-y-2.5">
                <div className="flex items-start gap-2.5">
                  <span className="flex items-center justify-center w-8 h-8 rounded-lg shrink-0 bg-[var(--g100)] text-[var(--tsub)]" title={categoryLabel(catOf(a))}>
                    <Icon name={categoryIcon(catOf(a))} size="sm" />
                  </span>
                  <div className="flex-1 min-w-0">
                    <div className="text-[13px] font-semibold text-[var(--text)] leading-snug">{a.name}</div>
                    <div className="text-[11px] font-mono text-[var(--blue)] mt-0.5">{a.assetId} · {a.code}</div>
                  </div>
                  <StatusBadge status={a.status} />
                </div>
                <div className="pl-[42px] grid grid-cols-1 gap-1 text-[12px]">
                  <div className="text-[var(--tsub)]">Serial: <span className="font-mono text-[var(--text)]">{a.serialNumber}</span></div>
                  <div className="text-[var(--tsub)]">ที่อยู่: <span className="text-[var(--text)]">{a.location}</span></div>
                  <div className="text-[var(--tsub)]">Site: <span className="text-[var(--text)]">{a.siteName}</span></div>
                  {a.user && <div className="text-[var(--tsub)]">ผู้ใช้: <span className="text-[var(--text)]">{a.user}</span></div>}
                </div>
                <div className="pl-[42px] flex items-center gap-2">
                  <a href={`/trace.html?serial=${encodeURIComponent(a.serialNumber)}&from=internal`} target="_blank" className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)]">
                    <Icon name="description" size="xs" /> Trace
                  </a>
                  <a href={`/qr.html?serial=${encodeURIComponent(a.serialNumber)}`} target="_blank" className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)]">
                    <Icon name="qr_code" size="xs" /> QR
                  </a>
                  <button onClick={() => doTransfer(a)} className="ml-auto flex items-center gap-1 px-3 py-1.5 rounded-lg bg-[var(--blue)] text-white text-[12px] font-semibold">
                    <Icon name="local_shipping" size="xs" /> โอนย้าย
                  </button>
                </div>
              </div>
            ))}
          </div>

          <Pagination total={filtered.length} page={page} onPage={setPage} pageSize={pageSize} onPageSize={setPageSize} />
        </div>
      </div>

      {/* Transfer */}
      {transfer && (
        <TransferModal
          open
          onClose={() => setTransfer(null)}
          onSuccess={load}
          serial={transfer.serial}
          current={transfer.current}
        />
      )}

      {/* History modal */}
      {history && <HistoryModal data={history} onClose={() => setHistory(null)} />}

      {/* Add Asset / Bulk Add */}
      {addOpen && <AddAssetModal assets={assets} parts={parts} categoryList={categoryList} onClose={() => setAddOpen(false)} onDone={() => { setAddOpen(false); load(); }} />}
      {bulkOpen && <BulkAddModal parts={parts} categoryList={categoryList} onClose={() => setBulkOpen(false)} onDone={() => { setBulkOpen(false); load(); }} />}
    </div>
  );
}

function PartItem({ icon, label, sub, count, active, onClick }) {
  return (
    <button
      onClick={onClick}
      className={`w-full flex items-center gap-2 px-2.5 py-2 rounded-lg text-left transition-colors ${
        active ? 'bg-[var(--blue-l)] text-[var(--blue)]' : 'hover:bg-[var(--surface2)]'
      }`}
    >
      <span className={`flex items-center justify-center w-6 h-6 rounded-md shrink-0 ${active ? 'bg-[var(--blue-l)] text-[var(--blue)]' : 'bg-[var(--g100)] text-[var(--tsub)]'}`}>
        <Icon name={icon} size="xs" />
      </span>
      <span className="flex-1 min-w-0">
        <span className="block text-[12px] font-semibold truncate">{label}</span>
        {sub && <span className="block text-[10px] text-[var(--tmuted)] truncate">{sub}</span>}
      </span>
      <span className={`text-[11px] font-bold ${active ? 'text-[var(--blue)]' : 'text-[var(--tmuted)]'}`}>{count}</span>
    </button>
  );
}

function StatChip({ label, value, tone }) {
  const tones = {
    blue: 'bg-[var(--blue-l)] text-[var(--blue)]',
    green: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]',
    red: 'bg-[var(--red-l)] text-[var(--red)]',
  };
  return (
    <div className={`flex items-center gap-2 px-3 py-1.5 rounded-full ${tones[tone]}`}>
      <span className="text-[11px] font-medium">{label}</span>
      <span className="text-[13px] font-bold">{value}</span>
    </div>
  );
}

function HistoryModal({ data, onClose }) {
  const { serial, logs } = data;
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center p-4 bg-black/40 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
      <div className="w-full max-w-2xl bg-white rounded-2xl shadow-xl my-8" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-[var(--g100)]">
          <div className="flex items-center gap-2">
            <span className="text-[var(--blue)]"><Icon name="assignment" size="sm" /></span>
            <span className="text-[15px] font-bold">ประวัติ</span>
            <span className="px-2 py-0.5 rounded-full bg-[var(--blue-l)] text-[var(--blue)] font-mono text-[11px]">{serial}</span>
            <span className="text-[12px] text-[var(--tmuted)]">{logs.length} รายการ</span>
          </div>
          <div className="flex items-center gap-2">
            <a href={`/trace.html?serial=${encodeURIComponent(serial)}&from=internal`} target="_blank" className="text-[12px] text-[var(--blue)] font-semibold">ดูหน้าเต็ม →</a>
            <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)]">
              <Icon name="close" size="sm" />
            </button>
          </div>
        </div>
        <div className="max-h-[60vh] overflow-y-auto p-5 space-y-3">
          {logs.length === 0 && <div className="flex items-center justify-center gap-2 text-center py-10 text-[var(--tmuted)]"><Icon name="inbox" size="sm" /> ยังไม่มีประวัติในระบบ</div>}
          {logs.map((l, i) => (
            <HistoryItem key={i} log={l} isLatest={i === 0} />
          ))}
        </div>
      </div>
    </div>
  );
}

function HistoryItem({ log, isLatest }) {
  const action = log.action || '';
  let icon = 'chevron_right';
  const iconClass = isLatest ? 'bg-[var(--blue-l)] text-[var(--blue)]' : 'bg-[var(--g100)] text-[var(--tsub)]';
  if (isLatest) icon = 'place';
  else if (action.includes('ลงทะเบียน') || action.includes('เพิ่ม')) icon = 'add';
  else if (action.includes('ซ่อม')) icon = 'build';
  else if (action.includes('คืน')) icon = 'undo';
  return (
    <div className="flex gap-3 items-start">
      <span className={`w-7 h-7 flex items-center justify-center rounded-full ${iconClass}`}><Icon name={icon} size="xs" /></span>
      <div className="flex-1 min-w-0">
        <div className="flex items-center justify-between gap-2">
          <span className="text-[12px] font-semibold text-[var(--text)]">{action}</span>
          <span className="text-[11px] text-[var(--tmuted)] whitespace-nowrap">{log.date}</span>
        </div>
        <div className="text-[12px] text-[var(--tsub)] mt-0.5">
          {log.from && log.from !== '-' ? `${log.from} → ` : ''}{log.to}
        </div>
        {(log.user && log.user !== '-') && <div className="flex items-center gap-1 text-[11px] text-[var(--tmuted)] mt-0.5"><Icon name="person" size="xs" /> {log.user}{log.remark && log.remark !== '-' ? ` · ${log.remark}` : ''}</div>}
      </div>
    </div>
  );
}

// ───── เพิ่ม Asset เดี่ยว (ทุก user) — จำลอง addAssetNewPart ─────
function AddAssetModal({ assets, parts, categoryList = CATEGORY_FALLBACK, onClose, onDone }) {
  const [partNumber, setPartNumber] = useState('');
  const [partName, setPartName] = useState('');
  const [category, setCategory] = useState('');
  const [unit, setUnit] = useState('ชิ้น');
  const [description, setDescription] = useState('');
  const [status, setStatus] = useState('ใช้งานได้');
  const [location, setLocation] = useState('Stock');
  const [siteName, setSiteName] = useState('Intranin');
  const [user, setUser] = useState('');
  const [sites, setSites] = useState([]);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    axios.get('/api/farm-sites').then(({ data }) => setSites(data || [])).catch(() => {});
  }, []);

  const maxId = useMemo(() => assets.reduce((m, a) => Math.max(m, parseInt(a.assetId) || 0), 0) + 1, [assets]);
  const assetId = String(maxId).padStart(4, '0');

  // Serial อัตโนมัติ: SN-<PART>-<ddmmyyyy>-<run next>
  const serialPreview = useMemo(() => {
    const pn = partNumber.trim().toUpperCase();
    if (!pn) return '';
    const now = new Date();
    const ds = String(now.getDate()).padStart(2, '0') + String(now.getMonth() + 1).padStart(2, '0') + now.getFullYear();
    const prefix = `SN-${pn}-${ds}`;
    let maxNum = 0;
    assets.forEach((a) => {
      const s = a.serialNumber || '';
      const p = s.split('-');
      if (p.length >= 3 && p[p.length - 2] === ds) {
        const n = parseInt(p[p.length - 1]);
        if (!isNaN(n) && n > maxNum) maxNum = n;
      }
    });
    return `${prefix}-${String(maxNum + 1).padStart(4, '0')}`;
  }, [partNumber, assets]);

  async function submit() {
    const pn = partNumber.trim().toUpperCase();
    if (!pn || !partName.trim()) return alert('กรุณากรอก Part Number และชื่ออุปกรณ์');
    const exists = parts.some((p) => p.partNumber === pn);
    setBusy(true);
    try {
      if (!exists) {
        const r = await axios.post('/api/add-part', { partNumber: pn, partName: partName.trim(), category, description, unit });
        if (r.data && r.data.error) { alert('ไม่สำเร็จ: ' + r.data.error); return; }
      }
      const r = await axios.post('/api/add-asset', {
        assetId,
        name: partName.trim(),
        code: pn,
        partNumber: pn,
        serialNumber: serialPreview,
        status,
        location: location.trim() || '-',
        siteName,
        user: user.trim() || '',
      });
      if (r.data.success) {
        alert(`เพิ่ม Part + Asset สำเร็จ\nAsset ID: ${assetId}\nSerial: ${serialPreview}`);
        onDone();
      } else alert('ไม่สำเร็จ: ' + (r.data.error || 'เพิ่ม Asset ไม่สำเร็จ'));
    } catch (e) {
      alert('เกิดข้อผิดพลาด: ' + (e.response?.data?.error || e.message));
    } finally {
      setBusy(false);
    }
  }

  return (
    <AssetModal title="เพิ่ม Part + Asset ใหม่" submitLabel="เพิ่ม Part + Asset" onSubmit={submit} busy={busy} onClose={onClose}>
      <p className="text-[11px] text-[var(--tmuted)] -mt-1">สร้าง Part ใหม่ในคลังและเพิ่ม Asset 1 ชิ้น — Asset ID/Serial รันอัตโนมัติ</p>
      <AField label="Part Number *"><input value={partNumber} onChange={(e) => setPartNumber(e.target.value.toUpperCase())} placeholder="เช่น SEN.TEMP-RS485" className={ain} /></AField>
      <AField label="ชื่ออุปกรณ์ *"><input value={partName} onChange={(e) => setPartName(e.target.value)} className={ain} /></AField>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <AField label="หมวดหมู่">
          <select value={category} onChange={(e) => setCategory(e.target.value)} className={ain}>
            <option value="">— เลือกหมวด —</option>
            {categoryList.map((c) => <option key={c.name} value={c.name}>{c.label} ({c.name})</option>)}
          </select>
        </AField>
        <AField label="หน่วย"><select value={unit} onChange={(e) => setUnit(e.target.value)} className={ain}><option>ชิ้น</option><option>ชุด</option><option>ลัง</option><option>ตัว</option></select></AField>
      </div>
      <AField label="รายละเอียด"><input value={description} onChange={(e) => setDescription(e.target.value)} className={ain} /></AField>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <AField label="Asset ID (อัตโนมัติ)"><input value={assetId} disabled className={ain + ' bg-[var(--g100)] text-[var(--tmuted)]'} /></AField>
        <AField label="Serial (อัตโนมัติ)"><input value={serialPreview || '— กรอก Part Number —'} disabled className={ain + ' bg-[var(--g100)] text-[var(--tmuted)]'} /></AField>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <AField label="สถานะ"><select value={status} onChange={(e) => setStatus(e.target.value)} className={ain}><option>ใช้งานได้</option><option>สำรอง</option><option>ส่งซ่อม</option><option>ชำรุด/สูญหาย</option></select></AField>
        <AField label="ไซต์งาน"><select value={siteName} onChange={(e) => setSiteName(e.target.value)} className={ain}>
          <option value="Intranin">บริษัท Intranin (คลังกลาง)</option>
          {sites.filter((s) => s.siteName && s.siteName !== 'Intranin').map((s) => <option key={s.siteId} value={s.siteName}>{s.siteName}</option>)}
        </select></AField>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <AField label="Location">
          <select value={location} onChange={(e) => setLocation(e.target.value)} className={ain}>
            <option value="Stock">Stock (คลังกลาง)</option>
            {sites.filter((s) => s.siteName && s.siteName !== 'Intranin').map((s) => <option key={s.siteId} value={s.siteName}>{s.siteName}</option>)}
          </select>
        </AField>
        <AField label="ผู้รับผิดชอบ"><input value={user} onChange={(e) => setUser(e.target.value)} placeholder="ชื่อผู้รับผิดชอบ" className={ain} /></AField>
      </div>
    </AssetModal>
  );
}

// ───── เพิ่มหลายชิ้น (ทุก user) — จำลอง confirmBulkAdd ─────
function BulkAddModal({ parts, categoryList = CATEGORY_FALLBACK, onClose, onDone }) {
  const [partNumber, setPartNumber] = useState('');
  const [partName, setPartName] = useState('');
  const [category, setCategory] = useState('');
  const [qty, setQty] = useState('');
  const [status, setStatus] = useState('ใช้งานได้');
  const [siteName, setSiteName] = useState('Intranin');
  const [location, setLocation] = useState('Stock');
  const [user, setUser] = useState('');
  const [sites, setSites] = useState([]);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    axios.get('/api/farm-sites').then(({ data }) => setSites(data || [])).catch(() => {});
  }, []);

  function selectPart(pn) {
    const p = parts.find((x) => x.partNumber === pn);
    if (p) { setPartNumber(p.partNumber); setPartName(p.partName); setCategory(p.category || ''); }
    else { setPartNumber(''); setPartName(''); setCategory(''); }
  }

  const preview = useMemo(() => {
    const n = parseInt(qty) || 0;
    if (!partNumber.trim() || !n) return [];
    const now = new Date();
    const ds = String(now.getDate()).padStart(2, '0') + String(now.getMonth() + 1).padStart(2, '0') + now.getFullYear();
    const prefix = `SN-${partNumber.trim().toUpperCase()}-${ds}`;
    return Array.from({ length: Math.min(n, 5) }, (_, i) => `${prefix}-${String(i + 1).padStart(4, '0')}`);
  }, [partNumber, qty]);

  async function submit() {
    const n = parseInt(qty);
    if (!partNumber.trim() || !partName.trim() || !n) return alert('กรอกข้อมูลให้ครบ');
    setBusy(true);
    try {
      const { data } = await axios.post('/api/bulk-add-asset', {
        partNumber: partNumber.trim().toUpperCase(),
        partName: partName.trim(),
        qty: n,
        status,
        siteName,
        location: location.trim(),
        user: user.trim(),
      });
      if (data.success) {
        alert(`เพิ่ม ${data.added} ชิ้นสำเร็จ\nSerial: ${data.firstSerial} ~ ${data.lastSerial}`);
        onDone();
      } else alert('เกิดข้อผิดพลาด: ' + (data.error || ''));
    } catch (e) {
      alert('เกิดข้อผิดพลาด: ' + (e.response?.data?.error || e.message));
    } finally {
      setBusy(false);
    }
  }

  return (
    <AssetModal title="เพิ่ม Asset หลายชิ้น" submitLabel="เพิ่ม Asset" onSubmit={submit} busy={busy} onClose={onClose}>
      <p className="text-[11px] text-[var(--tmuted)] -mt-1">สร้าง Asset หลายชิ้นพร้อมกันภายใต้ Part เดียว — Serial รันอัตโนมัติ</p>
      <AField label="เลือก Part (จากคลัง)">
        <select value={partNumber} onChange={(e) => selectPart(e.target.value)} className={ain}>
          <option value="">-- เลือก Part --</option>
          {parts.map((p) => <option key={p.partNumber} value={p.partNumber}>{p.partNumber} - {p.partName} ({p.totalQty || 0} ชิ้น)</option>)}
        </select>
      </AField>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <AField label="Part Number"><input value={partNumber} onChange={(e) => setPartNumber(e.target.value.toUpperCase())} className={ain} /></AField>
        <AField label="ชื่ออุปกรณ์"><input value={partName} onChange={(e) => setPartName(e.target.value)} className={ain} /></AField>
      </div>
      <AField label="จำนวน *"><input type="number" min="1" value={qty} onChange={(e) => setQty(e.target.value)} className={ain} /></AField>
      {preview.length > 0 && (
        <div className="px-3 py-2 rounded-lg bg-[var(--blue-l)] border border-[var(--blue-b)] text-[11px] text-[var(--blue)]">
          <div className="font-bold mb-1">Serial ที่จะสร้าง ({parseInt(qty) || 0} ชิ้น):</div>
          {preview.map((s) => <div key={s} className="font-mono">{s}</div>)}
          {parseInt(qty) > 5 && <div>... และอีก {parseInt(qty) - 5} ชิ้น</div>}
        </div>
      )}
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <AField label="สถานะ"><select value={status} onChange={(e) => setStatus(e.target.value)} className={ain}><option>ใช้งานได้</option><option>สำรอง</option><option>ส่งซ่อม</option><option>ชำรุด/สูญหาย</option></select></AField>
        <AField label="ไซต์งาน"><select value={siteName} onChange={(e) => setSiteName(e.target.value)} className={ain}>
          <option value="Intranin">บริษัท Intranin (คลังกลาง)</option>
          {sites.filter((s) => s.siteName && s.siteName !== 'Intranin').map((s) => <option key={s.siteId} value={s.siteName}>{s.siteName}</option>)}
        </select></AField>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <AField label="หมวดหมู่">
          <select value={category} onChange={(e) => setCategory(e.target.value)} className={ain}>
            <option value="">— เลือกหมวด —</option>
            {categoryList.map((c) => <option key={c.name} value={c.name}>{c.label} ({c.name})</option>)}
          </select>
        </AField>
        <AField label="Location">
          <select value={location} onChange={(e) => setLocation(e.target.value)} className={ain}>
            <option value="Stock">Stock (คลังกลาง)</option>
            {sites.filter((s) => s.siteName && s.siteName !== 'Intranin').map((s) => <option key={s.siteId} value={s.siteName}>{s.siteName}</option>)}
          </select>
        </AField>
      </div>
      <AField label="ผู้รับผิดชอบ"><input value={user} onChange={(e) => setUser(e.target.value)} placeholder="ชื่อผู้รับผิดชอบ" className={ain} /></AField>
    </AssetModal>
  );
}

function AssetModal({ title, submitLabel, onSubmit, busy, onClose, children }) {
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center p-4 bg-black/40 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
      <div className="w-full max-w-lg bg-white rounded-2xl shadow-xl my-8" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-[var(--g100)]">
          <span className="text-[15px] font-bold text-[var(--text)]">{title}</span>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)]"><Icon name="close" size="sm" /></button>
        </div>
        <div className="p-5 space-y-3">{children}</div>
        <div className="flex justify-end gap-2 px-5 py-4 border-t border-[var(--g100)] bg-[var(--surface2)] rounded-b-2xl">
          <button onClick={onClose} className="px-4 py-2 rounded-lg border border-[var(--g300)] text-[13px] text-[var(--tsub)]">ยกเลิก</button>
          <button onClick={onSubmit} disabled={busy} className="flex items-center gap-1 px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold disabled:opacity-60">
            <Icon name="save" size="sm" /> {busy ? 'กำลังบันทึก...' : submitLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

function AField({ label, children }) {
  return <div className="space-y-1"><label className="block text-[12px] font-medium text-[var(--tsub)]">{label}</label>{children}</div>;
}

const ain = 'w-full h-9 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px] focus:outline-none focus:border-[var(--blue)]';

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