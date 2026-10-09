import { useEffect, useState, useMemo, useCallback } from 'react';
import axios from 'axios';
import Icon from '../components/ui/Icon.jsx';
import StatusBadge from '../components/ui/StatusBadge.jsx';
import StatPill from '../components/ui/StatPill.jsx';
import Pagination from '../components/ui/Pagination.jsx';
import LocationPath from '../components/ui/LocationPath.jsx';
import TransferModal from '../components/TransferModal.jsx';
import { AssetHistoryModal } from '../components/AssetHistory.jsx';
import AddDeviceModal from '../components/AddDeviceModal.jsx';
import { useBusy, BusyOverlay } from '../components/ui/Busy.jsx';
import { showToast, ToastHost } from '../components/ui/Toast.jsx';
import { CATEGORY_FALLBACK, categoryIcon, categoryLabel } from '../data/categories.js';
import { buildLocation } from '../utils/location.js';
import { matchesInboundFilters } from '../utils/inbound.js';

export default function Asset() {
  const [assets, setAssets] = useState([]);
  const [parts, setParts] = useState([]);
  const [sites, setSites] = useState([]);
  const [categoryList, setCategoryList] = useState(CATEGORY_FALLBACK);
  const [categoryFilter, setCategoryFilter] = useState('');
  const [currentPart, setCurrentPart] = useState('');
  const [partSearch, setPartSearch] = useState('');
  const [assetSearch, setAssetSearch] = useState('');
  const [batchSearch, setBatchSearch] = useState('');
  const [receivedFrom, setReceivedFrom] = useState('');
  const [receivedTo, setReceivedTo] = useState('');
  const [recentBatches, setRecentBatches] = useState([]);
  const [batchHistoryOpen, setBatchHistoryOpen] = useState(false);
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(20);
  const [transfer, setTransfer] = useState(null);
  const [selectedSerials, setSelectedSerials] = useState([]);
  const [historySerial, setHistorySerial] = useState('');
  const [addOpen, setAddOpen] = useState(false);
  const [loading, setLoading] = useState(true);
  const busy = useBusy();

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const [a, p, s, c, b] = await Promise.all([
        axios.get('/api/assets'),
        axios.get('/api/part-catalog'),
        axios.get('/api/farm-sites'),
        axios.get('/api/categories'),
        axios.get('/api/inbound-pos?status=active&limit=12').catch(() => ({ data: [] })),
      ]);
      setAssets(a.data || []);
      setParts(p.data || []);
      setSites(s.data || []);
      if (c.data && c.data.length) setCategoryList(c.data);
      setRecentBatches(b.data || []);
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
    const bk = batchSearch.trim().toLowerCase();
    return assets.filter((a) => {
      if (currentPart && a.partNumber !== currentPart) return false;
      if (categoryFilter && catOf(a) !== categoryFilter) return false;
      if (!matchesInboundFilters(a, { query: bk, from: receivedFrom, to: receivedTo })) return false;
      if (!k) return true;
      return (
        (a.assetId || '').toLowerCase().includes(k) ||
        (a.code || '').toLowerCase().includes(k) ||
        (a.name || '').toLowerCase().includes(k) ||
        (a.serialNumber || '').toLowerCase().includes(k)
      );
    });
  }, [assets, currentPart, categoryFilter, catOf, assetSearch, batchSearch, receivedFrom, receivedTo]);

  const paged = filtered.slice((page - 1) * pageSize, page * pageSize);

  useEffect(() => { setPage(1); }, [currentPart, categoryFilter, assetSearch, batchSearch, receivedFrom, receivedTo, pageSize]);

  const currentPartName = currentPart ? (parts.find((p) => p.partNumber === currentPart)?.partName || currentPart) : '';
  const stats = useMemo(() => {
    const usable = filtered.filter((a) => (a.status || '').includes('ใช้งานได้')).length;
    const repair = filtered.filter((a) => (a.status || '').includes('ซ่อม')).length;
    return { total: filtered.length, usable, repair };
  }, [filtered]);

  function openHistory(serial) { if (serial) setHistorySerial(serial); }

  function doTransfer(a) {
    setTransfer({
      serial: a.serialNumber,
      current: { status: a.status, location: a.location, siteName: a.siteName, user: a.user },
    });
  }

  function toggleSelected(a) {
    const serial = a.serialNumber;
    setSelectedSerials((prev) => prev.includes(serial) ? prev.filter((s) => s !== serial) : [...prev, serial]);
  }

  function startBulkTransfer() {
    const selected = assets.filter((a) => selectedSerials.includes(a.serialNumber) && a.serialNumber);
    if (selected.length) setTransfer({ assets: selected });
  }

  function selectPO(po) {
    const batchIds = new Set((po.batches || []).map((batch) => batch.batchId));
    const serials = assets.filter((a) => (po.poNumber ? a.poNumber === po.poNumber : batchIds.has(a.batchId)) && a.serialNumber).map((a) => a.serialNumber);
    setCurrentPart(''); setCategoryFilter(''); setAssetSearch(''); setReceivedFrom(''); setReceivedTo('');
    setBatchSearch(po.poNumber || po.batchId || '');
    setSelectedSerials(serials);
    showToast(`เลือกอุปกรณ์ PO ${po.poNumber || po.batchId} จำนวน ${serials.length} ชิ้นแล้ว`);
  }

  function exportCSV() {
    const heads = ['Asset ID', 'Code', 'Name', 'Serial', 'Batch ID', 'Received At', 'PO Number', 'Supplier', 'Status', 'อยู่ในชุด', 'ตำแหน่ง (ฟาร์ม › โรงเรือน › จุดติดตั้ง)', 'User', 'Category'];
    const rows = [heads];
    filtered.forEach((a) => rows.push([
      a.assetId, a.code, a.name, a.serialNumber, a.batchId || '', a.receivedAt || '', a.poNumber || '', a.supplier || '', a.status,
      a.bundleName || a.bundleId || '',
      buildLocation({ siteName: a.siteName, houseName: a.houseName, houseId: a.houseId, location: a.location, bundleId: a.bundleId }).full,
      a.user, catOf(a),
    ]));
    dlCSV(`asset_export_${stamp()}.csv`, rows);
  }

  return (
    <div className="space-y-4">
      {!transfer && !addOpen && <ToastHost />}
      {/* Header */}
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-2">
          <button
            onClick={() => setAddOpen(true)}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)]"
          >
            <Icon name="add" size="sm" /> เพิ่มอุปกรณ์ใหม่
          </button>
          <button onClick={load} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[13px] hover:bg-[var(--surface2)]">
            <Icon name="refresh" size="sm" /> รีเฟรช
          </button>
        </div>
      </div>

      {selectedSerials.length > 0 && (
        <div className="flex flex-wrap items-center justify-between gap-2 rounded-lg border border-[var(--blue-b)] bg-[var(--blue-l)] px-3 py-2">
          <span className="text-[13px] font-medium text-[var(--blue-d)]">เลือกแล้ว {selectedSerials.length} ชิ้น</span>
          <div className="flex gap-2">
            <button onClick={() => setSelectedSerials([])} className="min-h-10 px-3 rounded-lg border border-[var(--g300)] bg-[var(--surface)] text-[13px]">ล้างที่เลือก</button>
            <button onClick={startBulkTransfer} className="min-h-10 px-3 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold">ย้ายที่เลือก</button>
          </div>
        </div>
      )}

      {/* Stats */}
      <div className="flex flex-wrap gap-2">
        <StatPill label="รวม" value={stats.total} icon="inventory_2" tone="blue" />
        <StatPill label="ใช้งาน" value={stats.usable} icon="check_circle" tone="green" />
        {stats.repair > 0 && <StatPill label="ซ่อม" value={stats.repair} icon="build" tone="red" />}
        {currentPartName && <div className="px-3 py-1.5 rounded-full bg-[var(--g100)] text-[12px] text-[var(--tsub)]"><Icon name="inventory_2" size="xs" /> {currentPartName}</div>}
      </div>

      <section className="rounded-xl border border-[var(--g200)] bg-[var(--surface)] p-3 sm:p-4 space-y-2">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="text-[13px] font-semibold text-[var(--text)]">PO ที่กำลังดำเนินการ</div>
            <button onClick={() => setBatchHistoryOpen(true)} className="min-h-11 px-3 rounded-lg border border-[var(--g300)] text-[12px] font-semibold text-[var(--blue)] hover:bg-[var(--blue-l)]">ดูประวัติ PO ทั้งหมด</button>
          </div>
          {recentBatches.length > 0 ? (
          <div className="flex gap-2 overflow-x-auto pb-1">
            {recentBatches.map((po) => (
              <button key={po.key} onClick={() => selectPO(po)} className="min-h-11 shrink-0 rounded-lg border border-[var(--g200)] px-3 text-left hover:border-[var(--blue)]">
                <span className="block text-[12px] font-semibold text-[var(--blue)]">{po.poNumber ? `PO ${po.poNumber}` : po.batchId}</span>
                <span className="block text-[11px] text-[var(--tsub)]">คงคลัง {po.stockCount}/{po.assetCount} ชิ้น · {po.batchCount} ล็อต</span>
              </button>
            ))}
          </div>
          ) : <div className="text-[12px] text-[var(--tmuted)]">ไม่มี PO ที่ยังรอดำเนินการ</div>}
      </section>

      <div className="grid grid-cols-1 xl:grid-cols-[240px_minmax(0,1fr)] gap-4">
        {/* Part sidebar — จอใหญ่เท่านั้น */}
        <div className="hidden xl:block rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden xl:max-h-[75vh] xl:sticky xl:top-[var(--topbar-h)]">
          <div className="p-2 border-b border-[var(--g100)]">
            <input
              value={partSearch}
              onChange={(e) => setPartSearch(e.target.value)}
              placeholder="ค้นหา Part..."
              className="w-full h-9 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px]"
            />
          </div>
          <div className="overflow-y-auto xl:max-h-[60vh] p-1.5">
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
        <div className="rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
          {/* Mobile part selector */}
          <div className="xl:hidden flex items-center gap-2 p-2.5 border-b border-[var(--g100)]">
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
            <div className="w-full grid grid-cols-1 sm:grid-cols-3 gap-2">
              <input value={batchSearch} onChange={(e) => setBatchSearch(e.target.value)} placeholder="ค้นหา Batch ID / PO / Supplier" className="h-11 min-w-0 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] px-3 text-[13px]" />
              <label className="flex min-w-0 items-center gap-2 rounded-lg border border-[var(--g200)] px-2"><span className="shrink-0 text-[11px] text-[var(--tsub)]">รับเข้าตั้งแต่</span><input aria-label="รับเข้าตั้งแต่" type="date" value={receivedFrom} onChange={(e) => setReceivedFrom(e.target.value)} className="h-10 min-w-0 flex-1 bg-transparent text-[12px]" /></label>
              <label className="flex min-w-0 items-center gap-2 rounded-lg border border-[var(--g200)] px-2"><span className="shrink-0 text-[11px] text-[var(--tsub)]">ถึง</span><input aria-label="รับเข้าถึง" type="date" value={receivedTo} onChange={(e) => setReceivedTo(e.target.value)} className="h-10 min-w-0 flex-1 bg-transparent text-[12px]" /></label>
            </div>
          </div>

          {/* Table (จอใหญ่) */}
          <div className="hidden md:block overflow-x-auto">
            <table className="w-full min-w-[900px] text-[12px]">
<thead>
              <tr className="text-left text-[12px] text-[var(--tsub)] border-b border-[var(--g200)] bg-[var(--surface2)]">
                  <th className="w-9 px-2 py-2.5 font-medium">เลือก</th>
                  <th className="px-2.5 py-2.5 font-medium">ชื่อ / รหัส</th>
                  <th className="px-2.5 py-2.5 font-medium">ตำแหน่งฟาร์ม / ชุด</th>
                  <th className="px-2.5 py-2.5 font-medium">Serial</th>
                  <th className="px-2.5 py-2.5 font-medium">ล็อตรับเข้า</th>
                  <th className="px-2.5 py-2.5 font-medium">Status</th>
                  <th className="hidden px-2.5 py-2.5 font-medium 2xl:table-cell">User</th>
                  <th className="sticky right-0 z-20 border-l border-[var(--g200)] bg-[var(--surface2)] px-2 py-2.5 text-center font-medium shadow-[-10px_0_10px_-10px_rgba(15,23,42,0.25)]">จัดการ</th>
                </tr>
              </thead>
              <tbody>
                {loading && <tr><td colSpan={8} className="text-center py-10 text-[var(--tmuted)]">กำลังโหลด...</td></tr>}
                {!loading && paged.length === 0 && (
                  <tr><td colSpan={8} className="text-center py-10 text-[var(--tmuted)]">ไม่พบอุปกรณ์ใน Part นี้</td></tr>
                )}
                {paged.map((a) => (
                  <tr key={a.serialNumber + a.assetId} className="group border-b border-[var(--g100)] hover:bg-[var(--surface2)]">
                    <td className="px-2 py-2"><input aria-label={`เลือก ${a.serialNumber}`} type="checkbox" checked={selectedSerials.includes(a.serialNumber)} onChange={() => toggleSelected(a)} /></td>
                    <td className="px-2.5 py-2">
                      {/* Step 2: รวมรหัสเข้าคอลัมน์ชื่อ (ชื่อ + รหัส·AssetID) ประหยัดความกว้างให้ตารางลงตัวที่ 1024px */}
                      <div className="flex items-center gap-1 text-[13px] font-medium text-[var(--text)]">
                        <span title={categoryLabel(catOf(a))} className="flex shrink-0 text-[var(--tsub)]"><Icon name={categoryIcon(catOf(a))} size="xs" /></span>
                        <span className="min-w-0 max-w-[160px] truncate" title={a.name}>{a.name}</span>
                      </div>
                      <div className="font-mono text-[10px] text-[var(--tmuted)]"><span className="block max-w-[160px] truncate" title={`${a.code} · ${a.assetId}`}>{a.code} · {a.assetId}</span></div>
                    </td>
                    <td className="px-2.5 py-2">
                      {/* Step 2: ตำแหน่งฟาร์ม/โรงเรือน = Primary Visual Indicator ย้ายขึ้นมาหลังชื่อ ไม่ต้องเปิด modal */}
                      <div className="min-w-0 max-w-[180px]">
                        {a.bundleId && (
                          <span
                            title={`อยู่ในชุด ${a.bundleName || a.bundleId}`}
                            className="mb-0.5 inline-flex max-w-full items-center gap-1 rounded bg-[var(--blue-l)] px-1.5 py-px text-[10px] font-medium text-[var(--blue)] border border-[var(--blue-b)]"
                          >
                            <Icon name="inventory_2" size="xs" className="flex-shrink-0" />
                            <span className="truncate">{a.bundleName || a.bundleId}</span>
                          </span>
                        )}
                        <LocationPath
                          variant="primary"
                          siteName={a.siteName}
                          houseName={a.houseName}
                          houseId={a.houseId}
                          location={a.location}
                          bundleId={a.bundleId}
                        />
                      </div>
                    </td>
                    <td className="px-2.5 py-2 font-mono text-[11px] text-[var(--blue)]"><span className="block max-w-[100px] truncate" title={a.serialNumber}>{a.serialNumber}</span></td>
                    <td className="px-2.5 py-2"><InboundBadge asset={a} /></td>
                    <td className="px-2.5 py-2"><StatusBadge status={a.status} /></td>
                    <td className="hidden px-2.5 py-2 text-[var(--tsub)] 2xl:table-cell">{a.user}</td>
                    <td className="sticky right-0 z-10 border-l border-[var(--g200)] bg-[var(--surface)] px-1.5 py-2 shadow-[-10px_0_10px_-10px_rgba(15,23,42,0.25)] group-hover:bg-[var(--surface2)]">
                      <div className="flex items-center justify-center gap-0.5">
                        <button onClick={() => openHistory(a.serialNumber)} title="ดูประวัติ" className="h-8 w-8 flex items-center justify-center rounded-lg text-[var(--blue)] hover:bg-[var(--blue-l)]">
                          <Icon name="description" size="sm" />
                        </button>
                        <a href={`/qr?serial=${encodeURIComponent(a.serialNumber)}`} target="_blank" title="QR" className="h-8 w-8 flex items-center justify-center rounded-lg text-[var(--tsub)] hover:bg-[var(--surface2)]">
                          <Icon name="qr_code" size="sm" />
                        </a>
                        <button onClick={() => doTransfer(a)} title="โอนย้าย" className="h-8 w-8 flex items-center justify-center rounded-lg text-[var(--blue)] hover:bg-[var(--blue-l)]">
                          <Icon name="local_shipping" size="sm" />
                        </button>
                      </div>
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
                  <input aria-label={`เลือก ${a.serialNumber}`} className="mt-2 h-5 w-5 shrink-0" type="checkbox" checked={selectedSerials.includes(a.serialNumber)} onChange={() => toggleSelected(a)} />
                  <span className="flex items-center justify-center w-8 h-8 rounded-lg shrink-0 bg-[var(--g100)] text-[var(--tsub)]" title={categoryLabel(catOf(a))}>
                    <Icon name={categoryIcon(catOf(a))} size="sm" />
                  </span>
                  <div className="flex-1 min-w-0">
                    <div className="text-[13px] font-semibold text-[var(--text)] leading-snug">{a.name}</div>
                    <div className="text-[11px] font-mono text-[var(--blue)] mt-0.5">{a.assetId} · {a.code}</div>
                  </div>
                  <StatusBadge status={a.status} />
                </div>
                {/* Step 2: ตำแหน่งฟาร์ม/โรงเรือน = ข้อมูลแรกที่ช่างต้องเห็นบนการ์ด (Primary Indicator) */}
                <div className="pl-[42px]">
                  <LocationPath
                    variant="primary"
                    siteName={a.siteName}
                    houseName={a.houseName}
                    houseId={a.houseId}
                    location={a.location}
                    bundleId={a.bundleId}
                    showChips
                  />
                </div>
                <div className="pl-[42px] grid grid-cols-1 gap-1 text-[12px]">
                  <div className="text-[var(--tsub)]">Serial: <span className="font-mono text-[var(--text)]">{a.serialNumber}</span></div>
                  <InboundBadge asset={a} />
                  {a.bundleId && (
                    <div className="flex items-center gap-1.5">
                      <span className="text-[var(--tsub)]">อยู่ในชุด:</span>
                      <span className="inline-flex items-center gap-1 rounded-md bg-[var(--blue-l)] border border-[var(--blue-b)] px-1.5 py-0.5 text-[11px] font-medium text-[var(--blue)]">
                        <Icon name="inventory_2" size="xs" /> {a.bundleName || a.bundleId}
                      </span>
                    </div>
                  )}
                  {a.user && <div className="text-[var(--tsub)]">ผู้ใช้: <span className="text-[var(--text)]">{a.user}</span></div>}
                </div>
                <div className="pl-[42px] flex items-center gap-2">
                  <button onClick={() => openHistory(a.serialNumber)} className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)]">
                    <Icon name="description" size="xs" /> ประวัติ
                  </button>
                  <a href={`/qr?serial=${encodeURIComponent(a.serialNumber)}`} target="_blank" className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)]">
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
          onSuccess={async (result) => {
            if (result?.failedSerials) setSelectedSerials(result.failedSerials);
            else if (transfer.assets) setSelectedSerials([]);
            await load();
          }}
          assets={transfer.assets}
          serial={transfer.serial || transfer.assets?.[0]?.serialNumber}
          current={transfer.current}
        />
      )}

      {/* History modal */}
      {historySerial && <AssetHistoryModal key={historySerial} serial={historySerial} onClose={() => setHistorySerial('')} />}
      {batchHistoryOpen && <BatchHistoryModal onClose={() => setBatchHistoryOpen(false)} onSelectPO={(po) => { setBatchHistoryOpen(false); selectPO(po); }} onClosed={load} />}

      {/* เพิ่มอุปกรณ์ใหม่ (โฟลว์เดียว: เลือก/สร้าง Part → Serial อัตโนมัติ กันซ้ำ) */}
      {addOpen && <AddDeviceModal open onClose={() => setAddOpen(false)} onDone={() => { setAddOpen(false); load(); }} />}

      <BusyOverlay label={busy.busyLabel} />
    </div>
  );
}

function formatBatchDate(value) {
  return value ? new Intl.DateTimeFormat('th-TH', { timeZone: 'Asia/Bangkok', dateStyle: 'medium' }).format(new Date(value)) : 'ไม่ระบุวัน';
}

function InboundBadge({ asset }) {
  if (!asset.batchId) return <span className="inline-flex rounded-full bg-[var(--g100)] px-2 py-1 text-[11px] text-[var(--tsub)]">ไม่ระบุล็อต</span>;
  // Step 2: ย่อคอลัมน์ล็อต — ยึดความกว้างไม่เกิน 120px (รายละเอียดอยู่ใน title) กันตารางล้นขวาที่ 1024px
  return (
    <span
      title={`รับเข้า ${formatBatchDate(asset.receivedAt)}${asset.receivedAt ? ` (${inboundAge(asset.receivedAt)})` : ''}`}
      className="inline-flex max-w-[120px] items-center gap-x-1 whitespace-nowrap rounded-lg bg-[var(--blue-l)] px-2 py-1 text-[11px] text-[var(--blue)]"
    >
      <strong className="truncate">{asset.batchId}</strong>
      {asset.receivedAt && <span className="hidden 2xl:inline">· {formatBatchDate(asset.receivedAt)}</span>}
    </span>
  );
}

function BatchHistoryModal({ onClose, onSelectPO, onClosed }) {
  const [pos, setPOs] = useState([]);
  const [query, setQuery] = useState('');
  const [tab, setTab] = useState('active');
  const [loading, setLoading] = useState(true);
  const reload = useCallback(() => axios.get('/api/inbound-pos').then(({ data }) => setPOs(data || [])).catch((error) => { console.error('load inbound PO history', error); showToast('โหลดประวัติ PO ไม่สำเร็จ', { type: 'err' }); }).finally(() => setLoading(false)), []);
  useEffect(() => { reload(); }, [reload]);
  const filtered = pos.filter((po) => po.status === tab && [po.poNumber, po.supplier, ...po.batches.map((batch) => batch.batchId)].some((value) => String(value || '').toLowerCase().includes(query.trim().toLowerCase())));
  async function closePO(po) {
    try { await axios.post(`/api/inbound-pos/${encodeURIComponent(po.poNumber)}/close`); showToast(`ปิด PO ${po.poNumber} แล้ว`); await reload(); onClosed?.(); }
    catch (error) { showToast(error.response?.data?.error || 'ปิด PO ไม่สำเร็จ', { type: 'err' }); }
  }
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto bg-black/40 p-4 backdrop-blur-sm" onClick={onClose}>
      <section role="dialog" aria-modal="true" aria-labelledby="batch-history-title" className="my-6 w-full max-w-2xl overflow-hidden rounded-2xl bg-[var(--surface)] shadow-xl" onClick={(event) => event.stopPropagation()}>
        <header className="flex items-center justify-between gap-3 border-b border-[var(--g100)] px-4 py-3 sm:px-5"><div><h2 id="batch-history-title" className="text-[15px] font-bold text-[var(--text)]">ประวัติ PO รับเข้าทั้งหมด</h2><p className="text-[12px] text-[var(--tmuted)]">{pos.length} PO · รวมอุปกรณ์จากทุกล็อต</p></div><button onClick={onClose} aria-label="ปิดประวัติ PO" className="flex h-11 w-11 shrink-0 items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)]"><Icon name="close" size="sm" /></button></header>
        <div className="space-y-3 p-4 sm:p-5">
          <div className="grid grid-cols-2 gap-2"><button onClick={() => setTab('active')} className={`min-h-11 rounded-lg border text-[13px] font-semibold ${tab === 'active' ? 'border-[var(--blue)] bg-[var(--blue-l)] text-[var(--blue)]' : 'border-[var(--g200)]'}`}>กำลังดำเนินการ ({pos.filter((po) => po.status === 'active').length})</button><button onClick={() => setTab('closed')} className={`min-h-11 rounded-lg border text-[13px] font-semibold ${tab === 'closed' ? 'border-[var(--blue)] bg-[var(--blue-l)] text-[var(--blue)]' : 'border-[var(--g200)]'}`}>ปิดแล้ว ({pos.filter((po) => po.status === 'closed').length})</button></div>
          <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="ค้น PO / Batch ID / Supplier" className="h-11 w-full rounded-lg border border-[var(--g200)] bg-[var(--surface2)] px-3 text-[13px] focus:border-[var(--blue)] focus:outline-none" autoFocus />
          <div className="max-h-[60vh] space-y-2 overflow-y-auto">{loading ? <div className="py-8 text-center text-[13px] text-[var(--tmuted)]">กำลังโหลดประวัติ PO...</div> : filtered.length ? filtered.map((po) => <div key={po.key} className="flex flex-col gap-2 rounded-xl border border-[var(--g200)] p-3 sm:flex-row sm:items-center sm:justify-between"><div className="min-w-0"><div className="break-all text-[13px] font-bold text-[var(--blue)]">{po.poNumber ? `PO ${po.poNumber}` : po.batchId}</div><div className="text-[12px] text-[var(--tsub)]">{po.assetCount} ชิ้น · คงคลัง {po.stockCount} · {po.batchCount} ล็อต · รับเข้าล่าสุด {formatBatchDate(po.receivedAt)}</div>{po.batches.map((batch) => <div key={batch.batchId} className="break-all text-[11px] text-[var(--tmuted)]">{batch.batchId} ({batch.count})</div>)}{po.manuallyClosed && <div className="text-[11px] text-[var(--tmuted)]">ปิดโดย {po.closedBy || 'ผู้ใช้'} · {formatBatchDate(po.closedAt)}</div>}</div><div className="flex flex-col gap-2 sm:shrink-0"><button onClick={() => onSelectPO(po)} className="min-h-11 rounded-lg bg-[var(--blue)] px-3 text-[12px] font-semibold text-white">เลือกอุปกรณ์ทั้ง PO</button>{tab === 'active' && po.poNumber && <button onClick={() => closePO(po)} className="min-h-11 rounded-lg border border-[var(--g300)] px-3 text-[12px] font-semibold text-[var(--tsub)]">ปิด PO</button>}</div></div>) : <div className="py-8 text-center text-[13px] text-[var(--tmuted)]">{pos.length ? 'ไม่พบ PO ในรายการนี้' : 'ยังไม่มีข้อมูล PO รับเข้า'}</div>}</div>
        </div>
      </section>
    </div>
  );
}
function inboundAge(value) {
  const days = Math.max(0, Math.floor((Date.now() - new Date(value).getTime()) / 86400000));
  return days === 0 ? 'วันนี้' : `${days} วันที่แล้ว`;
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
  const busy = useBusy();

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
    if (!pn || !partName.trim()) return showToast('กรุณากรอก Part Number และชื่ออุปกรณ์', { type: 'warn' });
    const exists = parts.some((p) => p.partNumber === pn);
    await busy.run('กำลังเพิ่ม Part + Asset...', async () => {
      try {
        if (!exists) {
          const r = await axios.post('/api/add-part', { partNumber: pn, partName: partName.trim(), category, description, unit });
          if (r.data && r.data.error) { showToast('ไม่สำเร็จ: ' + r.data.error, { type: 'err' }); return; }
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
          showToast(`เพิ่ม Part + Asset สำเร็จ · ${assetId} · ${serialPreview}`);
          onDone();
        } else showToast('ไม่สำเร็จ: ' + (r.data.error || 'เพิ่ม Asset ไม่สำเร็จ'), { type: 'err' });
      } catch (e) {
        showToast('เกิดข้อผิดพลาด: ' + (e.response?.data?.error || e.message), { type: 'err' });
      }
    });
  }

  return (
    <AssetModal title="เพิ่ม Part + Asset ใหม่" submitLabel="เพิ่ม Part + Asset" onSubmit={submit} busy={busy.busy} busyLabel={busy.busyLabel} onClose={onClose}>
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
  const busy = useBusy();

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
    if (!partNumber.trim() || !partName.trim() || !n) return showToast('กรอกข้อมูลให้ครบ', { type: 'warn' });
    await busy.run('กำลังเพิ่ม Asset หลายชิ้น...', async () => {
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
          showToast(`เพิ่ม ${data.added} ชิ้นสำเร็จ · Serial ${data.firstSerial} ถึง ${data.lastSerial}`);
          onDone();
        } else showToast('เกิดข้อผิดพลาด: ' + (data.error || ''), { type: 'err' });
      } catch (e) {
        showToast('เกิดข้อผิดพลาด: ' + (e.response?.data?.error || e.message), { type: 'err' });
      }
    });
  }

  return (
    <AssetModal title="เพิ่ม Asset หลายชิ้น" submitLabel="เพิ่ม Asset" onSubmit={submit} busy={busy.busy} busyLabel={busy.busyLabel} onClose={onClose}>
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

function AssetModal({ title, submitLabel, onSubmit, busy, busyLabel, onClose, children }) {
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center p-4 bg-black/40 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
      <div className="w-full max-w-lg bg-[var(--surface)] rounded-2xl shadow-xl my-8" onClick={(e) => e.stopPropagation()}>
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
      <BusyOverlay label={busyLabel} />
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
