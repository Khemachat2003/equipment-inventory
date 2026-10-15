// ══════════════════════════════════════════════════════════════════════════════
// Asset — หน้า "ทะเบียนรายชิ้น" (Part/Asset Master List) — Enterprise Industrial Asset Management
// ──────────────────────────────────────────────────────────────────────────────
// B2 Strict Terminology — หน้าจอภาษาไทยล้วน ไม่มีศัพท์ระบบหลุดขึ้นหน้าจอ:
//   • สถานะข้อมูล "In Bundle" (backend: routes/bundle.js) → แสดง "อยู่ในชุดติดตั้ง" (StatusBadge)
//   • ตำแหน่งคลังแสดง "คลังกลาง" (buildLocation — utils/location.js)
//   • ช่องวันที่ = ThaiDateInput รูปแบบไทย "วว/ดด/ปปปป" (ไม่โชว์ mm/dd/yyyy ของ browser)
//   • Sidebar/ช่องค้นหาใช้คำ "รายการอุปกรณ์" แทน "Part"
// โครงหน้า: Top Hero (หัวหน้า + Primary Actions + KPI Glass) → PO Glass Cards
//   → Sidebar รายการอุปกรณ์ (จัดกลุ่มหมวดหมู่ + ชิปนับ Glass) → Unified Filter Bar
//   → Enterprise Data Table (แถว 52px · ตัวหนังสือ ink · Modern Pill สถานะ · Action Icons + Tooltip)
// ══════════════════════════════════════════════════════════════════════════════

import { useEffect, useState, useMemo, useCallback } from 'react';
import axios from 'axios';
import Icon from '../components/ui/Icon.jsx';
import StatusBadge from '../components/ui/StatusBadge.jsx';
import KpiGlass from '../components/ui/KpiGlass.jsx';
import ActionIcon from '../components/ui/ActionIcon.jsx';
import CodeBadge from '../components/ui/CodeBadge.jsx';
import ThaiDateInput from '../components/ui/ThaiDateInput.jsx';
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

// Min-Height แถวตาราง 52px — height บน <td> ทำงานเป็น "ความสูงขั้นต่ำ" ตาม CSS table spec (เนื้อหายาวกว่าจะขยาย)
const DT_TD = 'px-3 py-2.5 h-[52px] align-middle';

export default function Asset() {
  const [assets, setAssets] = useState([]);
  const [parts, setParts] = useState([]);
  const [categoryList, setCategoryList] = useState(CATEGORY_FALLBACK);
  const [categoryFilter, setCategoryFilter] = useState('');
  const [currentPart, setCurrentPart] = useState('');
  const [partSearch, setPartSearch] = useState('');
  // Unified Search — ช่องเดียวค้นได้ทั้ง ชื่อ/รหัส/AssetID/Serial และ ล็อต (Batch) / PO / ผู้ขาย
  const [search, setSearch] = useState('');
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
  const [receiveOpen, setReceiveOpen] = useState(false); // "รับเข้า" — AddDeviceModal แบบขยายแผนล็อต/PO ไว้ล่วงหน้า
  const [loading, setLoading] = useState(true);
  const busy = useBusy();

  // รับ ?status= จากหน้าแรก (การ์ดสรุปสถานะ 3 ใบ) → เปิดมากรองสถานะไว้ให้เลย
  // (อ่านจาก window.location ตรง ๆ แนวเดียวกับ Stock ?q= — ไม่ต้องพึ่ง Router context)
  const [statusFilter, setStatusFilter] = useState(() => {
    try { return new URLSearchParams(window.location.search).get('status') || ''; } catch { return ''; }
  });

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const [a, p, c, b] = await Promise.all([
        axios.get('/api/assets'),
        axios.get('/api/part-catalog'),
        axios.get('/api/categories'),
        axios.get('/api/inbound-pos?status=active&limit=12').catch(() => ({ data: [] })),
      ]);
      setAssets(a.data || []);
      setParts(p.data || []);
      if (c.data && c.data.length) setCategoryList(c.data);
      setRecentBatches(b.data || []);
    } catch (e) {
      console.error('load asset', e);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { load(); }, [load]);

  // Sidebar: จัดกลุ่มตาม Part Number (ค่าข้อมูลจริง — คงเดิมทุกประการ)
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
        partName: parts.find((p) => p.partNumber === key)?.partName || v.first.name || key,
        category: parts.find((p) => p.partNumber === key)?.category || '',
      }))
      .filter((x) => !q || x.partNumber.toLowerCase().includes(q) || x.partName.toLowerCase().includes(q))
      .sort((a, b) => a.partNumber.localeCompare(b.partNumber, 'th'));
  }, [assets, parts, partSearch]);

  // Sidebar: แบ่งกลุ่มด้วยเส้นคั่นหมวดหมู่บาง ๆ (B2) — กลุ่ม "ไม่ระบุหมวดหมู่" ไปท้ายสุด
  const partCategoryGroups = useMemo(() => {
    const groups = [];
    const byLabel = {};
    partGroups.forEach((g) => {
      const label = categoryLabel(g.category) || 'ไม่ระบุหมวดหมู่';
      if (!byLabel[label]) { byLabel[label] = { label, total: 0, items: [] }; groups.push(byLabel[label]); }
      byLabel[label].items.push(g);
      byLabel[label].total += g.count;
    });
    groups.sort((a, b) => {
      if (a.label === 'ไม่ระบุหมวดหมู่') return 1;
      if (b.label === 'ไม่ระบุหมวดหมู่') return -1;
      return a.label.localeCompare(b.label, 'th');
    });
    return groups;
  }, [partGroups]);

  // Table: filter by part + unified search + category + ช่วงวันรับเข้า
  const catOf = useCallback((a) => parts.find((p) => p.partNumber === (a.partNumber || a.code))?.category || '', [parts]);
  const filtered = useMemo(() => {
    const k = search.trim().toLowerCase();
    return assets.filter((a) => {
      if (statusFilter && !(a.status || '').includes(statusFilter)) return false;
      if (currentPart && a.partNumber !== currentPart) return false;
      if (categoryFilter && catOf(a) !== categoryFilter) return false;
      if (!matchesInboundFilters(a, { from: receivedFrom, to: receivedTo })) return false;
      if (!k) return true;
      // Unified: เจอจากข้อมูลอุปกรณ์ หรือจาก ล็อต/PO/ผู้ขาย อย่างใดอย่างหนึ่ง
      const assetHit = [a.assetId, a.code, a.name, a.serialNumber].some((v) => String(v || '').toLowerCase().includes(k));
      const inboundHit = matchesInboundFilters(a, { query: search });
      return assetHit || inboundHit;
    });
  }, [assets, statusFilter, currentPart, categoryFilter, catOf, search, receivedFrom, receivedTo]);

  const paged = filtered.slice((page - 1) * pageSize, page * pageSize);

  useEffect(() => { setPage(1); }, [currentPart, categoryFilter, search, receivedFrom, receivedTo, pageSize, statusFilter]);

  const currentPartName = currentPart ? (parts.find((p) => p.partNumber === currentPart)?.partName || currentPart) : '';
  const stats = useMemo(() => {
    const usable = filtered.filter((a) => (a.status || '').includes('ใช้งานได้')).length;
    const repair = filtered.filter((a) => (a.status || '').includes('ซ่อม')).length;
    const bundled = filtered.filter((a) => !!a.bundleId).length;
    return { total: filtered.length, usable, repair, bundled };
  }, [filtered]);

  // เลือกทั้งหมดในหน้านี้ (checkbox หัวตาราง)
  const pageSerials = paged.map((a) => a.serialNumber).filter(Boolean);
  const allPageSelected = pageSerials.length > 0 && pageSerials.every((s) => selectedSerials.includes(s));
  function toggleAllPage() {
    setSelectedSerials((prev) => {
      if (pageSerials.every((s) => prev.includes(s))) return prev.filter((s) => !pageSerials.includes(s));
      return [...new Set([...prev, ...pageSerials])];
    });
  }

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
    setCurrentPart(''); setCategoryFilter(''); setReceivedFrom(''); setReceivedTo('');
    setSearch(po.poNumber || po.batchId || '');
    setSelectedSerials(serials);
    showToast(`เลือกอุปกรณ์ PO ${po.poNumber || po.batchId} จำนวน ${serials.length} ชิ้นแล้ว`);
  }

  function exportCSV() {
    const heads = ['Asset ID', 'Code', 'Name', 'Serial', 'Batch ID', 'Received At', 'PO Number', 'Supplier', 'Status', 'อยู่ในชุด', 'ตำแหน่ง (ฟาร์ม › โรงเรือน › จุดติดตั้ง)', 'User', 'Category'];
    const rows = [heads];
    filtered.forEach((a) => rows.push([
      a.assetId, a.code, a.name, a.serialNumber, a.batchId || '', a.receivedAt || '', a.poNumber || '', a.supplier || '',
      a.status || '',
      a.bundleName || a.bundleId || '',
      buildLocation({ siteName: a.siteName, houseName: a.houseName, houseId: a.houseId, location: a.location, bundleId: a.bundleId }).full,
      a.user, catOf(a),
    ]));
    dlCSV(`asset_export_${stamp()}.csv`, rows);
  }

  return (
    <div className="space-y-4">
      {!transfer && !addOpen && !receiveOpen && <ToastHost />}

      {/* ═══ Top Hero Bar — หัวหน้า + Primary Action Buttons + KPI Glassmorphic ═══ */}
      <section className="rounded-2xl border border-[var(--border)] bg-[var(--surface)]/85 backdrop-blur-md shadow-[var(--sh-sm)] p-3.5 sm:p-4 space-y-3.5">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-2.5 min-w-0">
            <span className="flex items-center justify-center w-10 h-10 rounded-xl bg-[var(--ink)] text-white shrink-0 shadow-[var(--sh-xs)]"><Icon name="devices" size="md" /></span>
            <div className="min-w-0">
              <h1 className="text-[16px] sm:text-[18px] font-extrabold text-[var(--ink)] leading-tight truncate">ทะเบียนรายชิ้น</h1>
              <p className="text-[11.5px] text-[var(--tsub)] truncate">ติดตามอุปกรณ์ทุกชิ้น — สถานะ ตำแหน่ง และประวัติ</p>
            </div>
          </div>
          {/* Primary Action Buttons — เด่นชัด มีไอคอนนำหน้าสะดุดตา */}
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => setAddOpen(true)}
              className="flex items-center gap-1.5 h-10 px-4 rounded-xl bg-[var(--blue)] text-white text-[13px] font-bold shadow-[0_2px_10px_var(--blue-glow)] transition hover:bg-[var(--blue-d)] hover:shadow-[var(--sh-md)] active:translate-y-px"
            >
              <Icon name="add" size="sm" /> เพิ่มอุปกรณ์ใหม่
            </button>
            <button
              onClick={() => setReceiveOpen(true)}
              className="flex items-center gap-1.5 h-10 px-4 rounded-xl bg-[var(--emerald)] text-white text-[13px] font-bold shadow-[0_2px_10px_rgba(0,200,150,.25)] transition hover:bg-[var(--emerald-d)] hover:shadow-[var(--sh-md)] active:translate-y-px"
            >
              <Icon name="inbox" size="sm" /> รับเข้า
            </button>
            <button
              onClick={load}
              title="รีเฟรชข้อมูล"
              aria-label="รีเฟรชข้อมูล"
              className="flex items-center justify-center w-10 h-10 rounded-xl border border-[var(--g200)] bg-[var(--surface)] text-[var(--tsub)] transition hover:text-[var(--ink)] hover:border-[var(--g300)] hover:shadow-[var(--sh-xs)] active:translate-y-px"
            >
              <Icon name="refresh" size="sm" />
            </button>
          </div>
        </div>

        {/* KPI Glassmorphic Cards — ตัวเลข Hero สะดุดตา อ่านได้จากระยะ */}
        <div className="flex flex-wrap items-stretch gap-2">
          <KpiGlass testId="kpi-total" icon="inventory_2" tone="slate" label="อุปกรณ์ทั้งหมด" value={stats.total} unit="ชิ้น" />
          <KpiGlass testId="kpi-usable" icon="check_circle" tone="green" label="ใช้งานได้" value={stats.usable} unit="ชิ้น" />
          <KpiGlass testId="kpi-bundled" icon="folder_open" tone="blue" label="อยู่ในชุดติดตั้ง" value={stats.bundled} unit="ชิ้น" />
          {stats.repair > 0 && <KpiGlass testId="kpi-repair" icon="build" tone="amber" label="ส่งซ่อม" value={stats.repair} unit="ชิ้น" />}
        </div>

        {/* ชิปตัวกรองที่กำลังใช้งาน — รายการอุปกรณ์ที่เลือก / สถานะจาก deep-link หน้าแรก */}
        {(currentPartName || statusFilter) && (
          <div className="flex flex-wrap items-center gap-2">
            {currentPartName && (
              <span className="flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-[var(--g100)] border border-[var(--g200)] text-[12px] font-semibold text-[var(--ink)]">
                <Icon name="inventory_2" size="xs" /> {currentPartName}
                <button
                  onClick={() => setCurrentPart('')}
                  title="ล้างตัวกรองรายการอุปกรณ์"
                  aria-label="ล้างตัวกรองรายการอุปกรณ์"
                  className="flex items-center justify-center w-5 h-5 rounded-full hover:bg-[var(--g200)] focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--blue-glow)]"
                >
                  <Icon name="close" size="xs" />
                </button>
              </span>
            )}
            {statusFilter && (
              <span className="flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-[var(--blue-l)] border border-[var(--blue-b)] text-[12px] font-semibold text-[var(--blue-d)]">
                <Icon name="filter_alt" size="xs" /> สถานะ: {statusFilter}
                <button
                  onClick={() => setStatusFilter('')}
                  title="ล้างตัวกรองสถานะ"
                  aria-label="ล้างตัวกรองสถานะ"
                  className="flex items-center justify-center w-5 h-5 rounded-full hover:bg-[var(--blue-b)]/60 focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--blue-glow)]"
                >
                  <Icon name="close" size="xs" />
                </button>
              </span>
            )}
          </div>
        )}
      </section>

      {/* แถบเลือก (Bulk select) — โชว์เมื่อมีการเลือกชิ้น */}
      {selectedSerials.length > 0 && (
        <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border border-[var(--blue-b)] bg-[var(--blue-l)]/80 backdrop-blur-sm px-3.5 py-2.5 shadow-[var(--sh-xs)]">
          <span className="flex items-center gap-1.5 text-[13px] font-semibold text-[var(--blue-d)]"><Icon name="check_circle" size="xs" /> เลือกแล้ว {selectedSerials.length} ชิ้น</span>
          <div className="flex gap-2">
            <button onClick={() => setSelectedSerials([])} className="min-h-10 px-3 rounded-lg border border-[var(--g300)] bg-[var(--surface)] text-[13px] text-[var(--tsub)] hover:text-[var(--ink)]">ล้างที่เลือก</button>
            <button onClick={startBulkTransfer} className="min-h-10 px-3 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)]">ย้ายที่เลือก</button>
          </div>
        </div>
      )}

      {/* PO ที่กำลังดำเนินการ — Glassmorphic PO Cards (กดเลือกอุปกรณ์ของ PO นั้นทันที) */}
      <section className="rounded-2xl border border-[var(--border)] bg-[var(--surface)]/85 backdrop-blur-md shadow-[var(--sh-sm)] p-3.5 sm:p-4 space-y-2.5">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center gap-2 min-w-0">
            <span className="flex items-center justify-center w-8 h-8 rounded-lg bg-[var(--blue-l)] border border-[var(--blue-b)] text-[var(--blue)] shrink-0"><Icon name="inbox" size="sm" /></span>
            <span className="text-[13px] font-bold text-[var(--ink)]">PO ที่กำลังดำเนินการ</span>
          </div>
          <button onClick={() => setBatchHistoryOpen(true)} className="min-h-11 px-3 rounded-xl border border-[var(--g200)] bg-[var(--surface)] text-[12.5px] font-semibold text-[var(--blue)] hover:bg-[var(--blue-l)] transition-colors">ดูประวัติ PO ทั้งหมด</button>
        </div>
        {recentBatches.length > 0 ? (
          <div className="flex gap-2.5 overflow-x-auto pb-1">
            {recentBatches.map((po) => (
              <button
                key={po.key}
                onClick={() => selectPO(po)}
                className="relative min-h-11 shrink-0 overflow-hidden rounded-xl border border-[var(--g200)] bg-white/70 backdrop-blur-sm px-3.5 py-2 text-left shadow-[var(--sh-xs)] transition-all hover:border-[var(--blue)] hover:shadow-[var(--sh-md)] hover:-translate-y-px"
              >
                <span aria-hidden="true" className="absolute inset-y-0 left-0 w-[3px] bg-[var(--blue)]" />
                <span className="block text-[12.5px] font-bold text-[var(--blue)]">{po.poNumber ? `PO ${po.poNumber}` : po.batchId}</span>
                <span className="block text-[11px] text-[var(--tsub)] mt-0.5">คงคลัง {po.stockCount}/{po.assetCount} ชิ้น · {po.batchCount} ล็อต</span>
              </button>
            ))}
          </div>
        ) : (
          <div className="text-[12px] text-[var(--tmuted)]">ไม่มี PO ที่ยังรอดำเนินการ</div>
        )}
      </section>

      <div className="grid grid-cols-1 xl:grid-cols-[248px_minmax(0,1fr)] gap-4 items-start">
        {/* ═══ Sidebar รายการอุปกรณ์ (จอใหญ่) — แบ่งหมวดหมู่ด้วยเส้นบางคมชัด + ชิปนับแบบ Glass ═══ */}
        <div data-testid="part-sidebar" className="hidden xl:block rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden xl:max-h-[75vh] xl:sticky xl:top-[var(--topbar-h)]">
          <div className="px-3 py-3 border-b border-[var(--border)] flex items-center justify-between gap-2">
            <span className="flex items-center gap-2 min-w-0 text-[12.5px] font-bold text-[var(--ink)]">
              <span className="flex items-center justify-center w-7 h-7 rounded-lg bg-[var(--ink)] text-white shrink-0"><Icon name="devices" size="xs" /></span>
              รายการอุปกรณ์
            </span>
            <CountChip active={!!currentPart}>{assets.length}</CountChip>
          </div>
          <div className="p-2 border-b border-[var(--g100)]">
            <input
              value={partSearch}
              onChange={(e) => setPartSearch(e.target.value)}
              placeholder="ค้นหารายการอุปกรณ์..."
              className="w-full h-9 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px] text-[var(--ink)] placeholder:text-[var(--tmuted)] focus:outline-none focus:border-[var(--blue)] focus:ring-2 focus:ring-[var(--blue-glow)]"
            />
          </div>
          <div className="overflow-y-auto xl:max-h-[58vh] p-1.5">
            <PartItem
              icon="format_list_bulleted"
              label="ทุกรายการอุปกรณ์"
              count={assets.length}
              active={currentPart === ''}
              onClick={() => setCurrentPart('')}
            />
            {partCategoryGroups.map((group) => (
              <div key={group.label}>
                <div className="mt-1.5 flex items-center justify-between gap-2 border-t border-[var(--border)] px-2.5 pt-2 pb-1">
                  <span className="text-[10px] font-bold uppercase tracking-wider text-[var(--tmuted)] truncate">{group.label}</span>
                  <CountChip>{group.total}</CountChip>
                </div>
                {group.items.map((g) => (
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
            ))}
          </div>
        </div>

        {/* ═══ Table Card ═══ */}
        <div className="rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
          {/* Mobile part selector */}
          <div className="xl:hidden flex items-center gap-2 p-2.5 border-b border-[var(--g100)]">
            <Icon name="devices" size="sm" />
            <select
              value={currentPart}
              onChange={(e) => setCurrentPart(e.target.value)}
              aria-label="เลือกรายการอุปกรณ์"
              className="flex-1 h-9 px-2.5 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px] text-[var(--ink)]"
            >
              <option value="">ทุกรายการอุปกรณ์ ({assets.length})</option>
              {partGroups.map((g) => (
                <option key={g.partNumber} value={g.partNumber}>{g.partNumber} ({g.count})</option>
              ))}
            </select>
          </div>

          {/* ═══ Unified Search & Date Filter Bar — ยุบช่องค้นหาที่กระจัดกระจายให้เหลือแถบเดียวเรียบร้อย ═══ */}
          <div className="flex flex-wrap items-center gap-2.5 p-3 sm:p-3.5 border-b border-[var(--border)]">
            <div className="relative flex-1 min-w-[230px]">
              <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[var(--tmuted)] pointer-events-none"><Icon name="search" size="sm" /></span>
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="ค้นหาชื่ออุปกรณ์ · รหัส · Serial · ล็อต / PO / ผู้ขาย"
                className="w-full h-10 pl-9 pr-8 rounded-xl border border-[var(--g200)] bg-[var(--surface2)] text-[13px] text-[var(--ink)] placeholder:text-[var(--tmuted)] focus:outline-none focus:border-[var(--blue)] focus:ring-2 focus:ring-[var(--blue-glow)]"
              />
              {search && (
                <button onClick={() => setSearch('')} title="ล้างคำค้น" aria-label="ล้างคำค้น" className="absolute right-2 top-1/2 -translate-y-1/2 w-6 h-6 flex items-center justify-center rounded-md text-[var(--tmuted)] hover:bg-[var(--g100)] hover:text-[var(--ink)]"><Icon name="close" size="xs" /></button>
              )}
            </div>
            <select
              value={categoryFilter}
              onChange={(e) => setCategoryFilter(e.target.value)}
              aria-label="กรองตามหมวดหมู่"
              className="h-10 px-2.5 rounded-xl border border-[var(--g200)] bg-[var(--surface2)] text-[12.5px] text-[var(--ink)] min-w-[150px] focus:outline-none focus:border-[var(--blue)]"
            >
              <option value="">ทุกหมวดหมู่</option>
              {categoryList.map((c) => <option key={c.name} value={c.name}>{c.label}</option>)}
            </select>
            <ThaiDateInput label="รับเข้าตั้งแต่" ariaLabel="รับเข้าตั้งแต่" value={receivedFrom} onChange={setReceivedFrom} />
            <ThaiDateInput label="ถึง" ariaLabel="รับเข้าถึง" value={receivedTo} onChange={setReceivedTo} />
            <button
              onClick={exportCSV}
              title="ส่งออก CSV ทั้งหมดที่กรองอยู่"
              className="flex items-center gap-1.5 h-10 px-3.5 rounded-xl border border-[var(--g200)] bg-[var(--surface)] text-[12.5px] font-semibold text-[var(--tsub)] hover:text-[var(--ink)] hover:border-[var(--g300)] transition-colors shrink-0"
            >
              <Icon name="download" size="sm" /> CSV
            </button>
          </div>

          {/* ═══ Enterprise Data Table (จอใหญ่) — แถว 52px · ตัวหนังสือ ink · Modern Pill สถานะ · Action Icons + Tooltip ═══ */}
          <div className="hidden md:block overflow-x-auto">
            <table className="w-full min-w-[900px] text-[12.5px]">
              <thead>
                <tr className="text-left text-[11px] font-bold text-[var(--tmuted)] bg-[var(--surface2)] border-b border-[var(--border)]">
                  <th className="w-10 px-3 py-3">
                    <input
                      type="checkbox"
                      aria-label="เลือกทั้งหมดในหน้านี้"
                      checked={allPageSelected}
                      onChange={toggleAllPage}
                      className="h-[18px] w-[18px] accent-[var(--blue)] cursor-pointer align-middle"
                    />
                  </th>
                  <th className="px-3 py-3 whitespace-nowrap">รายการอุปกรณ์</th>
                  <th className="px-3 py-3 whitespace-nowrap">ตำแหน่ง / ชุดติดตั้ง</th>
                  <th className="px-3 py-3 whitespace-nowrap">Serial</th>
                  <th className="px-3 py-3 whitespace-nowrap">ล็อตรับเข้า</th>
                  <th className="px-3 py-3 whitespace-nowrap">สถานะ</th>
                  <th className="hidden px-3 py-3 whitespace-nowrap 2xl:table-cell">ผู้ใช้</th>
                  <th className="sticky right-0 z-20 border-l border-[var(--g200)] bg-[var(--surface2)] px-3 py-3 text-center font-bold whitespace-nowrap shadow-[-10px_0_10px_-10px_rgba(15,23,42,0.25)]">จัดการ</th>
                </tr>
              </thead>
              <tbody>
                {loading && <tr><td colSpan={8} className="text-center py-10 text-[var(--tmuted)]">กำลังโหลด...</td></tr>}
                {!loading && paged.length === 0 && (
                  <tr>
                    <td colSpan={8} className="text-center py-12 text-[var(--tmuted)]">
                      <div className="flex justify-center mb-2 text-[var(--tmuted)]"><Icon name="inventory_2" size="2xl" /></div>
                      ไม่พบอุปกรณ์ในรายการนี้
                    </td>
                  </tr>
                )}
                {paged.map((a) => (
                  <tr key={a.serialNumber + a.assetId} className="group border-b border-[var(--border)] transition-colors hover:bg-[var(--surface2)]">
                    <td className="w-10 px-3 py-2.5 h-[52px] align-middle">
                      <input
                        aria-label={`เลือก ${a.serialNumber}`}
                        type="checkbox"
                        className="h-[18px] w-[18px] accent-[var(--blue)] cursor-pointer"
                        checked={selectedSerials.includes(a.serialNumber)}
                        onChange={() => toggleSelected(a)}
                      />
                    </td>
                    <td className={DT_TD}>
                      {/* ชื่ออุปกรณ์ = ตัวหนาหลัก (ink) · รหัส/AssetID = Badge เทารอง (B2) */}
                      <div className="flex items-center gap-2.5 min-w-[210px]">
                        <span title={categoryLabel(catOf(a))} className="flex items-center justify-center w-8 h-8 rounded-lg bg-[var(--g100)] border border-[var(--g200)] text-[var(--tsub)] shrink-0"><Icon name={categoryIcon(catOf(a))} size="xs" /></span>
                        <div className="min-w-0">
                          <div className="text-[13px] font-bold text-[var(--ink)] truncate max-w-[170px]" title={a.name}>{a.name}</div>
                          <div className="mt-0.5 flex items-center gap-1.5 flex-wrap">
                            <CodeBadge>{a.code}</CodeBadge>
                            <span className="font-mono text-[10px] text-[var(--tmuted)]"><span className="block max-w-[70px] truncate">{a.assetId}</span></span>
                          </div>
                        </div>
                      </div>
                    </td>
                    <td className={DT_TD}>
                      {/* ตำแหน่งฟาร์ม/โรงเรือน = Primary Visual Indicator · แถบ "อยู่ในชุดติดตั้ง" อยู่เหนือ pill */}
                      <div className="min-w-0 max-w-[190px]">
                        {a.bundleId && (
                          <span
                            title={`อยู่ในชุดติดตั้ง ${a.bundleName || a.bundleId}`}
                            className="mb-1 inline-flex max-w-full items-center gap-1 rounded-md bg-[var(--blue-l)] px-1.5 py-0.5 text-[10px] font-bold text-[var(--blue)] border border-[var(--blue-b)]"
                          >
                            <Icon name="folder_open" size="xs" className="flex-shrink-0" />
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
                    <td className={DT_TD}>
                      <span className="font-mono text-[11.5px] font-medium text-[var(--ink)] block max-w-[110px] truncate" title={a.serialNumber}>{a.serialNumber}</span>
                    </td>
                    <td className={DT_TD}><InboundBadge asset={a} /></td>
                    <td className={DT_TD}><StatusBadge status={a.status} /></td>
                    <td className="hidden px-3 py-2.5 h-[52px] align-middle text-[var(--tsub)] 2xl:table-cell">{a.user || '—'}</td>
                    <td className="sticky right-0 z-10 border-l border-[var(--g200)] bg-[var(--surface)] px-2 py-2.5 h-[52px] align-middle shadow-[-10px_0_10px_-10px_rgba(15,23,42,0.25)] group-hover:bg-[var(--surface2)]">
                      <div className="flex items-center justify-center gap-0.5">
                        <ActionIcon icon="description" label="ดูประวัติการเคลื่อนไหว" tone="blue" onClick={() => openHistory(a.serialNumber)} />
                        <ActionIcon icon="qr_code" label="พิมพ์ฉลาก QR" href={`/qr?serial=${encodeURIComponent(a.serialNumber)}`} target="_blank" />
                        <ActionIcon icon="local_shipping" label="โอนย้ายอุปกรณ์" tone="blue" onClick={() => doTransfer(a)} />
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
              <div className="text-center py-10 text-[var(--tmuted)] text-[13px]">ไม่พบอุปกรณ์ในรายการนี้</div>
            )}
            {paged.map((a) => (
              <div key={a.serialNumber + a.assetId} className="p-3.5 space-y-2.5">
                <div className="flex items-start gap-2.5">
                  <input aria-label={`เลือก ${a.serialNumber}`} className="mt-2 h-5 w-5 shrink-0 accent-[var(--blue)]" type="checkbox" checked={selectedSerials.includes(a.serialNumber)} onChange={() => toggleSelected(a)} />
                  <span className="flex items-center justify-center w-8 h-8 rounded-lg shrink-0 bg-[var(--g100)] border border-[var(--g200)] text-[var(--tsub)]" title={categoryLabel(catOf(a))}>
                    <Icon name={categoryIcon(catOf(a))} size="sm" />
                  </span>
                  <div className="flex-1 min-w-0">
                    <div className="text-[13px] font-bold text-[var(--ink)] leading-snug">{a.name}</div>
                    <div className="mt-0.5 flex items-center gap-1.5 flex-wrap">
                      <CodeBadge>{a.code}</CodeBadge>
                      <span className="font-mono text-[10px] text-[var(--tmuted)]">{a.assetId}</span>
                    </div>
                  </div>
                  <StatusBadge status={a.status} />
                </div>
                {/* ตำแหน่งฟาร์ม/โรงเรือน = ข้อมูลแรกที่ช่างต้องเห็นบนการ์ด (Primary Indicator) */}
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
                  <div className="text-[var(--tsub)]">Serial: <span className="font-mono font-semibold text-[var(--ink)]">{a.serialNumber}</span></div>
                  <InboundBadge asset={a} />
                  {a.bundleId && (
                    <div className="flex items-center gap-1.5">
                      <span className="text-[var(--tsub)]">อยู่ในชุดติดตั้ง:</span>
                      <span className="inline-flex items-center gap-1 rounded-md bg-[var(--blue-l)] border border-[var(--blue-b)] px-1.5 py-0.5 text-[11px] font-bold text-[var(--blue)]">
                        <Icon name="folder_open" size="xs" /> {a.bundleName || a.bundleId}
                      </span>
                    </div>
                  )}
                  {a.user && <div className="text-[var(--tsub)]">ผู้ใช้: <span className="font-medium text-[var(--ink)]">{a.user}</span></div>}
                </div>
                <div className="pl-[42px] flex items-center gap-2">
                  <button onClick={() => openHistory(a.serialNumber)} className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)]">
                    <Icon name="description" size="xs" /> ประวัติ
                  </button>
                  <a href={`/qr?serial=${encodeURIComponent(a.serialNumber)}`} target="_blank" rel="noreferrer" className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)]">
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

      {/* เพิ่มอุปกรณ์ใหม่ / รับเข้า — โฟลว์เดียว (เลือก/สร้างรายการอุปกรณ์ → Serial อัตโนมัติ)
          "รับเข้า" เปิดมาพร้อมแผนล็อต/PO ขยายไว้ล่วงหน้า (presetAdvanced) */}
      {addOpen && <AddDeviceModal open onClose={() => setAddOpen(false)} onDone={() => { setAddOpen(false); load(); }} />}
      {receiveOpen && <AddDeviceModal open presetAdvanced onClose={() => setReceiveOpen(false)} onDone={() => { setReceiveOpen(false); load(); }} />}

      <BusyOverlay label={busy.busyLabel} />
    </div>
  );
}

function formatBatchDate(value) {
  return value ? new Intl.DateTimeFormat('th-TH', { timeZone: 'Asia/Bangkok', dateStyle: 'medium' }).format(new Date(value)) : 'ไม่ระบุวัน';
}

function InboundBadge({ asset }) {
  if (!asset.batchId) return <span className="inline-flex rounded-full bg-[var(--g100)] px-2 py-1 text-[11px] text-[var(--tsub)]">ไม่ระบุล็อต</span>;
  // ย่อคอลัมน์ล็อต — ยึดความกว้างไม่เกิน 120px (รายละเอียดอยู่ใน title) กันตารางล้นขวาที่ 1024px
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

// CountChip — ชิปนับจำนวนแบบ Glassmorphic (พื้นโปร่งแสง + ขอบบาง + เงาเบา) — Sidebar รายการอุปกรณ์
function CountChip({ children, active = false }) {
  return (
    <span className={`inline-flex items-center justify-center min-w-[22px] px-1.5 py-0.5 rounded-full border text-[10px] font-bold tabular-nums backdrop-blur-sm whitespace-nowrap shadow-[var(--sh-xs)] ${active ? 'border-[var(--blue-b)] bg-[var(--blue-l)] text-[var(--blue)]' : 'border-[var(--g200)] bg-white/70 text-[var(--g600)]'}`}>
      {children}
    </span>
  );
}

// PartItem — รายการอุปกรณ์ 1 กลุ่มใน Sidebar (ชื่อ Part Number = ค่าข้อมูลจริง · จำนวนเป็นชิป Glass)
function PartItem({ icon, label, sub, count, active, onClick }) {
  return (
    <button
      onClick={onClick}
      className={`w-full flex items-center gap-2 px-2.5 py-2 rounded-lg text-left transition-colors ${active ? 'bg-[var(--blue-l)]' : 'hover:bg-[var(--surface2)]'}`}
    >
      <span className={`flex items-center justify-center w-6 h-6 rounded-md border shrink-0 ${active ? 'bg-[var(--blue)] border-[var(--blue)] text-white' : 'bg-[var(--g100)] border-[var(--g200)] text-[var(--tsub)]'}`}>
        <Icon name={icon} size="xs" />
      </span>
      <span className="flex-1 min-w-0">
        <span className={`block text-[12px] truncate ${active ? 'font-bold text-[var(--blue)]' : 'font-semibold text-[var(--ink)]'}`}>{label}</span>
        {sub && <span className="block text-[10px] text-[var(--tmuted)] truncate">{sub}</span>}
      </span>
      <CountChip active={active}>{count}</CountChip>
    </button>
  );
}

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
