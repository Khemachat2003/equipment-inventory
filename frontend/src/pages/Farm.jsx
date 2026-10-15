import { useEffect, useState, useMemo, useCallback } from 'react';
import axios from 'axios';
import { useNavigate } from 'react-router-dom';
import Icon from '../components/ui/Icon.jsx';
import StatusBadge from '../components/ui/StatusBadge.jsx';
import TransferModal from '../components/TransferModal.jsx';
import LocationPath from '../components/ui/LocationPath.jsx';
import { useBusy, BusyOverlay } from '../components/ui/Busy.jsx';
import { suggestSiteId, suggestHouseId } from '../utils/farmId.js';
import { buildLocation } from '../utils/location.js';

const TYPE_ICONS = { 'สัตว์ปีก': 'egg', 'สัตว์บก': 'pets', 'สุกร': 'agriculture', 'อื่นๆ': 'category', 'ไม่ระบุ': 'help' };
const FARM_TYPES = ['สัตว์ปีก', 'สัตว์บก', 'สุกร', 'อื่นๆ'];

// B2 Clean Titles — ซ่อนรหัสชุด "(BDL-xxx)" ออกจาก "ชื่อหลัก" ของการ์ด (โชว์เฉพาะชื่อไทย อ่านง่าย)
// ใช้ตอน "แสดงผล" เท่านั้น — ค่าจริง (ค้นหา/ฟิลเตอร์/เปิดฟาร์ม) ยังอ้างชื่อเดิมเสมอ ไม่กระทบข้อมูล
const BUNDLE_CODE_RE = /\s*\(\s*BDL[^)]*\)/gi;
function cleanBundleTitle(value) {
  const s = String(value || '');
  const cleaned = s.replace(BUNDLE_CODE_RE, '').trim();
  return cleaned || s.trim();
}

// ═══ Enterprise Data Table helpers — มุมมอง "ตารางรายชิ้น" (B2 Strict Terminology + Modern Data Table) ═══

// สีสถานะของ "ชุดติดตั้ง" — เขียว = ติดตั้งแล้ว · น้ำเงิน = อยู่ในคลัง · ส้ม = ซ่อมบำรุง (ภาษาเดียวกับหน้า Bundle — BTONES)
const BUNDLE_TONES = { 'In Stock': 'blue', Deployed: 'green', Maintenance: 'amber' };

// สีสถานะของ "รายชิ้น" — เขียว = ใช้งานได้ · ส้ม = ส่งซ่อม · เทา = ชำรุด/สูญหาย
function farmAssetTone(status) {
  const s = String(status || '');
  if (s.includes('ซ่อม')) return 'amber';
  if (s.includes('ชำรุด') || s.includes('สูญ')) return 'gray';
  return 'green';
}

// CodeBadge — รหัสรอง (BDL-xxx / รหัสอุปกรณ์) เป็น Badge เทารองขนาดเล็ก — B2: ชื่อไทยเป็นตัวหนาหลัก รหัสย่อเป็นรอง
function CodeBadge({ children }) {
  return (
    <span className="inline-flex items-center whitespace-nowrap rounded-md border border-[var(--border)] bg-[var(--g100)] px-1.5 py-0.5 font-mono text-[10px] font-semibold text-[var(--tmuted)]">
      {children}
    </span>
  );
}

// KpiGlass — KPI Summary Card แบบ Glassmorphic (ภาษาเดียวกับ KPI Cards หน้า Bundle):
// พื้นโปร่งแสงเบลอ --surface/80 + backdrop-blur-md + ขอบบาง var(--border) + แถบไล่สี tone ด้านบน 3px + ตัวเลข Hero tabular-nums
const KPI_ACCENTS = {
  blue: 'from-[var(--blue)]',
  green: 'from-[var(--emerald)]',
  amber: 'from-[var(--amber)]',
  red: 'from-[var(--red)]',
};
const KPI_BOX = {
  blue: 'bg-[var(--blue-l)] text-[var(--blue)]',
  green: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]',
  amber: 'bg-[var(--amber-l)] text-[var(--amber-d)]',
  red: 'bg-[var(--red-l)] text-[var(--red)]',
};
function KpiGlass({ icon, tone = 'blue', label, value, unit }) {
  return (
    <div className="relative flex select-none items-center gap-2.5 min-w-0 overflow-hidden rounded-xl border border-[var(--border)] bg-[var(--surface)]/80 px-3 py-2 backdrop-blur-md shadow-[var(--sh-sm)] hover:shadow-[var(--sh-md)] hover:-translate-y-px transition-all">
      <span aria-hidden="true" className={`absolute inset-x-0 top-0 h-[3px] bg-gradient-to-r ${KPI_ACCENTS[tone] || KPI_ACCENTS.blue} to-transparent`} />
      <span className={`flex items-center justify-center w-8 h-8 rounded-lg shrink-0 ${KPI_BOX[tone] || KPI_BOX.blue}`}>
        <Icon name={icon} size="sm" />
      </span>
      <div className="leading-none min-w-0">
        <div className="text-[10px] font-semibold text-[var(--tmuted)] whitespace-nowrap">{label}</div>
        <div className="mt-1 text-[20px] font-extrabold text-[var(--text)] tabular-nums leading-none">
          {value} <span className="text-[11px] font-bold text-[var(--tsub)] whitespace-nowrap">{unit}</span>
        </div>
      </div>
    </div>
  );
}

// ActionIcon — Clean Action Group: ปุ่มไอคอนโปร่งตัวเบา + Hover Tooltip สวยงาม (แถบเม็ดเข้ม var(--ink) + เงานุ่ม)
// tooltip วาง "แนวนอนซ้ายของไอคอน" ภายในแถวเสมอ → ไม่โดน clip จาก overflow-x-auto ของ scroll container
function ActionIcon({ icon, label, href, target, onClick, tone = 'ink' }) {
  const btnCls = tone === 'blue'
    ? 'text-[var(--blue)] hover:bg-[var(--blue-l)]'
    : 'text-[var(--tsub)] hover:bg-[var(--g100)] hover:text-[var(--ink)]';
  const cls = `relative z-10 flex h-8 w-8 items-center justify-center rounded-lg transition-colors ${btnCls}`;
  return (
    <span className="group/action relative flex items-center justify-center">
      {href ? (
        <a href={href} target={target} rel={target ? 'noreferrer' : undefined} aria-label={label} className={cls}>
          <Icon name={icon} size="sm" />
        </a>
      ) : (
        <button type="button" aria-label={label} onClick={onClick} className={cls}>
          <Icon name={icon} size="sm" />
        </button>
      )}
      <span aria-hidden="true" className="pointer-events-none absolute right-[calc(100%+6px)] top-1/2 z-20 -translate-y-1/2 whitespace-nowrap rounded-lg bg-[var(--ink)] px-2 py-1 text-[10px] font-semibold text-white opacity-0 shadow-[var(--sh-md)] transition-opacity duration-150 group-hover/action:opacity-100">
        {label}
      </span>
    </span>
  );
}

// Min-Height แถวตาราง 52px — height บน <td> ทำงานเป็น "ความสูงขั้นต่ำ" ของ cell ตาม CSS table spec (เนื้อหายาวกว่าจะขยาย)
const DT_TD = 'px-4 py-2.5 h-[52px] align-middle';

export default function Farm() {
  const navigate = useNavigate();

  const [assets, setAssets] = useState([]);
  const [bundles, setBundles] = useState([]);
  const [farmSearch, setFarmSearch] = useState('');
  const [currentFarm, setCurrentFarm] = useState('ALL');
  const [assetSearch, setAssetSearch] = useState('');
  const [transfer, setTransfer] = useState(null);
  const [addLocOpen, setAddLocOpen] = useState(false);
  // ⑦a — ช่องค้นหา-first "ของชิ้นนี้อยู่ที่ไหน?" (ค้นทุกฟาร์ม ตอบเป็นการ์ดตำแหน่ง ไม่ขึ้นกับฟาร์มที่เลือกอยู่)
  const [findQ, setFindQ] = useState('');
  // ⑦b — มุมมองหน้า: 'overview' = การ์ดภาพรวมฟาร์ม (default) | 'detail' = ตารางรายชิ้น (โครงเดิม)
  const [view, setView] = useState('overview');
  const [loading, setLoading] = useState(true);

  const load = useCallback(async () => {
    try {
      const [a, b] = await Promise.all([axios.get('/api/assets'), axios.get('/api/bundles')]);
      setAssets(a.data || []);
      setBundles(b.data || []);
    } catch (e) {
      console.error('farm load', e);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { load(); }, [load]);

  // Build farm map: นับอุปกรณ์ตามฟาร์มจริง กฎเดียวกับ Dashboard —
  // อุปกรณ์ใน Bundle ที่ถูก deploy ไปฟาร์มไหน ให้นับที่ฟาร์มนั้น (ไม่นับด้วย siteName ตัวเองซึ่งเป็นชื่อชุด)
  const { farmMap, farmTypes, farmTotals, statTotals } = useMemo(() => {
    const map = {};
    const types = {};
    const totals = {};
    const stats = {};
    const bundleFarm = {};
    bundles.forEach((b) => {
      if (b.status === 'Deployed' && (b.location || '').trim()) bundleFarm[b.bundleId] = (b.location || '').trim();
    });
    assets.forEach((a) => {
      const inDeployed = !!a.bundleId && !!bundleFarm[a.bundleId];
      if (a.bundleId && !inDeployed) return;
      const f = inDeployed ? bundleFarm[a.bundleId] : (a.siteName || 'ไม่ระบุไซต์').trim();
      if (!f) return;
      totals[f] = (totals[f] || 0) + 1;
      const isOk = !!a.status && a.status.includes('ใช้งานได้');
      const isRep = !!a.status && a.status.includes('ซ่อม');
      if (isOk || isRep) {
        if (!stats[f]) stats[f] = { ok: 0, rep: 0 };
        if (isOk) stats[f].ok++;
        if (isRep) stats[f].rep++;
      }
      if (inDeployed) return;
      if (!map[f]) map[f] = [];
      map[f].push(a);
      if (a.farmType && a.farmType !== '-') types[f] = a.farmType;
    });
    bundles.forEach((b) => {
      if (b.status === 'Deployed' && b.location) {
        const f = b.location.trim();
        if (f) { if (!map[f]) map[f] = []; if (totals[f] === undefined) totals[f] = 0; }
      }
    });
    return { farmMap: map, farmTypes: types, farmTotals: totals, statTotals: stats };
  }, [assets, bundles]);

  const groups = useMemo(() => {
    const g = {};
    Object.keys(farmMap).sort().forEach((f) => {
      const type = farmTypes[f] || 'อื่นๆ';
      if (!g[type]) g[type] = [];
      g[type].push(f);
    });
    const q = farmSearch.toLowerCase();
    const out = {};
    Object.keys(g).sort().forEach((type) => {
      const filtered = g[type].filter((f) => f.toLowerCase().includes(q));
      if (filtered.length) out[type] = filtered;
    });
    return out;
  }, [farmMap, farmTypes, farmSearch]);

  const totalsAll = Object.values(farmTotals).reduce((a, b) => a + b, 0);
  const scopeTotal = currentFarm === 'ALL' ? totalsAll : (farmTotals[currentFarm] || 0);
  const scopeOk = currentFarm === 'ALL' ? Object.values(statTotals).reduce((s, x) => s + (x.ok || 0), 0) : (statTotals[currentFarm]?.ok || 0);
  const scopeRep = currentFarm === 'ALL' ? Object.values(statTotals).reduce((s, x) => s + (x.rep || 0), 0) : (statTotals[currentFarm]?.rep || 0);

  function bundlesFor(farm) {
    const deployed = bundles.filter((b) => b.status === 'Deployed' && b.location);
    return farm === 'ALL' ? deployed : deployed.filter((b) => (b.location || '').trim() === farm);
  }

  const list = currentFarm === 'ALL' ? assets.filter((a) => !a.bundleId) : (farmMap[currentFarm] || []);
  const shownBundles = useMemo(() => bundlesFor(currentFarm), [bundles, currentFarm, farmMap]);

  // filter
  const kw = assetSearch.toLowerCase().trim();
  const filteredAssets = kw
    ? list.filter((a) =>
        (a.assetId || '').toLowerCase().includes(kw) ||
        (a.code || '').toLowerCase().includes(kw) ||
        (a.name || '').toLowerCase().includes(kw) ||
        (a.serialNumber || '').toLowerCase().includes(kw))
    : list;
  const filteredBundles = kw
    ? shownBundles.filter((b) =>
        (b.bundleId || '').toLowerCase().includes(kw) || (b.bundleName || '').toLowerCase().includes(kw))
    : shownBundles;

  const ok = scopeOk;
  const rep = scopeRep;

  function openBundle(bundleId) {
    navigate(`/bundle?open=${encodeURIComponent(bundleId)}`);
  }

  // ⑦b — สรุปต่อฟาร์ม (จากข้อมูลที่โหลดอยู่แล้ว ไม่เพิ่ม request):
  // อุปกรณ์ = ตัวเดี่ยว + สมาชิกชุดที่ deploy มาฟาร์มนี้ (กฎเดียวกับตาราง) · ชุด = ชุดที่ deploy · โรงเรือน = houseName/houseId ไม่ซ้ำ
  const farmOverview = useMemo(() => {
    const bundleFarm = {};
    bundles.forEach((b) => {
      if (b.status === 'Deployed' && (b.location || '').trim()) bundleFarm[b.bundleId] = b.location.trim();
    });
    const types = {};
    const assetsOf = {};
    const bundleCount = {};
    const houses = {};
    const stats = {};
    assets.forEach((a) => {
      const inDeployed = !!a.bundleId && !!bundleFarm[a.bundleId];
      const f = inDeployed ? bundleFarm[a.bundleId] : (a.siteName || '').trim();
      if (!f) return;
      assetsOf[f] = (assetsOf[f] || 0) + 1;
      if (inDeployed) bundleCount[f] = (bundleCount[f] || 0) + 1;
      if (!inDeployed && a.farmType && a.farmType !== '-') types[f] = types[f] || a.farmType;
      const h = (a.houseName || a.houseId || '').trim();
      if (h) { if (!houses[f]) houses[f] = new Set(); houses[f].add(h); }
      const isOk = !!a.status && a.status.includes('ใช้งานได้');
      const isRep = !!a.status && a.status.includes('ซ่อม');
      if (isOk || isRep) {
        if (!stats[f]) stats[f] = { ok: 0, rep: 0 };
        if (isOk) stats[f].ok++;
        if (isRep) stats[f].rep++;
      }
    });
    // ฟาร์มที่มีชุด deploy แต่ยังไม่มีสมาชิกเลย → โผล่การ์ดเปล่า ๆ ให้เห็นว่าฟาร์มนี้มีชุดอยู่
    Object.keys(bundleFarm).forEach((bid) => { if (!assetsOf[bundleFarm[bid]]) assetsOf[bundleFarm[bid]] = 0; });
    return Object.keys(assetsOf)
      .map((f) => ({
        name: f,
        type: types[f] || (f === 'ไม่ระบุไซต์' ? 'ไม่ระบุ' : 'อื่นๆ'),
        assets: assetsOf[f],
        bundles: bundleCount[f] || 0,
        houses: houses[f] ? houses[f].size : 0,
        ok: stats[f]?.ok || 0,
        rep: stats[f]?.rep || 0,
      }))
      .sort((a, b) => a.name.localeCompare(b.name, 'th'));
  }, [assets, bundles]);

  // ⑦a — ผลค้นหาแบบ global (ทุกฟาร์ม): ฟาร์ม + อุปกรณ์รายชิ้น + ชุด — พิมพ์ 2 ตัวขึ้นไปถึงค้น
  // ช่องเดียวค้นครบ: ชื่ออุปกรณ์ · Serial (SN) · รหัส · ชื่อฟาร์ม (ค้นได้ทั้งชื่อจริงและชื่อแบบ B2 Clean ที่ตัดรหัส BDL-xxx แล้ว)
  const find = useMemo(() => {
    const q = findQ.toLowerCase().trim();
    if (q.length < 2) return { farms: [], assets: [], bundles: [], active: false };
    const ff = farmOverview.filter((f) =>
      [f.name, cleanBundleTitle(f.name)].some((v) => (v || '').toLowerCase().includes(q)));
    const fa = assets.filter((a) =>
      [a.assetId, a.code, a.name, a.serialNumber].some((v) => (v || '').toLowerCase().includes(q)));
    const fb = bundles.filter((b) =>
      [b.bundleId, b.bundleName, cleanBundleTitle(b.bundleName)].some((v) => (v || '').toLowerCase().includes(q)));
    return { farms: ff, assets: fa, bundles: fb, active: true };
  }, [findQ, assets, bundles, farmOverview]);

  return (
    <div className="space-y-4">
      {/* Top Action Bar — ปุ่มเพิ่มซ้าย · สวิตช์มุมมองขวา — ระยะห่างสมดุล ไม่ลอยชิดขอบ/ติดช่องค้นหา */}
      <div className="flex flex-wrap items-center justify-between gap-3 pt-1">
        {/* เพิ่มฟาร์ม/โรงเรือน — ทุกคนที่ล็อกอินเพิ่มได้ (เดิม gate ไว้เฉพาะ admin ทำให้ user ติดขั้นตอนโอนย้าย) */}
        {/* เพิ่มฟาร์ม/โรงเรือน — ปุ่มเดียว แผงเดียวทำต่อเนื่อง (เดิม 2 ปุ่ม 2 modal + alert) */}
        <div className="flex gap-2 pl-0.5">
          <button onClick={() => setAddLocOpen(true)} className="flex items-center gap-1.5 h-9 px-4 rounded-xl bg-[var(--blue)] text-white text-[13px] font-semibold shadow-[var(--sh-sm)] transition hover:bg-[var(--blue-d)] hover:shadow-[var(--sh-md)] active:translate-y-px">
            <Icon name="add" size="sm" /> เพิ่มฟาร์ม / โรงเรือน
          </button>
        </div>
        {/* ⑦b — สวิตช์มุมมอง: การ์ดภาพรวมฟาร์ม (default) / ตารางรายชิ้น (โครงเดิม) */}
        <div className="flex items-center gap-1 p-1 rounded-xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-xs)]">
          <button
            onClick={() => setView('overview')}
            className={`flex items-center gap-1 px-3 h-8 rounded-lg text-[12px] font-semibold whitespace-nowrap transition-colors ${view === 'overview' ? 'bg-[var(--surface)] text-[var(--blue)] shadow-sm' : 'text-[var(--tsub)] hover:text-[var(--text)]'}`}
          >
            <Icon name="grid_on" size="xs" /> ภาพรวมฟาร์ม
          </button>
          <button
            onClick={() => setView('detail')}
            className={`flex items-center gap-1 px-3 h-8 rounded-lg text-[12px] font-semibold whitespace-nowrap transition-colors ${view === 'detail' ? 'bg-[var(--surface)] text-[var(--blue)] shadow-sm' : 'text-[var(--tsub)] hover:text-[var(--text)]'}`}
          >
            <Icon name="table_view" size="xs" /> ตารางรายชิ้น
          </button>
        </div>
      </div>

      {/* ⑦a — ช่องค้นหาลอยเม็ดเดียว (Floating Glass Search) — ยุบการ์ดขาวซ้อนชั้นเดิมให้เหลือบาร์เดียว:
          ไอคอนค้นหาเด่นชัด + พื้นโปร่งแสงเบลอ + เงาลอย sh-md — ค้นครบ ชื่ออุปกรณ์ · Serial · รหัส · ชื่อฟาร์ม (ทุกฟาร์ม) */}
      <div className="space-y-2.5">
        <div className="relative">
          <span className="absolute left-2 top-1/2 -translate-y-1/2 z-10 flex items-center justify-center w-9 h-9 rounded-xl bg-[var(--blue)] text-white shadow-[0_2px_10px_var(--blue-glow)] pointer-events-none">
            <Icon name="search" size="sm" weight="medium" />
          </span>
          <input
            value={findQ}
            onChange={(e) => setFindQ(e.target.value)}
            placeholder="ค้นหา ชื่ออุปกรณ์ · Serial (SN) · รหัส · ชื่อฟาร์ม — รวมทุกฟาร์มในช่องเดียว"
            className="w-full h-[52px] pl-[54px] pr-10 rounded-2xl border border-[var(--g200)] bg-[var(--surface)]/85 backdrop-blur-md shadow-[var(--sh-md)] text-[14px] font-medium text-[var(--text)] placeholder:text-[var(--tmuted)] placeholder:font-normal focus:outline-none focus:border-[var(--blue)] focus:shadow-[0_0_0_4px_var(--blue-glow),var(--sh-md)] transition"
          />
          {findQ && (
            <button onClick={() => setFindQ('')} title="ล้างคำค้น" className="absolute right-2.5 top-1/2 -translate-y-1/2 w-7 h-7 flex items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--g100)]"><Icon name="close" size="xs" /></button>
          )}
        </div>
        {find.active && (
          <div className="rounded-2xl border border-[var(--g200)] bg-[var(--surface)]/95 backdrop-blur-md shadow-[var(--sh-lg)] p-3 space-y-2.5">
            <div className="text-[12px] font-semibold text-[var(--tsub)]">
              พบ {find.farms.length + find.assets.length + find.bundles.length} รายการ (ฟาร์ม {find.farms.length} · อุปกรณ์ {find.assets.length} · ชุด {find.bundles.length})
            </div>
            {find.farms.length > 0 && (
              <div className="space-y-2">
                {find.farms.slice(0, 4).map((f) => (
                  <FindFarmCard key={f.name} f={f} onOpen={(n) => { setCurrentFarm(n); setView('detail'); }} />
                ))}
                {find.farms.length > 4 && <div className="text-[12px] text-[var(--tmuted)] pl-1">+ อีก {find.farms.length - 4} ฟาร์ม — พิมพ์คำให้เจาะจงขึ้นเพื่อลดผล</div>}
              </div>
            )}
            {find.bundles.length > 0 && (
              <div className="space-y-2">
                {find.bundles.slice(0, 4).map((b) => <FindBundleCard key={b.bundleId} b={b} onOpen={openBundle} />)}
                {find.bundles.length > 4 && <div className="text-[12px] text-[var(--tmuted)] pl-1">+ อีก {find.bundles.length - 4} ชุด — พิมพ์คำให้เจาะจงขึ้นเพื่อลดผล</div>}
              </div>
            )}
            {find.assets.length > 0 && (
              <div className="space-y-2">
                {find.assets.slice(0, 8).map((a) => (
                  <FindAssetCard
                    key={(a.serialNumber || '') + (a.assetId || '')}
                    a={a}
                    onTransfer={(x) => setTransfer({
                      serial: x.serialNumber,
                      current: { status: x.status, location: x.location, siteName: x.siteName, user: x.user },
                    })}
                  />
                ))}
                {find.assets.length > 8 && <div className="text-[12px] text-[var(--tmuted)] pl-1">+ อีก {find.assets.length - 8} ชิ้น — พิมพ์คำให้เจาะจงขึ้นเพื่อลดผล</div>}
              </div>
            )}
            {find.farms.length === 0 && find.assets.length === 0 && find.bundles.length === 0 && (
              <div className="text-center py-4 text-[13px] text-[var(--tmuted)]">
                <div className="flex justify-center mb-1"><Icon name="search_off" size="md" /></div>
                <div>ไม่พบของที่ตรงกับ "{findQ.trim()}"</div>
              </div>
            )}
          </div>
        )}
      </div>

      {view === 'overview' ? (
        <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">
          {farmOverview.map((f) => (
            <FarmOverviewCard
              key={f.name}
              name={f.name}
              type={f.type}
              assets={f.assets}
              bundles={f.bundles}
              houses={f.houses}
              ok={f.ok}
              rep={f.rep}
              onOpen={(n) => { setCurrentFarm(n); setView('detail'); }}
            />
          ))}
          {farmOverview.length === 0 && (
            <div className="col-span-full text-center py-12 text-[var(--tmuted)] text-[13px]">
              <div className="flex justify-center mb-2"><Icon name="factory" size="2xl" /></div>
              ยังไม่มีฟาร์ม — เพิ่มฟาร์มแรกได้จากปุ่มด้านบน
            </div>
          )}
        </div>
      ) : (
      <div className="grid grid-cols-1 lg:grid-cols-[212px_1fr] gap-4 items-start">
        {/* Farm sidebar — การ์ด Glass เรียว + ตัวเลขนับเป็น Chip แบบ Glassmorphic */}
        <div className="hidden lg:block rounded-2xl bg-[var(--surface)]/85 backdrop-blur-md border border-[var(--border)] shadow-[var(--sh-sm)] overflow-hidden lg:sticky lg:top-[var(--topbar-h)]">
          <div className="p-2 border-b border-[var(--g100)]">
            <input
              value={farmSearch}
              onChange={(e) => setFarmSearch(e.target.value)}
              placeholder="ค้นหาฟาร์ม..."
              className="w-full h-8 px-2.5 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[12px] text-[var(--ink)] focus:outline-none focus:border-[var(--blue)] focus:ring-2 focus:ring-[var(--blue-glow)]"
            />
          </div>
          <div className="overflow-y-auto lg:max-h-[65vh] p-1.5 space-y-0.5">
            <FarmItem icon="public" label="ทุกฟาร์ม" count={totalsAll} active={currentFarm === 'ALL'} onClick={() => setCurrentFarm('ALL')} />
            {Object.keys(groups).sort().map((type) => (
              <div key={type}>
                <div className="px-2 pt-2.5 pb-1 flex items-center gap-1 text-[10px] font-bold text-[var(--tmuted)]"><Icon name={TYPE_ICONS[type] || 'grass'} size="xs" /> {type}</div>
                {groups[type].map((f) => (
                  <FarmItem
                    key={f}
                    icon={TYPE_ICONS[type] || 'grass'}
                    label={cleanBundleTitle(f)}
                    bundle={bundlesFor(f).length > 0}
                    count={farmTotals[f] || 0}
                    active={currentFarm === f}
                    onClick={() => setCurrentFarm(f)}
                  />
                ))}
              </div>
            ))}
          </div>
        </div>

        {/* Main — Enterprise Industrial Data Table */}
        <div className="rounded-2xl bg-[var(--surface)] border border-[var(--border)] shadow-[var(--sh-sm)] overflow-hidden">
          {/* Header — ชื่อฟาร์มซ้าย · KPI Summary Cards (Glassmorphic) มุมขวาบนตาราง */}
          <div className="p-4 border-b border-[var(--border)] space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-2 min-w-0">
                <span className="flex items-center justify-center w-8 h-8 rounded-lg bg-[var(--ink)] text-white shrink-0"><Icon name="grid_on" size="xs" /></span>
                <div className="text-[14px] font-bold text-[var(--ink)] truncate">{currentFarm === 'ALL' ? 'ทุกฟาร์ม' : cleanBundleTitle(currentFarm)}</div>
              </div>
              {/* KPI Summary Cards — B2: "125 ชิ้น / 45 ใช้งานได้ / 13 ชุดติดตั้ง" (ไม่มีศัพท์ BUNDLE) */}
              <div className="flex flex-wrap items-center gap-2">
                <KpiGlass icon="inventory_2" tone="blue" label="อุปกรณ์" value={scopeTotal} unit="ชิ้น" />
                <KpiGlass icon="check_circle" tone="green" label="สภาพพร้อม" value={ok} unit="ใช้งานได้" />
                <KpiGlass icon="folder_open" tone="amber" label="ติดตั้งที่ฟาร์ม" value={shownBundles.length} unit="ชุดติดตั้ง" />
                {rep > 0 && <KpiGlass icon="build" tone="red" label="ต้องซ่อม" value={rep} unit="ชิ้น" />}
              </div>
            </div>
            {/* Mobile farm selector */}
            <div className="lg:hidden flex items-center gap-2">
              <Icon name="agriculture" size="sm" />
              <select
                value={currentFarm}
                onChange={(e) => setCurrentFarm(e.target.value)}
                className="flex-1 h-9 px-2.5 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px]"
              >
                <option value="ALL">ทุกฟาร์ม ({totalsAll})</option>
                {Object.entries(groups)
                  .sort(([a], [b]) => a.localeCompare(b, 'th'))
                  .flatMap(([, farms]) => farms.map((f) => (
                    <option key={f} value={f}>{cleanBundleTitle(f)} ({farmTotals[f] || 0})</option>
                  )))}
              </select>
            </div>
            <div className="relative max-w-sm">
              <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[var(--tmuted)] pointer-events-none"><Icon name="search" size="sm" /></span>
              <input
                value={assetSearch}
                onChange={(e) => setAssetSearch(e.target.value)}
                placeholder="ค้นหาอุปกรณ์ / ชุด..."
                className="w-full h-9 pl-9 pr-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px] text-[var(--ink)] focus:outline-none focus:border-[var(--blue)] focus:ring-2 focus:ring-[var(--blue-glow)]"
              />
            </div>
          </div>

          {/* ── Desktop: Enterprise Data Table — แถว Min-Height 52px · เส้นคั่นบาง var(--border) · Hover ฟ้าอ่อนโปร่งแสง · ตัวหนังสือ var(--ink) ── */}
          <div className="hidden md:block overflow-x-auto">
            <table className="w-full text-[12.5px]">
              <thead>
                <tr className="text-left text-[11px] font-bold text-[var(--tmuted)] bg-[var(--surface2)] border-b border-[var(--border)]">
                  <th className="px-4 py-3 whitespace-nowrap">รายการ</th>
                  <th className="px-4 py-3 whitespace-nowrap">Serial</th>
                  <th className="px-4 py-3 whitespace-nowrap">สถานะ</th>
                  <th className="px-4 py-3 whitespace-nowrap">ตำแหน่ง</th>
                  <th className="px-4 py-3 whitespace-nowrap">ผู้ใช้</th>
                  <th className="px-4 py-3 text-center whitespace-nowrap">จัดการ</th>
                </tr>
              </thead>
              <tbody>
                {loading && <tr><td colSpan={6} className="text-center py-10 text-[var(--tmuted)]">กำลังโหลด...</td></tr>}
                {!loading && filteredAssets.length === 0 && filteredBundles.length === 0 && (
                  <tr><td colSpan={6} className="text-center py-12 text-[var(--tmuted)]">
                    <div className="flex justify-center mb-2 text-[var(--tmuted)]"><Icon name="factory" size="2xl" /></div>ไม่พบอุปกรณ์ในไซต์งานนี้
                  </td></tr>
                )}
                {/* แถว "ชุดติดตั้ง" — ชื่อไทยเป็นตัวหนาหลัก รหัส BDL ย่อเป็น Badge เทารอง (B2) · สถานะ Pill สีชัด */}
                {filteredBundles.map((b) => (
                  <tr key={'bundle-' + b.bundleId} className="border-b border-[var(--border)] bg-[var(--blue-l)]/35 transition-colors hover:bg-[var(--blue-l)]/70 cursor-pointer" onClick={() => openBundle(b.bundleId)}>
                    <td className={DT_TD}>
                      <div className="flex items-center gap-2.5 min-w-[230px]">
                        <span className="flex items-center justify-center w-8 h-8 rounded-lg bg-[var(--surface)] border border-[var(--border)] text-[var(--blue)] shrink-0 shadow-[var(--sh-xs)]"><Icon name="folder_open" size="xs" /></span>
                        <div className="min-w-0">
                          <div className="text-[13px] font-bold text-[var(--ink)] truncate" title={b.bundleName}>{cleanBundleTitle(b.bundleName || b.bundleId)}</div>
                          <div className="mt-0.5 flex items-center gap-1.5 flex-wrap">
                            <CodeBadge>{b.bundleId}</CodeBadge>
                            <span className="inline-flex items-center gap-0.5 whitespace-nowrap rounded-full bg-[var(--blue-l)] border border-[var(--blue-b)] px-1.5 py-px text-[10px] font-bold text-[var(--blue)]"><Icon name="inventory_2" size="xs" /> {b.assetIds?.length || 0} ชิ้นในชุด</span>
                          </div>
                        </div>
                      </div>
                    </td>
                    <td className={DT_TD + ' text-center text-[var(--tmuted)]'}>—</td>
                    <td className={DT_TD}><StatusBadge status={b.status} tone={BUNDLE_TONES[b.status]} /></td>
                    <td className={DT_TD + ' font-medium text-[var(--ink)]'}>{(b.location || '').trim() || 'คลังกลาง'}</td>
                    <td className={DT_TD + ' text-center text-[var(--tmuted)]'}>—</td>
                    <td className={DT_TD}>
                      <div className="flex items-center justify-center gap-1" onClick={(e) => e.stopPropagation()}>
                        <ActionIcon icon="visibility" label="ดูรายละเอียดชุด" tone="blue" onClick={() => openBundle(b.bundleId)} />
                        <ActionIcon icon="local_shipping" label="ย้ายชุดไปฟาร์ม" tone="blue" onClick={() => navigate(`/bundle?deploy=${encodeURIComponent(b.bundleId)}`)} />
                      </div>
                    </td>
                  </tr>
                ))}
                {/* แถว "รายชิ้น" — Modern Pill Badges สีชัด + Clean Action Group Icons พร้อม Hover Tooltip */}
                {filteredAssets.map((a) => (
                  <tr key={a.serialNumber + a.assetId} className="border-b border-[var(--border)] transition-colors hover:bg-[var(--blue-l)]/45">
                    <td className={DT_TD}>
                      <div className="flex items-center gap-2.5 min-w-[230px]">
                        <span className="flex items-center justify-center w-8 h-8 rounded-lg bg-[var(--g100)] text-[var(--tsub)] shrink-0"><Icon name="devices" size="xs" /></span>
                        <div className="min-w-0">
                          <div className="text-[13px] font-bold text-[var(--ink)] truncate" title={a.name}>{a.name}</div>
                          <div className="mt-0.5 flex items-center gap-1.5 flex-wrap">
                            {a.code && <CodeBadge>{a.code}</CodeBadge>}
                            <span className="font-mono text-[10px] text-[var(--tmuted)]">{a.assetId}</span>
                          </div>
                        </div>
                      </div>
                    </td>
                    <td className={DT_TD}><span className="block max-w-[150px] truncate font-mono text-[12px] font-semibold text-[var(--ink)]" title={a.serialNumber}>{a.serialNumber}</span></td>
                    <td className={DT_TD}><StatusBadge status={a.status} tone={farmAssetTone(a.status)} /></td>
                    <td className={DT_TD}>
                      <div className="min-w-0 max-w-[260px]">
                        {a.bundleId && (
                          <span className="mb-0.5 inline-flex max-w-full items-center gap-1 rounded bg-[var(--blue-l)] px-1.5 py-px text-[10px] font-semibold text-[var(--blue)] border border-[var(--blue-b)]">
                            <Icon name="inventory_2" size="xs" className="flex-shrink-0" />
                            <span className="truncate">{a.bundleName || a.bundleId}</span>
                          </span>
                        )}
                        <LocationPath tone="ink" siteName={a.siteName} houseName={a.houseName} houseId={a.houseId} location={a.location} bundleId={a.bundleId} />
                      </div>
                    </td>
                    <td className={DT_TD + ' text-[var(--tsub)]'}>{a.user || '—'}</td>
                    <td className={DT_TD}>
                      <div className="flex items-center justify-center gap-1">
                        <ActionIcon icon="description" label="ดูประวัติการเคลื่อนไหว (Trace)" href={`/trace/${encodeURIComponent(a.serialNumber)}`} target="_blank" />
                        <ActionIcon icon="qr_code" label="ฉลาก QR" href={`/qr?serial=${encodeURIComponent(a.serialNumber)}`} target="_blank" />
                        <ActionIcon icon="local_shipping" label="โอนย้าย" tone="blue" onClick={() => setTransfer({
                          serial: a.serialNumber,
                          current: { status: a.status, location: a.location, siteName: a.siteName, user: a.user },
                        })} />
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Mobile card list — B2 Clean (ตำแหน่งผ่าน buildLocation — "คลังกลาง" ภาษาไทยเท่านั้น) + สถานะ Pill สีชัด */}
          <div className="md:hidden divide-y divide-[var(--border)]">
            {loading && <div className="text-center py-10 text-[var(--tmuted)] text-[13px]">กำลังโหลด...</div>}
            {!loading && filteredAssets.length === 0 && filteredBundles.length === 0 && (
              <div className="text-center py-12 text-[var(--tmuted)] text-[13px]">ไม่พบอุปกรณ์ในไซต์งานนี้</div>
            )}
            {filteredBundles.map((b) => (
              <div key={'bundle-' + b.bundleId} className="p-3.5 space-y-2.5 bg-[var(--blue-l)]/40" onClick={() => openBundle(b.bundleId)}>
                <div className="flex items-center gap-2.5">
                  <span className="flex items-center justify-center w-8 h-8 rounded-lg shrink-0 bg-[var(--surface)] text-[var(--blue)] shadow-sm">
                    <Icon name="folder_open" size="sm" />
                  </span>
                  <div className="flex-1 min-w-0">
                    <div className="text-[13px] font-bold text-[var(--ink)] leading-snug">{cleanBundleTitle(b.bundleName || b.bundleId)}</div>
                    <div className="mt-0.5"><CodeBadge>{b.bundleId}</CodeBadge></div>
                  </div>
                  <StatusBadge status={b.status} tone={BUNDLE_TONES[b.status]} />
                </div>
                <div className="pl-[42px] text-[12px] font-medium text-[var(--ink)]">{b.assetIds?.length || 0} ชิ้นในชุด · {(b.location || '').trim() || 'คลังกลาง'}</div>
                <div className="pl-[42px]">
                  <button onClick={(e) => { e.stopPropagation(); openBundle(b.bundleId); }} className="flex items-center gap-1 px-3 py-1.5 rounded-lg bg-[var(--blue)] text-white text-[12px] font-semibold">
                    <Icon name="folder_open" size="xs" /> ดูชุด
                  </button>
                </div>
              </div>
            ))}
            {filteredAssets.map((a) => {
              const mLoc = buildLocation({ siteName: a.siteName, houseName: a.houseName, houseId: a.houseId, location: a.location, bundleId: a.bundleId });
              return (
                <div key={a.serialNumber + a.assetId} className="p-3.5 space-y-2.5">
                  <div className="flex items-start gap-2.5">
                    <span className="flex items-center justify-center w-8 h-8 rounded-lg shrink-0 bg-[var(--g100)] text-[var(--tsub)]">
                      <Icon name="devices" size="sm" />
                    </span>
                    <div className="flex-1 min-w-0">
                      <div className="text-[13px] font-bold text-[var(--ink)] leading-snug">{a.name}</div>
                      <div className="mt-0.5 flex items-center gap-1.5 flex-wrap">
                        {a.code && <CodeBadge>{a.code}</CodeBadge>}
                        <span className="font-mono text-[10px] text-[var(--tmuted)]">{a.assetId}</span>
                      </div>
                    </div>
                    <StatusBadge status={a.status} tone={farmAssetTone(a.status)} />
                  </div>
                  <div className="pl-[42px] grid grid-cols-1 gap-1 text-[12px]">
                    <div className="text-[var(--tsub)]">Serial: <span className="font-mono font-semibold text-[var(--ink)]">{a.serialNumber}</span></div>
                    <div className="text-[var(--tsub)]">ที่อยู่: <span className="font-medium text-[var(--ink)]">{mLoc.full || 'ไม่ระบุตำแหน่ง'}</span></div>
                    {a.user && <div className="text-[var(--tsub)]">ผู้ใช้: <span className="text-[var(--ink)]">{a.user}</span></div>}
                  </div>
                  <div className="pl-[42px] flex items-center gap-2">
                    <a href={`/trace/${encodeURIComponent(a.serialNumber)}`} target="_blank" className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)]">
                      <Icon name="description" size="xs" /> Trace
                    </a>
                    <a href={`/qr?serial=${encodeURIComponent(a.serialNumber)}`} target="_blank" className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)]">
                      <Icon name="qr_code" size="xs" /> QR
                    </a>
                    <button onClick={() => setTransfer({
                      serial: a.serialNumber,
                      current: { status: a.status, location: a.location, siteName: a.siteName, user: a.user },
                    })} className="ml-auto flex items-center gap-1 px-3 py-1.5 rounded-lg bg-[var(--blue)] text-white text-[12px] font-semibold">
                      <Icon name="local_shipping" size="xs" /> โอนย้าย
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
      )}

      {transfer && (
        <TransferModal
          open
          onClose={() => setTransfer(null)}
          onSuccess={load}
          serial={transfer.serial}
          current={transfer.current}
        />
      )}

      {addLocOpen && <AddFarmOrHouseModal onClose={() => setAddLocOpen(false)} onDone={load} />}
    </div>
  );
}

// ⑦b — การ์ดภาพรวมฟาร์ม (มุมมอง default ของหน้า) — Modern Depth & B2 Clean UI:
//   • หัวการ์ด Dark Slate (--ink = #0F172A) + Category Badge คมชัด — Visual Separation จากพื้นหลัง (ภาษาเดียวกับ FarmAccessCard หน้าแรก)
//   • B2 Clean Titles — ซ่อนรหัสชุด "(BDL-xxx)" ออกจากชื่อหลัก (ชื่อจริงยังใช้ค้นหา/เปิดฟาร์มตามเดิม)
//   • Glassmorphic Metric Chips — ขอบบาง var(--border) + ตัวเลข hero ใหญ่อ่านง่าย
//   • Health Pill สีชัด + Glow — เขียวเมื่อปกติ / ส้มอำพันเรืองแสงเมื่อมีรายการต้องซ่อม (ช่างเห็นจากระยะไกล)
// กดการ์ด → เข้ามุมมองรายชิ้นของฟาร์มนั้น (ตารางเดิมกลายเป็นชั้นลึก)
function FarmOverviewCard({ name, type, assets, bundles, houses, ok, rep, onOpen }) {
  const icon = TYPE_ICONS[type] || 'help';
  const title = cleanBundleTitle(name); // B2 Clean — ชื่อหลักไม่มีรหัสชุด
  return (
    <button
      onClick={() => onOpen(name)}
      title={`เปิดมุมมองรายชิ้นของ ${title}`}
      className="group text-left rounded-[var(--r5)] bg-[var(--surface)] border border-[var(--border)] shadow-[var(--sh-sm)] overflow-hidden hover:shadow-[var(--sh-md)] hover:-translate-y-0.5 transition-all focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--blue-glow)]"
    >
      {/* Dark Slate Header — แถบสีเข้มคมชัด แยกชั้นการ์ดออกจากพื้นหลัง */}
      <div className="flex items-center gap-2.5 bg-[var(--ink)] px-4 py-3">
        <span className="w-9 h-9 rounded-xl bg-white/10 ring-1 ring-white/15 text-white/90 flex items-center justify-center shrink-0 backdrop-blur-sm">
          <Icon name={icon} size="sm" />
        </span>
        <div className="flex-1 min-w-0">
          <div className="text-[14px] font-bold text-white truncate" title={name}>{title}</div>
          {title !== name && <div className="text-[10px] text-white/45 truncate">{name}</div>}
        </div>
        {/* Category Badge — ประเภทฟาร์มบนพื้นเข้ม อ่านชัดทันที */}
        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-white/10 ring-1 ring-white/20 text-[10px] font-bold text-white/90 whitespace-nowrap shrink-0">
          <Icon name={icon} size="xs" weight="medium" /> {type}
        </span>
        <span className="text-white/50 group-hover:text-white group-hover:translate-x-0.5 transition-all shrink-0">
          <Icon name="chevron_right" size="sm" />
        </span>
      </div>

      <div className="p-4 pt-3 space-y-3">
        {/* Glassmorphic Metric Chips — Pill ขอบบางคมชัด + ตัวเลข hero ใหญ่ */}
        <div className="flex flex-wrap gap-1.5">
          <MetricChip icon="inventory_2" label="อุปกรณ์" value={assets} />
          <MetricChip icon="folder_open" label="ชุด" value={bundles} />
          <MetricChip icon="home_work" label="โรงเรือน" value={houses} />
        </div>
        {/* High-Contrast Health Status — Pill สีชัดเจน + Glow Effect */}
        <div className="flex items-center gap-2 flex-wrap">
          <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-[var(--emerald-l)] text-[var(--emerald-d)] ring-1 ring-[var(--emerald-b)] text-[11px] font-bold">
            <Icon name="check_circle" size="xs" /> ใช้งานได้ {ok}
          </span>
          {rep > 0 ? (
            <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-[var(--amber-l)] text-[var(--amber-d)] ring-1 ring-[var(--amber-b)] text-[11px] font-bold shadow-[0_0_14px_var(--amber-b)]">
              <Icon name="build" size="xs" /> ต้องซ่อม {rep}
            </span>
          ) : (
            <span className="text-[11px] font-semibold text-[var(--tmuted)]">สถานะปกติ</span>
          )}
        </div>
      </div>
    </button>
  );
}

// MetricChip — Glassmorphic Pill Badge: ขอบบาง 1px var(--border) + พื้นโปร่งแสงเบลอ + ตัวเลข hero (tabular-nums)
function MetricChip({ icon, label, value }) {
  return (
    <span className="inline-flex items-center gap-2 pl-2 pr-3 py-1 rounded-full bg-[var(--surface2)]/80 backdrop-blur-sm border border-[var(--border)] shadow-[var(--sh-xs)]">
      <span className="flex items-center justify-center w-6 h-6 rounded-full bg-[var(--surface)] border border-[var(--border)] text-[var(--blue)] shrink-0">
        <Icon name={icon} size="xs" />
      </span>
      <span className="text-[18px] font-extrabold text-[var(--text)] tabular-nums leading-none">{value}</span>
      <span className="text-[10px] font-semibold text-[var(--tmuted)]">{label}</span>
    </span>
  );
}

// ⑦a — การ์ดตอบรายชิ้น: ตำแหน่งเต็ม (ฟาร์ม › โรงเรือน › จุดติดตั้ง) + สถานะ + ปุ่มต่อ (ประวัติ/ฉลาก/ย้าย)
// ตำแหน่งอ่านจาก buildLocation() เดิมของระบบ — จัดกรณี "อยู่ในชุด" (ฟาร์มอ่านจาก Location) และคลังกลางให้แล้ว
function FindAssetCard({ a, onTransfer }) {
  const loc = buildLocation({ siteName: a.siteName, houseName: a.houseName, houseId: a.houseId, location: a.location, bundleId: a.bundleId });
  return (
    <div className="flex flex-wrap items-center gap-x-3 gap-y-2 px-3.5 py-3 rounded-xl border border-[var(--g200)] bg-[var(--surface2)] hover:border-[var(--blue-b)]">
      <div className="flex-1 min-w-[230px]">
        <div className="flex items-center gap-2 flex-wrap">
          <span className="font-mono text-[13px] font-bold text-[var(--text)]">{a.serialNumber || a.assetId}</span>
          <StatusBadge status={a.status} />
          {a.bundleId && (
            <span className="inline-flex max-w-full items-center gap-1 rounded bg-[var(--blue-l)] px-1.5 py-px text-[10px] font-medium text-[var(--blue)] border border-[var(--blue-b)]">
              <Icon name="inventory_2" size="xs" className="flex-shrink-0" />
              <span className="truncate">{a.bundleName || a.bundleId}</span>
            </span>
          )}
        </div>
        <div className="text-[12px] text-[var(--tsub)]">{a.name}</div>
        <div className="mt-1.5 flex items-center gap-1.5 text-[13px] font-semibold text-[var(--text)] flex-wrap">
          <Icon name="place" size="sm" className="text-[var(--emerald-d)]" />
          {loc.full ? <span>{loc.full}</span> : <span className="font-normal text-[var(--tmuted)]">ไม่ระบุตำแหน่ง</span>}
        </div>
      </div>
      <div className="flex items-center gap-1">
        <a href={`/trace/${encodeURIComponent(a.serialNumber)}`} target="_blank" title="ดูประวัติ" className="h-8 w-8 flex items-center justify-center rounded-lg text-[var(--blue)] hover:bg-[var(--blue-l)]"><Icon name="description" size="sm" /></a>
        <a href={`/qr?serial=${encodeURIComponent(a.serialNumber)}`} target="_blank" title="ฉลาก QR" className="h-8 w-8 flex items-center justify-center rounded-lg text-[var(--tsub)] hover:bg-[var(--surface2)]"><Icon name="qr_code" size="sm" /></a>
        <button onClick={() => onTransfer(a)} title="ย้าย" className="h-8 w-8 flex items-center justify-center rounded-lg text-[var(--blue)] hover:bg-[var(--blue-l)]"><Icon name="local_shipping" size="sm" /></button>
      </div>
    </div>
  );
}

// ⑦a — การ์ดตอบระดับชุด: ชุดอยู่ที่ไหน + กี่ชิ้น + ปุ่มดูชุด
function FindBundleCard({ b, onOpen }) {
  const atStock = b.status === 'In Stock';
  const loc = (b.location || '').trim();
  return (
    <div className="flex flex-wrap items-center gap-x-3 gap-y-2 px-3.5 py-3 rounded-xl border border-[var(--blue-b)] bg-[var(--blue-l)]">
      <div className="flex-1 min-w-[230px]">
        <div className="flex items-center gap-2 flex-wrap">
          <span className="inline-flex items-center gap-1 text-[13px] font-bold text-[var(--text)]"><Icon name="folder_open" size="sm" /> {cleanBundleTitle(b.bundleName || b.bundleId)}</span>
          <span className="font-mono text-[11px] text-[var(--tmuted)]">{b.bundleId}</span>
          <StatusBadge status={b.status} />
        </div>
        <div className="mt-1.5 flex items-center gap-1.5 text-[13px] font-semibold text-[var(--text)] flex-wrap">
          <Icon name="place" size="sm" className="text-[var(--emerald-d)]" />
          {atStock ? 'อยู่ในคลังกลาง' : (loc || 'ไม่ระบุตำแหน่ง')}
          <span className="text-[12px] font-normal text-[var(--tmuted)]">· {(b.assetIds || []).length} ชิ้น</span>
        </div>
      </div>
      <button onClick={() => onOpen(b.bundleId)} className="flex items-center gap-1 px-2.5 py-1 rounded-lg bg-[var(--blue)] text-white text-[11px] font-semibold whitespace-nowrap"><Icon name="folder_open" size="xs" /> ดูชุด</button>
    </div>
  );
}

// ⑦a — ผลค้นหาระดับฟาร์ม: ชื่อฟาร์ม (B2 Clean) + ประเภท + สรุปยอด — กดเพื่อเปิดมุมมองรายชิ้นของฟาร์มทันที
function FindFarmCard({ f, onOpen }) {
  const icon = TYPE_ICONS[f.type] || 'help';
  return (
    <button onClick={() => onOpen(f.name)} className="w-full flex flex-wrap items-center gap-x-3 gap-y-2 px-3.5 py-3 rounded-xl border border-[var(--g200)] bg-[var(--surface2)]/70 hover:border-[var(--blue-b)] hover:bg-[var(--surface)] hover:shadow-[var(--sh-sm)] text-left transition-all">
      <span className="flex items-center justify-center w-9 h-9 rounded-xl bg-[var(--emerald-l)] text-[var(--emerald-d)] flex-shrink-0"><Icon name={icon} size="sm" /></span>
      <div className="flex-1 min-w-[200px]">
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-[13px] font-bold text-[var(--text)] truncate">{cleanBundleTitle(f.name)}</span>
          <span className="px-1.5 py-px rounded-full bg-[var(--g100)] border border-[var(--border)] text-[10px] font-bold text-[var(--tsub)]">{f.type}</span>
        </div>
        <div className="text-[11px] text-[var(--tmuted)] mt-0.5">{f.assets} อุปกรณ์ · {f.bundles} ชุด · {f.houses} โรงเรือน</div>
      </div>
      <span className="flex items-center gap-1 px-2.5 py-1 rounded-lg bg-[var(--blue)] text-white text-[11px] font-semibold whitespace-nowrap"><Icon name="grid_on" size="xs" /> เปิดฟาร์ม</span>
    </button>
  );
}

function FarmItem({ icon, label, count, active, onClick, bundle }) {
  return (
    <button onClick={onClick} className={`w-full flex items-center gap-2 px-2 py-1.5 rounded-lg text-left transition-colors ${active ? 'bg-[var(--blue-l)] text-[var(--blue)]' : 'hover:bg-[var(--surface2)]'}`}>
      <span className={`flex items-center justify-center w-7 h-7 rounded-md shrink-0 ${active ? 'bg-white/70 text-[var(--blue)]' : 'bg-[var(--g100)] text-[var(--tsub)]'}`}>
        <Icon name={icon} size="xs" />
      </span>
      <span className={`flex items-center gap-1 flex-1 min-w-0 text-[12px] font-semibold truncate ${active ? 'text-[var(--blue)]' : 'text-[var(--ink)]'}`}>{label}{bundle && <Icon name="folder_open" size="xs" className="shrink-0" />}</span>
      {/* ตัวเลขนับรายการ — Glassmorphic Chip อ่านง่าย */}
      <span className={`flex items-center justify-center min-w-[26px] px-1.5 py-0.5 rounded-full text-[10px] font-bold tabular-nums backdrop-blur-sm shrink-0 ${active ? 'bg-white/75 text-[var(--blue)] border border-[var(--blue-b)]' : 'bg-white/60 text-[var(--tsub)] border border-[var(--border)]'}`}>{count}</span>
    </button>
  );
}



// ---- เพิ่มฟาร์ม / โรงเรือน — แผงเดียวทำต่อเนื่อง 2 ชั้น (ล็อกอินแล้วเพิ่มได้ทุกคน — backend เปิดสิทธิ์แล้ว) ----
// เดิม: 2 ปุ่มเปิด 2 modal แยก + แจ้งผลด้วย alert() → รวมเป็นแผงเดียว:
//   ① ชั้นเพิ่มฟาร์ม (ชื่อ + รหัส auto จากชื่อ + ประเภท) — ข้อมูลเพิ่ม (จังหวัด/ผู้จัดการ/หมายเหตุ) พับเก็บ
//   ② บันทึกฟาร์มสำเร็จ → แผงชวน "เพิ่มโรงเรือนต่อ" พร้อมเลือกฟาร์มใหม่ให้เลย (รหัสโรงเรือน auto SITE-H01 ไล่เลข)
//      หรือกด "มีฟาร์มอยู่แล้ว" ข้ามไปเพิ่มโรงเรือนกับฟาร์มเดิมได้ทันที
// ผลสำเร็จ/ข้อผิดพลาด (รวม 409 กันเพิ่มซ้ำจาก backend) แสดงในแผงเอง ไม่เด้ง alert — backend เดิมทุก endpoint
function AddFarmOrHouseModal({ onClose, onDone }) {
  const [step, setStep] = useState('site'); // 'site' = ชั้นเพิ่มฟาร์ม | 'house' = ชั้นเพิ่มโรงเรือน
  const [flash, setFlash] = useState('');   // แถบสำเร็จในแผง
  const [err, setErr] = useState('');       // แถบข้อผิดพลาดในแผง
  // ── ชั้นฟาร์ม ──
  const [siteName, setSiteName] = useState('');
  const [farmType, setFarmType] = useState('สัตว์ปีก');
  const [siteId, setSiteId] = useState('');
  const [idTouched, setIdTouched] = useState(false);
  const [province, setProvince] = useState('');
  const [manager, setManager] = useState('');
  const [note, setNote] = useState('');
  const [moreOpen, setMoreOpen] = useState(false);
  // ── ชั้นโรงเรือน ──
  const [sites, setSites] = useState([]);
  const [houses, setHouses] = useState([]);
  const [hSiteId, setHSiteId] = useState('');
  const [houseId, setHouseId] = useState('');
  const [houseName, setHouseName] = useState('');
  const [houseType, setHouseType] = useState('');
  const [capacity, setCapacity] = useState('');
  const [hNote, setHNote] = useState('');
  const [hMoreOpen, setHMoreOpen] = useState(false);
  const busy = useBusy();

  useEffect(() => {
    axios.get('/api/farm-sites').then(({ data }) => setSites(data || [])).catch(() => {});
  }, []);

  function onSiteName(v) {
    setSiteName(v);
    if (!idTouched) setSiteId(suggestSiteId(v)); // เดารหัสจากชื่อ เช่น "Farm Bangpa" → FARMBANGPA
  }

  // เลือกฟาร์ม (ชั้นโรงเรือน) → โหลดโรงเรือนที่มีอยู่ เพื่อเดารหัสถัดไป (SITE-H01, H02, ...)
  useEffect(() => {
    setHouses([]);
    setHouseId('');
    if (!hSiteId) return;
    axios.get(`/api/farm-houses/${encodeURIComponent(hSiteId)}`)
      .then(({ data }) => {
        const list = data || [];
        setHouses(list);
        setHouseId(suggestHouseId(hSiteId, list.map((h) => h.houseId)));
      })
      .catch(() => setHouses([]));
  }, [hSiteId]);

  // เพิ่มโรงเรือนสำเร็จ → โหลดรายการใหม่ + เดารหัสถัดไปให้เพิ่มต่อได้ทันที
  function refreshHouses(sid) {
    axios.get(`/api/farm-houses/${encodeURIComponent(sid)}`)
      .then(({ data }) => {
        const list = data || [];
        setHouses(list);
        setHouseId(suggestHouseId(sid, list.map((h) => h.houseId)));
      })
      .catch(() => {});
  }

  async function submitSite() {
    if (!siteId.trim() || !siteName.trim()) { setErr('กรุณากรอกชื่อฟาร์มและรหัสฟาร์ม'); return; }
    setErr('');
    await busy.run('กำลังเพิ่มฟาร์ม...', async () => {
      try {
        const { data } = await axios.post('/api/add-farm-site', { siteId: siteId.trim().toUpperCase(), siteName: siteName.trim(), farmType, province, manager, note });
        if (data.success) {
          const newId = siteId.trim().toUpperCase();
          onDone && onDone(); // รีเฟรชข้อมูลหน้า (ฟาร์มใหม่โชว์ใน sidebar ทันที)
          setSites((prev) => (prev.some((s) => s.siteId === newId) ? prev : [...prev, { siteId: newId, siteName: siteName.trim() }]));
          setFlash(`เพิ่มฟาร์ม "${siteName.trim()}" สำเร็จ — เพิ่มโรงเรือนต่อเลยไหม?`);
          setHSiteId(newId); // เลือกฟาร์มใหม่ให้เลย → โหลดโรงเรือน + เดารหัส H01
          setHouseName('');
          setStep('house');
        } else {
          setErr(data.error || 'เพิ่มฟาร์มไม่สำเร็จ');
        }
      } catch (e) {
        setErr(e.response?.data?.error || 'ไม่สามารถเชื่อมต่อได้');
      }
    });
  }

  async function submitHouse() {
    if (!hSiteId || !houseName.trim() || !houseId.trim()) { setErr('กรุณาเลือกฟาร์มและกรอกชื่อ/รหัสโรงเรือน'); return; }
    setErr('');
    await busy.run('กำลังเพิ่มโรงเรือน...', async () => {
      try {
        const { data } = await axios.post('/api/add-farm-house', { houseId: houseId.trim().toUpperCase(), siteId: hSiteId, houseName: houseName.trim(), houseType, capacity, note: hNote });
        if (data.success) {
          onDone && onDone();
          setFlash(`เพิ่มโรงเรือน "${houseName.trim()}" สำเร็จ — เพิ่มต่อได้เลย หรือปิดหน้าต่างนี้ได้`);
          setHouseName('');
          refreshHouses(hSiteId);
        } else {
          setErr(data.error || 'เพิ่มโรงเรือนไม่สำเร็จ');
        }
      } catch (e) {
        setErr(e.response?.data?.error || 'ไม่สามารถเชื่อมต่อได้');
      }
    });
  }

  return (
    <Modal
      onClose={onClose}
      title={step === 'site' ? 'เพิ่มฟาร์ม / โรงเรือน' : 'เพิ่มโรงเรือน'}
      submitLabel={step === 'site' ? 'บันทึกฟาร์ม' : 'บันทึกโรงเรือน'}
      onSubmit={step === 'site' ? submitSite : submitHouse}
      saving={busy.busy}
      busyLabel={busy.busyLabel}
    >
      {/* ตัวบอกขั้น 1-2 */}
      <div className="flex items-center gap-2 text-[12px]">
        <span className={`px-2.5 py-1 rounded-full font-semibold border ${step === 'site' ? 'bg-[var(--blue-l)] text-[var(--blue)] border-[var(--blue)]' : 'bg-[var(--surface2)] text-[var(--tsub)] border-[var(--g200)]'}`}>1 · ฟาร์ม</span>
        <Icon name="chevron_right" size="xs" className="text-[var(--tmuted)]" />
        <span className={`px-2.5 py-1 rounded-full font-semibold border ${step === 'house' ? 'bg-[var(--blue-l)] text-[var(--blue)] border-[var(--blue)]' : 'bg-[var(--surface2)] text-[var(--tsub)] border-[var(--g200)]'}`}>2 · โรงเรือน</span>
      </div>
      {/* แถบสำเร็จ / ข้อผิดพลาด — แสดงในแผง (แทน alert) */}
      {flash && (
        <div className="flex items-start gap-1.5 px-3 py-2 rounded-lg bg-[var(--emerald-l)] border border-[var(--emerald-b)] text-[12px] font-medium text-[var(--emerald-d)]">
          <Icon name="check_circle" size="sm" /> <span>{flash}</span>
        </div>
      )}
      {err && (
        <div className="flex items-start gap-1.5 px-3 py-2 rounded-lg bg-[var(--red-l)] text-[12px] font-medium text-[var(--red)]">
          <Icon name="warning" size="sm" /> <span>{err}</span>
        </div>
      )}

      {step === 'site' ? (
        <>
          <F2 label="ชื่อฟาร์ม *"><input value={siteName} onChange={(e) => onSiteName(e.target.value)} placeholder="เช่น ฟาร์มโคนมบางแพ" className={inp} /></F2>
          <F2 label="รหัสฟาร์ม (Site ID) * — เดาให้จากชื่อ แก้ได้">
            <input value={siteId} onChange={(e) => { setSiteId(e.target.value.toUpperCase()); setIdTouched(true); }} placeholder="ตัวอังกฤษ/ตัวเลข เช่น FARM01" className={inp} />
          </F2>
          <F2 label="ประเภทฟาร์ม"><select value={farmType} onChange={(e) => setFarmType(e.target.value)} className={inp}>{FARM_TYPES.map((t) => <option key={t}>{t}</option>)}</select></F2>
          <button type="button" onClick={() => setMoreOpen((v) => !v)} className="flex items-center gap-1 text-[12px] font-semibold text-[var(--blue)]">
            <Icon name={moreOpen ? 'expand_less' : 'expand_more'} size="sm" /> ข้อมูลฟาร์มเพิ่ม (จังหวัด · ผู้จัดการ · หมายเหตุ)
          </button>
          {moreOpen && (
            <>
              <F2 label="จังหวัด"><input value={province} onChange={(e) => setProvince(e.target.value)} className={inp} /></F2>
              <F2 label="ผู้จัดการฟาร์ม"><input value={manager} onChange={(e) => setManager(e.target.value)} className={inp} /></F2>
              <F2 label="หมายเหตุ"><input value={note} onChange={(e) => setNote(e.target.value)} className={inp} /></F2>
            </>
          )}
          {/* มีฟาร์มอยู่แล้ว → ข้ามไปชั้นเพิ่มโรงเรือนได้ทันที */}
          <button type="button" onClick={() => { setFlash(''); setErr(''); setStep('house'); }} className="text-[12px] font-semibold text-[var(--blue)] hover:underline self-start">
            มีฟาร์มอยู่แล้ว — ข้ามไปเพิ่มโรงเรือนเลย
          </button>
        </>
      ) : (
        <>
          <F2 label="ฟาร์ม *">
            <select value={hSiteId} onChange={(e) => setHSiteId(e.target.value)} className={inp}>
              <option value="">— เลือกฟาร์ม —</option>
              {sites.map((s) => <option key={s.siteId} value={s.siteId}>{s.siteName} ({s.siteId})</option>)}
            </select>
          </F2>
          <F2 label="ชื่อโรงเรือน *"><input value={houseName} onChange={(e) => setHouseName(e.target.value)} placeholder="เช่น โรงเรือน 1" className={inp} /></F2>
          <F2 label={`รหัสโรงเรือน — auto จากฟาร์ม แก้ได้${houses.length ? ` (ปัจจุบันมี ${houses.length} โรงเรือน)` : ''}`}>
            <input value={houseId} onChange={(e) => setHouseId(e.target.value.toUpperCase())} placeholder="เช่น SITE-H01" className={inp} />
          </F2>
          <button type="button" onClick={() => setHMoreOpen((v) => !v)} className="flex items-center gap-1 text-[12px] font-semibold text-[var(--blue)]">
            <Icon name={hMoreOpen ? 'expand_less' : 'expand_more'} size="sm" /> ข้อมูลโรงเรือนเพิ่ม (ประเภท · ความจุ · หมายเหตุ)
          </button>
          {hMoreOpen && (
            <>
              <F2 label="ประเภทโรงเรือน"><input value={houseType} onChange={(e) => setHouseType(e.target.value)} placeholder="เช่น เปิด ระบบปิด" className={inp} /></F2>
              <F2 label="ความจุ"><input value={capacity} onChange={(e) => setCapacity(e.target.value)} className={inp} /></F2>
              <F2 label="หมายเหตุ"><input value={hNote} onChange={(e) => setHNote(e.target.value)} className={inp} /></F2>
            </>
          )}
          <button type="button" onClick={() => setStep('site')} className="text-[12px] font-semibold text-[var(--tsub)] hover:underline self-start">
            ← ย้อนกลับไปเพิ่มฟาร์มอีก
          </button>
        </>
      )}
    </Modal>
  );
}

function Modal({ onClose, title, submitLabel, onSubmit, saving, busyLabel, children }) {
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center p-4 bg-black/40 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
      <div className="w-full max-w-md bg-[var(--surface)] rounded-2xl shadow-xl my-8" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-[var(--g100)]">
          <span className="text-[15px] font-bold text-[var(--text)]">{title}</span>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)]"><Icon name="close" size="sm" /></button>
        </div>
        <div className="p-5 space-y-4">{children}</div>
        <div className="flex justify-end gap-2 px-5 py-4 border-t border-[var(--g100)] bg-[var(--surface2)] rounded-b-2xl">
          <button onClick={onClose} className="px-4 py-2 rounded-lg border border-[var(--g300)] text-[13px] text-[var(--tsub)]">ยกเลิก</button>
          <button onClick={onSubmit} disabled={saving} className="px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold disabled:opacity-60">{saving ? 'กำลังบันทึก...' : '✓ ' + submitLabel}</button>
        </div>
      </div>
      <BusyOverlay label={busyLabel} />
    </div>
  );
}

function F2({ label, children }) {
  return <div className="space-y-1"><label className="block text-[12px] font-medium text-[var(--tsub)]">{label}</label>{children}</div>;
}

const inp = 'w-full h-9 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px] focus:outline-none focus:border-[var(--blue)]';