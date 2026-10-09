// Home — หน้าแรกแบบ Workdesk (UX simplify)
// หลักการ: "1 หน้าจอ = งานที่ผู้ใช้อยากทำ" — ค้นหาได้ทุกอย่างจากที่เดียว
// ลำดับความสำคัญ (Visual Hierarchy): (1) ค้นหา → ผลลัพธ์ทำต่อทันที
//   (2) งานที่ทำบ่อย / Operations (ช่าง+จัดซื้อกดใช้ตลอดวัน)
//   (3) Farm Health & Status + Quick Farm Access (compact 6 ฟาร์มแรก + ค้นหาในโซน + ดูฟาร์มทั้งหมด)
// ไม่แสดงกราฟ chart.js (กราฟเดิมอยู่ที่ /dashboard)

import { useEffect, useMemo, useState } from 'react';
import axios from 'axios';
import { useNavigate, useSearchParams } from 'react-router-dom';
import Icon from '../components/ui/Icon.jsx';
import StatusBadge from '../components/ui/StatusBadge.jsx';
import LocationPath from '../components/ui/LocationPath.jsx';
import AddDeviceModal from '../components/AddDeviceModal.jsx';
import TransferModal from '../components/TransferModal.jsx';
import { AssetHistoryModal } from '../components/AssetHistory.jsx';
import FarmDetailModal from '../components/FarmDetailModal.jsx';
import { buildFarmOverview } from '../utils/farmMonitor.js';

const IMAGE_URL = (code, ext) =>
  `https://cdn.jsdelivr.net/gh/Khemachat2003/stock-image@main/images/${code}.${ext || 'jpg'}?v=4`;

const QUICK_ACTIONS = [
  { to: '/asset', icon: 'list_alt', box: 'bg-[var(--purple-l)] text-[var(--purple)]', title: 'ทะเบียนรายชิ้น', desc: 'ดู/ค้นหารายชิ้นทั้งหมด — สถานะ, ที่ตั้ง และประวัติ' },
  { to: '/scan', icon: 'document_scanner', box: 'bg-[var(--blue-l)] text-[var(--blue)]', title: 'ย้าย/โอนอุปกรณ์', desc: 'สแกนบาร์โค้ด (หรือพิมพ์รหัส) แล้วเลือกปลายทาง' },
  { to: '/stock', icon: 'inventory_2', box: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]', title: 'เบิก–คืนของ', desc: 'หยิบของจากคลังไปใช้ที่งาน หรือคืนกลับคลัง' },
  { to: '/qr', icon: 'qr_code_2', box: 'bg-[var(--amber-l)] text-[var(--amber-d)]', title: 'พิมพ์ฉลาก', desc: 'พิมพ์ฉลาก QR / Barcode ติดอุปกรณ์' },
];

// Scalable Farm Cards (รองรับฟาร์มเพิ่มขึ้นในอนาคต): หน้าแรกโชว์แค่ 6 ฟาร์มแรกแบบ Compact Grid
// ที่เหลือกด "ดูฟาร์มทั้งหมด" เพื่อขยายดู — กันบล็อกการ์ดล้นหน้าจอ (Grid Overflow) เมื่อฟาร์มเยอะขึ้น
const FARM_PREVIEW_COUNT = 6;



export default function Home() {
  const navigate = useNavigate();
  const [searchParams, setSearchParams] = useSearchParams();
  const [q, setQ] = useState('');
  const [assets, setAssets] = useState([]);
  const [stock, setStock] = useState([]);
  const [bundles, setBundles] = useState([]);
  const [farms, setFarms] = useState([]);
  const [stats, setStats] = useState(null);
  const [recent, setRecent] = useState([]);
  const [loading, setLoading] = useState(true);
  const [addOpen, setAddOpen] = useState(false);
  const [transfer, setTransfer] = useState(null); // ปุ่ม "ย้าย" เปิดฟอร์มโอนย้ายทันทีบนหน้าแรก (ไม่ต้องเด้งไป /scan ก่อน)
  const [historySerial, setHistorySerial] = useState('');
  const [farmOpen, setFarmOpen] = useState(null); // ชื่อฟาร์มที่เปิด FarmDetailModal (Quick Farm Access — ดูรายละเอียดฟาร์มได้จากหน้าแรก)
  const [farmQuery, setFarmQuery] = useState(''); // In-Zone Quick Search: พิมพ์ค้นหาฟาร์มในโซน Farm Health ได้ทันที ไม่ต้องเลื่อนหา
  const [showAllFarms, setShowAllFarms] = useState(false); // "ดูฟาร์มทั้งหมด" — ขยาดเกิน FARM_PREVIEW_COUNT ฟาร์มแรก
  const [dataVersion, setDataVersion] = useState(0); // เพิ่มของเสร็จ → โหลดสถิติใหม่

  useEffect(() => {
    let alive = true;
    (async () => {
      // allSettled: API ตัวใดล้มเหลว (เช่น quota/เน็ต) → ส่วนอื่นยังแสดงได้
      const [a, s, d, h, b, f] = await Promise.allSettled([
        axios.get('/api/assets'),
        axios.get('/api/stock'),
        axios.get('/api/dashboard-full'),
        axios.get('/api/asset-history-recent'),
        axios.get('/api/bundles'),
        axios.get('/api/farms'),
      ]);
      if (!alive) return;
      if (a.status === 'fulfilled') setAssets(a.value.data || []);
      if (s.status === 'fulfilled') setStock(s.value.data || []);
      if (d.status === 'fulfilled') setStats(d.value.data || null);
      if (h.status === 'fulfilled') setRecent((h.value.data || []).slice(0, 8));
      if (b.status === 'fulfilled') setBundles(b.value.data || []);
      if (f.status === 'fulfilled') setFarms(f.value.data || []);
      setLoading(false);
    })();
    return () => { alive = false; };
  }, [dataVersion]);

  // Deep-link แชร์ลิงก์ฟาร์มได้: /?farm=ชื่อฟาร์ม → เปิดรายละเอียดฟาร์มทันที (แล้วล้างพารามิเตอร์)
  useEffect(() => {
    const farmParam = searchParams.get('farm');
    if (farmParam) {
      setFarmOpen(farmParam);
      setSearchParams({}, { replace: true });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [searchParams]);

  const results = useMemo(() => {
    const k = q.trim().toLowerCase();
    if (!k) return null;
    // ค้นได้ทั้ง: ชื่อ/รหัส/AssetID/Serial + PO / ล็อตรับเข้า (Batch) / Supplier
    // → ช่าง/จัดซื้อพิมพ์เลข PO แล้วเห็นทันทีว่าแต่ละชิ้นติดตั้งอยู่ที่ฟาร์มไหน โรงเรือนใด
    const hitAssets = assets
      .filter((a) => [a.name, a.code, a.assetId, a.serialNumber, a.poNumber, a.batchId, a.supplier].some((v) => (v || '').toLowerCase().includes(k)))
      .slice(0, 8);
    const hitStock = stock
      .filter((i) => [i.code, i.name].some((v) => (v || '').toLowerCase().includes(k)))
      .slice(0, 6);
    return { assets: hitAssets, stock: hitStock, total: hitAssets.length + hitStock.length };
  }, [q, assets, stock]);

  // ── Farm Health (Step 3) — สรุปรายฟาร์มด้วยกฎเดียวกับ Farm Monitor (/farm) ──
  // รวมข้อมูลฟาร์ม 2 แหล่ง: /api/farms (farmName/farmType — แหล่งหลัก) + /api/dashboard-full.farmSites (จังหวัด/ประเภท)
  const farmMetaList = useMemo(() => {
    const map = new Map();
    (stats?.farmSites || []).forEach((f) => {
      if (f?.siteName) map.set(f.siteName, { siteName: f.siteName, farmType: f.farmType, province: f.province });
    });
    (farms || []).forEach((f) => {
      const name = f?.farmName || f?.siteName;
      if (name && map.has(name)) map.set(name, { ...map.get(name), farmType: f.farmType || map.get(name).farmType });
      else if (name) map.set(name, { siteName: name, farmType: f?.farmType, province: f?.province });
    });
    return [...map.values()];
  }, [farms, stats]);
  const farmOverview = useMemo(
    () => buildFarmOverview({ assets, bundles, farms: farmMetaList }),
    [assets, bundles, farmMetaList],
  );
  const farmTotalBundles = useMemo(
    () => bundles.reduce((s, b) => s + (b.status === 'Deployed' ? 1 : 0), 0),
    [bundles],
  );
  const openedFarm = useMemo(
    () => farmOverview.find((f) => f.name === farmOpen) || { name: farmOpen, type: 'อื่นๆ', province: '' },
    [farmOverview, farmOpen],
  );
  // ── Scalable Farm Cards: ค้นหาในโซน (ชื่อ/ประเภท/จังหวัด) → โชว์ทุกผลลัพธ์ที่ตรง
  //    ไม่ค้น → โชว์ 6 ฟาร์มแรก (Compact Grid) หรือทั้งหมดถ้ากด "ดูฟาร์มทั้งหมด" ──
  const farmFilterActive = farmQuery.trim() !== '';
  const visibleFarms = useMemo(() => {
    const k = farmQuery.trim().toLowerCase();
    if (k) {
      return farmOverview.filter((f) => [f.name, f.type, f.province].some((v) => (v || '').toLowerCase().includes(k)));
    }
    return showAllFarms ? farmOverview : farmOverview.slice(0, FARM_PREVIEW_COUNT);
  }, [farmOverview, farmQuery, showAllFarms]);

  // ── สถิติเพิ่มเติม (derive จากข้อมูลที่มีอยู่ ไม่ต้องเรียก API เพิ่ม) ──
  const derived = useMemo(() => {
    const totalAssets = stats?.assets?.length ?? assets.length;
    const statusCount = stats?.statusCount || {};
    const statuses = Object.entries(statusCount)
      .map(([label, value]) => ({ label, value }))
      .sort((a, b) => b.value - a.value);
    const repairCount = (statusCount['ส่งซ่อม'] || 0) + (statusCount['ชำรุด/สูญหาย'] || 0);

    let bundleAssets = 0;
    const bundleSet = new Set();
    (stats?.assets || assets).forEach((a) => {
      if (a.bundleId) { bundleSet.add(a.bundleId); bundleAssets++; }
    });

    // Serial ลงทะเบียนล่าสุด — parse วันที่จากรูปแบบ SN-{part}-{ddmmyyyy}-{seq}
    const recentSerials = (stats?.assets || assets)
      .map((a) => {
        const m = /^SN-(.+)-(\d{2})(\d{2})(\d{4})-(\d+)$/.exec(a.serialNumber || '');
        if (!m) return null;
        return { ...a, ts: new Date(+m[4], +m[3] - 1, +m[2]).getTime(), seq: +m[5] };
      })
      .filter(Boolean)
      .sort((x, y) => (y.ts - x.ts) || (y.seq - x.seq))
      .slice(0, 6);

    const installedPct = totalAssets > 0
      ? Math.round(((stats?.totalFarmAssets ?? 0) / totalAssets) * 100)
      : 0;

    // สินค้าใกล้หมด (คลัง Office ≤ 2) — เรียงน้อยสุดก่อน
    const lowStock = stock
      .filter((i) => parseInt(i.office) <= 2)
      .sort((a, b) => a.office - b.office)
      .slice(0, 6);
    const outCount = lowStock.filter((i) => parseInt(i.office) === 0).length;

    return {
      totalAssets, statuses, repairCount,
      readyCount: Math.max(0, totalAssets - repairCount),
      bundleCount: bundleSet.size, bundleAssets,
      installedPct, recentSerials, lowStock, outCount,
    };
  }, [stats, assets, stock]);

  const todayStr = new Date().toLocaleDateString('th-TH', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

  return (
    <div className="space-y-5 max-w-5xl">
      {/* ══ Hero — ทักทาย + ค้นหา + KPI แบบ glass ══ */}
      <div className="relative overflow-hidden rounded-2xl bg-[var(--ink)] text-white shadow-[var(--sh-md)]">
        <div className="absolute -top-24 -right-14 w-72 h-72 rounded-full bg-[var(--blue)]/25 blur-3xl pointer-events-none" />
        <div className="absolute -bottom-28 left-24 w-64 h-64 rounded-full bg-[var(--emerald)]/15 blur-3xl pointer-events-none" />
        <div className="relative p-5 sm:p-6">
          {/* แถวทักทาย + รีเฟรช */}
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="min-w-0">
              <div className="text-[11px] font-semibold uppercase tracking-widest text-white/45">{todayStr}</div>
            </div>
            <button
              onClick={() => setDataVersion((v) => v + 1)}
              title="รีเฟรชข้อมูล"
              className="h-8 px-2.5 rounded-lg bg-white/10 border border-white/15 text-[12px] text-white/80 hover:bg-white/20 hover:text-white flex items-center gap-1.5 transition-colors shrink-0"
            >
              <Icon name="refresh" size="sm" /> รีเฟรช
            </button>
          </div>

          {/* Search + action */}
          <div className="mt-4 flex flex-col sm:flex-row gap-2">
            <div className="relative flex-1">
              <span className="absolute left-3.5 top-1/2 -translate-y-1/2 text-[var(--tmuted)] pointer-events-none">
                <Icon name="search" size="sm" />
              </span>
              <input
                value={q}
                onChange={(e) => setQ(e.target.value)}
                onKeyDown={(e) => { if (e.key === 'Escape') setQ(''); }}
                placeholder="พิมพ์ชื่อ / รหัส / Serial / PO / ล็อตรับเข้า..."
                className="w-full h-12 pl-10 pr-3 rounded-xl border border-transparent bg-[var(--surface)] text-[14px] text-[var(--text)] focus:outline-none focus:ring-2 focus:ring-white/50 shadow-[var(--sh-sm)] placeholder:text-[var(--tmuted)]"
              />
            </div>
            <div className="flex gap-2">
              <button
                onClick={() => navigate('/scan')}
                title="สแกนบาร์โค้ด / QR"
                className="h-12 px-4 rounded-xl bg-white/10 border border-white/20 backdrop-blur text-white text-[13px] font-semibold hover:bg-white/20 flex items-center gap-1.5 transition-colors"
              >
                <Icon name="qr_code_scanner" size="sm" /> <span className="hidden min-[430px]:inline">สแกน</span>
              </button>
              <button
                onClick={() => setAddOpen(true)}
                title="เพิ่มอุปกรณ์ใหม่ — ระบบสร้าง Serial ให้อัตโนมัติ ไม่ซ้ำ"
                className="h-12 px-4 rounded-xl bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)] active:scale-[.98] transition flex items-center gap-1.5 whitespace-nowrap"
              >
                <Icon name="add" size="sm" /> เพิ่มอุปกรณ์ใหม่
              </button>
            </div>
          </div>

          {/* KPI glass tiles */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-5">
            {loading && !stats ? (
              Array.from({ length: 4 }).map((_, i) => (
                <div key={i} className="rounded-xl bg-white/[0.06] border border-white/10 p-3.5 animate-pulse">
                  <div className="h-8 w-8 rounded-lg bg-white/10" />
                  <div className="h-5 w-14 mt-3 rounded bg-white/10" />
                  <div className="h-2.5 w-20 mt-2 rounded bg-white/10" />
                </div>
              ))
            ) : (
              <>
                <HeroKpi icon="qr_code_2" glow="rgba(124,58,237,.30)" label="Serial ทั้งหมด" value={derived.totalAssets} unit="ชิ้น" />
                <HeroKpi icon="check_circle" glow="rgba(0,200,150,.28)" label="พร้อมใช้งาน" value={derived.readyCount} unit="ชิ้น" />
                <HeroKpi icon="build" glow="rgba(224,49,49,.28)" label="ส่งซ่อม/ชำรุด" value={derived.repairCount} unit="ชิ้น" />
                <HeroKpi icon="grid_on" glow="rgba(27,108,168,.35)" label="อยู่ในชุด (Bundle)" value={derived.bundleAssets} unit="ชิ้น" />
              </>
            )}
          </div>
        </div>
      </div>

      

      {/* ══ ผลลัพธ์ค้นหา — ทำต่อได้จากที่นี่เลย ไม่ต้องเปลี่ยนหน้า ══ */}
      {results && (
        <div className="space-y-3">
          {results.total === 0 ? (
            <div className="rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] p-6 text-center text-[13px] text-[var(--tmuted)]">
              ไม่พบ "{q}" — ลองพิมพ์ชื่อหรือรหัสอื่น หรือกดปุ่ม สแกน ด้านบน
            </div>
          ) : (
            <>
              {results.assets.length > 0 && (
                <div className="space-y-2">
                  <div className="text-[11px] font-semibold uppercase tracking-wider text-[var(--tmuted)]">
                    อุปกรณ์รายชิ้น ({results.assets.length})
                  </div>
                  {results.assets.map((a) => (
                    <div
                      key={`${a.serialNumber}-${a.assetId}`}
                      className="rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] p-3.5 flex flex-wrap items-center gap-3 hover:border-[var(--blue-b)] transition-colors"
                    >
                      <div className="flex-1 min-w-[220px]">
                        <div className="flex items-center gap-2 flex-wrap">
                          <span className="text-[13px] font-semibold text-[var(--text)] truncate">{a.name || a.code}</span>
                          <StatusBadge status={a.status} />
                          {(a.poNumber || a.batchId) && (
                            <span
                              className="inline-flex items-center gap-1 rounded bg-[var(--amber-l)] px-1.5 py-px text-[10px] font-semibold text-[var(--amber-d)] border border-[var(--amber-b)]"
                              title={a.poNumber ? `PO ${a.poNumber}` : `ล็อตรับเข้า ${a.batchId}`}
                            >
                              <Icon name="local_shipping" size="xs" className="shrink-0" />
                              {a.poNumber ? `PO ${a.poNumber}` : a.batchId}
                            </span>
                          )}
                        </div>
                        <div className="text-[11px] font-mono text-[var(--blue)] mt-0.5 truncate">
                          {a.serialNumber || a.assetId || a.code}
                        </div>
                        {/* Step 3 — ตำแหน่งฟาร์มเป็น primary pill: เห็นทันทีว่าติดอยู่ฟาร์มไหน โรงเรือนใด */}
                        <LocationPath
                          className="mt-1"
                          variant="primary"
                          siteName={a.siteName}
                          houseName={a.houseName}
                          houseId={a.houseId}
                          location={a.location}
                          bundleId={a.bundleId}
                        />
                      </div>
                      <div className="flex items-center gap-1.5 flex-wrap">
                        <button
                          onClick={() => setTransfer({ serial: a.serialNumber || a.assetId || '', current: { status: a.status, location: a.location, siteName: a.siteName, user: a.user } })}
                          title="ย้าย/โอนอุปกรณ์ — เปิดฟอร์มทันที"
                          className="h-8 px-2.5 rounded-lg bg-[var(--blue)] text-white text-[12px] font-semibold hover:bg-[var(--blue-d)] flex items-center gap-1"
                        >
                          <Icon name="local_shipping" size="sm" /> ย้าย
                        </button>
                        {/* B3 (รอบ 3) — เส้นทางย้ายทั้งชุดจากหน้าแรก: ชิ้นนี้อยู่ในชุด → เปิดหน้าชุดพร้อมฟอร์มย้ายทั้งชุด (ถ้ายังอยู่คลัง) */}
                        {a.bundleId && (
                          <button
                            onClick={() => navigate(`/bundle?deploy=${encodeURIComponent(a.bundleId)}`)}
                            title="ย้ายทั้งชุด — ถ้าชุดยังอยู่คลังจะเปิดฟอร์มย้ายทั้งชุดให้เลย (ถ้าติดตั้งแล้วจะเปิดหน้าชุดเพื่อย้ายรายชิ้น/คืนเข้าคลัง)"
                            className="h-8 px-2.5 rounded-lg bg-[var(--blue-l)] border border-[var(--blue-b)] text-[var(--blue)] text-[12px] font-semibold hover:bg-[var(--blue-b)] flex items-center gap-1"
                          >
                            <Icon name="inventory_2" size="sm" /> ย้ายทั้งชุด
                          </button>
                        )}
                        <button
                          onClick={() => a.serialNumber && setHistorySerial(a.serialNumber)}
                          disabled={!a.serialNumber}
                          title="ดูประวัติการเคลื่อนไหว"
                          className="h-8 px-2.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)] hover:bg-[var(--surface2)] flex items-center gap-1 disabled:opacity-40"
                        >
                          <Icon name="history" size="sm" /> ประวัติ
                        </button>
                        <a
                          href={`/qr?serial=${encodeURIComponent(a.serialNumber || '')}`}
                          target="_blank"
                          rel="noreferrer"
                          title="พิมพ์ฉลาก"
                          className="h-8 px-2.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)] hover:bg-[var(--surface2)] flex items-center gap-1"
                        >
                          <Icon name="qr_code" size="sm" /> ฉลาก
                        </a>
                      </div>
                    </div>
                  ))}
                </div>
              )}
              {results.stock.length > 0 && (
                <div className="space-y-2">
                  <div className="text-[11px] font-semibold uppercase tracking-wider text-[var(--tmuted)]">
                    ของในคลัง ({results.stock.length})
                  </div>
                  {results.stock.map((i) => (
                    <div key={i.code} className="rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] p-3.5 flex flex-wrap items-center gap-3 hover:border-[var(--emerald-b)] transition-colors">
                      <img
                        src={IMAGE_URL(i.code, i.ext)}
                        width="44"
                        height="44"
                        loading="lazy"
                        onError={(e) => { e.currentTarget.onerror = null; e.currentTarget.src = '/image/noimage.jpg'; }}
                        className="object-cover rounded-lg border border-[var(--g200)] shrink-0"
                        alt={i.name}
                      />
                      <div className="flex-1 min-w-[180px]">
                        <div className="text-[13px] font-semibold text-[var(--text)] truncate">{i.name}</div>
                        <div className="text-[11px] font-mono text-[var(--tmuted)]">{i.code}</div>
                      </div>
                      <div className="flex items-center gap-3 text-[12px] text-[var(--tsub)]">
                        <span>คลัง <strong className="tabular-nums text-[var(--text)]">{i.office}</strong></span>
                        <span>Site <strong className="tabular-nums text-[var(--text)]">{i.site}</strong></span>
                      </div>
                      <button
                        onClick={() => navigate(`/stock?q=${encodeURIComponent(i.code)}`)}
                        disabled={!(parseInt(i.office) > 0)}
                        title="ไปหน้าเบิก–คืนของ (ค้นหาค้างไว้ให้แล้ว)"
                        className="h-8 px-3 rounded-lg bg-[var(--emerald)] text-white text-[12px] font-semibold hover:bg-[var(--emerald-d)] disabled:opacity-40 flex items-center gap-1"
                      >
                        <Icon name="add_shopping_cart" size="sm" /> เบิก
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </>
          )}
        </div>
      )}

      {/* ══ งานที่ทำบ่อย — Operations (Priority 1: ช่าง/จัดซื้อใช้กดงานจริงตลอดวัน — ขึ้นก่อนโซนภาพรวม) ══ */}
      <div>
        <div className="text-[11px] font-semibold uppercase tracking-wider text-[var(--tmuted)] mb-2">งานที่ทำบ่อย</div>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
          {QUICK_ACTIONS.map((a) => (
            <button
              key={a.title}
              onClick={() => navigate(a.to)}
              className="group text-left rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] p-4 hover:border-[var(--blue)] hover:shadow-[var(--sh-md)] hover:-translate-y-0.5 transition-all"
            >
              <span className={`flex items-center justify-center w-11 h-11 rounded-xl ${a.box} transition-transform group-hover:scale-105`}>
                <Icon name={a.icon} size="md" />
              </span>
              <div className="text-[14px] font-bold text-[var(--text)] mt-3">{a.title}</div>
              <div className="text-[12px] text-[var(--tmuted)] mt-1 leading-snug">{a.desc}</div>
              <div className="text-[12px] font-semibold text-[var(--blue)] mt-2.5 flex items-center gap-1">
                เปิด <Icon name="chevron_right" size="xs" className="transition-transform group-hover:translate-x-0.5" />
              </div>
            </button>
          ))}
        </div>
      </div>

      {/* ══ Farm Health & Status — Executive Overview (Priority 2: เห็นสถานะทุกฟาร์มจากหน้าแรก) ══ */}
      <div>
        <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
          <div className="flex items-center gap-2 text-[11px] font-semibold uppercase tracking-wider text-[var(--tmuted)]">
            <Icon name="monitoring" size="xs" /> Farm Health &amp; Status
            {farmOverview.length > 0 && (
              <span className="px-1.5 py-0.5 rounded-full bg-[var(--g200)] text-[10px] font-bold text-[var(--tsub)] normal-case tracking-normal tabular-nums">
                {farmOverview.length} ฟาร์ม
              </span>
            )}
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            {/* In-Zone Quick Search — พิมพ์ชื่อฟาร์ม/จังหวัด/ประเภท → เจอการ์ดทันที ไม่ต้องเลื่อนหา */}
            <div className="relative">
              <span className="absolute left-2.5 top-1/2 -translate-y-1/2 text-[var(--tmuted)] pointer-events-none">
                <Icon name="search" size="xs" />
              </span>
              <input
                value={farmQuery}
                onChange={(e) => setFarmQuery(e.target.value)}
                placeholder="ค้นหาฟาร์ม..."
                className="h-8 w-[172px] pl-8 pr-7 rounded-lg border border-[var(--g200)] bg-[var(--surface)] text-[12px] text-[var(--text)] placeholder:text-[var(--tmuted)] focus:outline-none focus:ring-2 focus:ring-[var(--blue-glow)] focus:border-[var(--blue-b)] transition-shadow"
              />
              {farmQuery && (
                <button
                  onClick={() => setFarmQuery('')}
                  title="ล้างการค้นหาฟาร์ม"
                  className="absolute right-1.5 top-1/2 -translate-y-1/2 text-[var(--tmuted)] hover:text-[var(--text)] transition-colors"
                >
                  <Icon name="close" size="xs" />
                </button>
              )}
            </div>
            {/* View All — รองรับฟาร์มเพิ่มขึ้นในอนาคต: หน้าแรกโชว์ 6 ฟาร์มแรก ที่เหลือขยายดูที่นี่ */}
            {!farmFilterActive && farmOverview.length > FARM_PREVIEW_COUNT && (
              <button
                onClick={() => setShowAllFarms((v) => !v)}
                title={showAllFarms ? 'ย่อกลับเหลือ 6 ฟาร์มแรก' : `แสดงฟาร์มทั้งหมด ${farmOverview.length} ฟาร์ม`}
                className="text-[12px] font-semibold text-[var(--blue)] hover:underline flex items-center gap-1"
              >
                <Icon name={showAllFarms ? 'expand_less' : 'expand_more'} size="xs" />
                {showAllFarms ? 'ย่อ' : `ดูฟาร์มทั้งหมด (${farmOverview.length})`}
              </button>
            )}
            <button
              onClick={() => navigate('/farm')}
              title="เปิดหน้า Farm Monitor — ภาพรวมฟาร์มและค้นหาของว่าอยู่ฟาร์มไหน"
              className="text-[12px] font-semibold text-[var(--blue)] hover:underline flex items-center gap-1"
            >
              <Icon name="open_in_new" size="xs" /> เปิด Farm Monitor
            </button>
          </div>
        </div>

        {/* สรุปตัวเลขรวม: ฟาร์ม · ตู้ Bundle · อุปกรณ์รวม · ต้องดูแล */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-2.5 mb-3">
          <SumTile
            icon="agriculture" tone="green" label="ฟาร์มทั้งหมด"
            value={stats?.totalFarms ?? farmOverview.length}
            sub={`${farmOverview.filter((f) => f.assets > 0).length} ฟาร์มมีอุปกรณ์ติดตั้ง`}
          />
          <SumTile
            icon="grid_on" tone="purple" label="ตู้ควบคุม (Bundle)"
            value={bundles.length}
            sub={`${farmTotalBundles} ตู้ติดตั้งที่ฟาร์ม · ${Math.max(0, bundles.length - farmTotalBundles)} ตู้ในคลัง`}
          />
          <SumTile
            icon="devices" tone="blue" label="อุปกรณ์รวมทั้งระบบ"
            value={derived.totalAssets}
            sub={`${stats?.totalFarmAssets ?? 0} ติดตั้งที่ฟาร์ม · ${stats?.totalStockAssets ?? 0} รอติดตั้ง`}
          />
          <SumTile
            icon={derived.repairCount > 0 ? 'warning' : 'check_circle'}
            tone={derived.repairCount > 0 ? 'red' : 'green'}
            label="ต้องดูแล (ซ่อม/ชำรุด)"
            value={derived.repairCount}
            sub={derived.repairCount > 0 ? 'ควรเข้าตรวจเช็คหน้างาน' : 'สถานะปกติทั้งหมด'}
          />
        </div>

        {/* Quick Farm Access — Compact Grid: โชว์ 6 ฟาร์มแรก (กันล้นจอเมื่อฟาร์มเพิ่มขึ้น) + กดเปิด FarmDetailModal ทันที */}
        {loading && farmOverview.length === 0 ? (
          <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-3">
            {Array.from({ length: 3 }).map((_, i) => (
              <div key={i} className="rounded-2xl bg-[var(--surface)] border border-[var(--g200)] p-4 animate-pulse">
                <div className="flex items-center gap-2.5">
                  <div className="w-10 h-10 rounded-xl bg-[var(--g100)]" />
                  <div className="flex-1">
                    <div className="h-3.5 w-28 rounded bg-[var(--g100)]" />
                    <div className="h-2.5 w-20 rounded bg-[var(--g100)] mt-1.5" />
                  </div>
                </div>
                <div className="grid grid-cols-3 gap-2 mt-3">
                  <div className="h-11 rounded-lg bg-[var(--g100)]" />
                  <div className="h-11 rounded-lg bg-[var(--g100)]" />
                  <div className="h-11 rounded-lg bg-[var(--g100)]" />
                </div>
              </div>
            ))}
          </div>
        ) : farmOverview.length === 0 ? (
          <div className="rounded-2xl border border-dashed border-[var(--g300)] p-6 text-center text-[13px] text-[var(--tmuted)]">
            ยังไม่มีข้อมูลฟาร์ม — เพิ่มฟาร์มแรกได้ที่หน้า Farm Monitor
          </div>
        ) : visibleFarms.length === 0 ? (
          <div className="rounded-2xl border border-dashed border-[var(--g300)] p-6 text-center text-[13px] text-[var(--tmuted)]">
            ไม่พบฟาร์มที่ตรงกับ "{farmQuery}" — ลองค้นด้วยชื่อ/จังหวัดอื่น หรือเปิด Farm Monitor ดูทั้งหมด
          </div>
        ) : (
          <>
            <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-3">
              {visibleFarms.map((f) => (
                <FarmAccessCard key={f.name} farm={f} onOpen={() => setFarmOpen(f.name)} />
              ))}
            </div>
            {!farmFilterActive && !showAllFarms && farmOverview.length > FARM_PREVIEW_COUNT && (
              <div className="mt-2.5 text-center text-[11px] text-[var(--tmuted)]">
                แสดง {visibleFarms.length} จาก {farmOverview.length} ฟาร์ม — กด "ดูฟาร์มทั้งหมด" ด้านบนเพื่อดูที่เหลือ
              </div>
            )}
          </>
        )}
      </div>

      {/* ══ สถิติเชิงลึก — ลงทะเบียนล่าสุด + สถานะอุปกรณ์ ══ */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {/* Serial ลงทะเบียนล่าสุด (parse วันที่จาก Serial) */}
          <Card icon="app_registration" tone="purple" title="ลงทะเบียนล่าสุด">
            {!loading && derived.recentSerials.length > 0 ? (
              <div className="space-y-2">
                {derived.recentSerials.map((a) => (
                  <button
                    key={a.serialNumber}
                    type="button"
                    onClick={() => a.serialNumber && setHistorySerial(a.serialNumber)}
                    disabled={!a.serialNumber}
                    title="ดูประวัติการเคลื่อนไหว"
                    className="flex w-full items-center gap-2.5 rounded-xl border border-[var(--g200)] px-2.5 py-2 text-left hover:border-[var(--purple)] hover:bg-[var(--purple-l)]/30 transition-colors disabled:opacity-50"
                  >
                    <div className="flex-1 min-w-0">
                      <div className="text-[11px] font-mono text-[var(--purple)] truncate" title={a.serialNumber}>{a.serialNumber}</div>
                      <div className="text-[12px] font-semibold text-[var(--text)] truncate">{a.name || a.code}</div>
                    </div>
                    <StatusBadge status={a.status} />
                    <Icon name="chevron_right" size="xs" className="text-[var(--tmuted)] shrink-0" />
                  </button>
                ))}
                <a
                  href="/asset"
                  target="_blank"
                  rel="noreferrer"
                  className="block pt-1 text-center text-[12px] font-semibold text-[var(--blue)] hover:underline"
                >
                  ดูทะเบียนรายชิ้นทั้งหมด
                </a>
              </div>
            ) : !loading ? (
              <div className="py-4 text-center text-[12px] text-[var(--tmuted)]">ยังไม่มีข้อมูล Serial</div>
            ) : (
              <SkeletonRows rows={4} />
            )}
          </Card>

          {/* สถานะอุปกรณ์ (progress bars) */}
          <Card icon="fact_check" tone="emerald" title="สถานะอุปกรณ์">
            {!loading && derived.statuses.length > 0 ? (
              <div className="space-y-2.5">
                {derived.statuses.slice(0, 4).map((s) => {
                  const max = Math.max(1, ...derived.statuses.map((x) => x.value));
                  const pct = (s.value / max) * 100;
                  const tone = s.label.includes('ซ่อม') ? 'var(--red)' : s.label.includes('สำรอง') ? 'var(--amber)' : s.label.includes('ชำรุด') || s.label.includes('สูญ') ? 'var(--g400)' : 'var(--emerald)';
                  return (
                    <div key={s.label}>
                      <div className="flex items-center justify-between text-[12px]">
                        <span className="text-[var(--tsub)]">{s.label}</span>
                        <span className="font-bold text-[var(--text)] tabular-nums">{s.value}</span>
                      </div>
                      <div className="h-1.5 rounded-full bg-[var(--g100)] mt-1">
                        <div className="h-full rounded-full transition-all duration-700" style={{ width: `${pct}%`, background: tone }} />
                      </div>
                    </div>
                  );
                })}
                {derived.repairCount > 0 && (
                  <div className="flex items-center gap-2 mt-1 rounded-lg bg-[var(--red-l)] border border-[var(--red-b)] px-2.5 py-2 text-[11px] font-medium text-[var(--red)]">
                    <Icon name="warning" size="xs" /> มี {derived.repairCount} ชิ้นที่ส่งซ่อม/ชำรุด — ควรตรวจสอบ
                  </div>
                )}
              </div>
            ) : (
              <SkeletonRows rows={3} />
            )}
          </Card>
        </div>

      {/* ══ สต็อกใกล้หมด — เตือนล่วงหน้าก่อนของหมดจริง ══ */}
      <div className="rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
        <div className="flex items-center justify-between px-4 py-3 border-b border-[var(--g100)]">
          <div className="flex items-center gap-2 text-[13px] font-semibold text-[var(--text)]">
            <span className="flex items-center justify-center w-7 h-7 rounded-lg bg-[var(--amber-l)] text-[var(--amber-d)]">
              <Icon name="warning" size="sm" />
            </span>
            สต็อกใกล้หมด (คลัง Office ≤ 2)
            {derived.outCount > 0 && (
              <span className="px-1.5 py-0.5 rounded-full bg-[var(--red-l)] text-[var(--red)] text-[10px] font-bold">{derived.outCount} หมด</span>
            )}
          </div>
          <button onClick={() => navigate('/stock')} title="ไปหน้าเบิก–คืนของ" className="text-[12px] font-semibold text-[var(--blue)] hover:underline">
            เบิก–คืนของ
          </button>
        </div>
        {loading ? (
          <div className="px-4 py-4"><SkeletonRows rows={2} /></div>
        ) : derived.lowStock.length === 0 ? (
          <div className="px-4 py-5 flex items-center justify-center gap-2 text-[12px] text-[var(--emerald-d)]">
            <Icon name="check_circle" size="sm" /> สต็อกทุกรายการเพียงพอ — ไม่มีรายการใกล้หมด
          </div>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 p-3">
            {derived.lowStock.map((i) => {
              const qty = parseInt(i.office);
              const out = qty === 0;
              return (
                <button
                  key={i.code}
                  onClick={() => navigate(`/stock?q=${encodeURIComponent(i.code)}`)}
                  title="ไปหน้าเบิก–คืนของ (ค้นหาค้างไว้ให้แล้ว)"
                  className="flex items-center gap-3 rounded-xl border border-[var(--g200)] px-3 py-2.5 text-left hover:border-[var(--amber)] hover:bg-[var(--amber-l)]/40 transition-colors"
                >
                  <span className={`flex items-center justify-center w-10 h-8 rounded-lg text-[13px] font-bold tabular-nums shrink-0 ${out ? 'bg-[var(--red-l)] text-[var(--red)]' : 'bg-[var(--amber-l)] text-[var(--amber-d)]'}`}>
                    {qty}
                  </span>
                  <span className="flex-1 min-w-0">
                    <span className="block text-[12px] font-semibold text-[var(--text)] truncate">{i.name}</span>
                    <span className="block text-[10px] font-mono text-[var(--tmuted)]">{i.code}</span>
                  </span>
                  <span className={`px-1.5 py-0.5 rounded-full text-[10px] font-bold shrink-0 ${out ? 'bg-[var(--red-l)] text-[var(--red)]' : 'bg-[var(--amber-l)] text-[var(--amber-d)]'}`}>
                    {out ? 'หมด' : 'ใกล้หมด'}
                  </span>
                </button>
              );
            })}
          </div>
        )}
      </div>

      {/* ══ การโอนย้ายล่าสุด — timeline เรียงเหตุการณ์ ══ */}
      <div className="rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
        <div className="flex items-center justify-between px-4 py-3 border-b border-[var(--g100)]">
          <div className="flex items-center gap-2 text-[13px] font-semibold text-[var(--text)]">
            <span className="flex items-center justify-center w-7 h-7 rounded-lg bg-[var(--blue-l)] text-[var(--blue)]">
              <Icon name="local_shipping" size="sm" />
            </span>
            การโอนย้ายล่าสุด
          </div>
          <button onClick={() => navigate('/asset')} title="ไปหน้าทะเบียนรายชิ้น" className="text-[12px] font-semibold text-[var(--blue)] hover:underline">
            ดูทะเบียนรายชิ้น
          </button>
        </div>
        {recent.length === 0 ? (
          <div className="px-4 py-6 text-center text-[12px] text-[var(--tmuted)]">
            {loading ? 'กำลังโหลด...' : 'ยังไม่มีการโอนย้าย'}
          </div>
        ) : (
          <div className="divide-y divide-[var(--g100)]">
            {recent.map((r, idx) => {
              const chip = actionChip(r.action);
              return (
                <button type="button" key={`${r.serialNumber}-${idx}`} onClick={() => r.serialNumber && setHistorySerial(r.serialNumber)} className="flex w-full items-center gap-2.5 px-4 py-2.5 text-left text-[12px] hover:bg-[var(--surface2)] transition-colors">
                  <span className={`flex items-center justify-center w-7 h-7 rounded-lg shrink-0 ${chip.cls}`}>
                    <Icon name={chip.icon} size="xs" />
                  </span>
                  <span className="font-semibold text-[var(--text)] shrink-0 hidden min-[430px]:inline">{r.action}</span>
                  <span
                    className="font-mono text-[11px] text-[var(--blue)] bg-[var(--blue-l)] rounded-md px-1.5 py-0.5 truncate max-w-[180px]"
                    title={r.serialNumber}
                  >
                    {r.serialNumber}
                  </span>
                  <span className="truncate text-[var(--tsub)] hidden sm:inline" title={r.to}>→ {r.to}</span>
                  {r.user && r.user !== '-' && (
                    <span className="hidden lg:flex items-center gap-1 text-[var(--tmuted)] shrink-0" title={`โดย ${r.user}`}>
                      <Icon name="person" size="xs" /> {r.user}
                    </span>
                  )}
                  <span className="ml-auto shrink-0 text-[var(--tmuted)] tabular-nums">{r.date}</span>
                </button>
              );
            })}
          </div>
        )}
      </div>

      {/* เพิ่มอุปกรณ์ใหม่ (โฟลว์เดียว: เลือก/สร้าง Part → Serial อัตโนมัติ กันซ้ำ) */}
      {addOpen && (
        <AddDeviceModal
          open
          onClose={() => setAddOpen(false)}
          onDone={() => { setAddOpen(false); setDataVersion((v) => v + 1); }}
        />
      )}

      {/* ย้าย/โอนอุปกรณ์ — ปุ่ม "ย้าย" ในผลค้นหาเปิดฟอร์มทันทีบนหน้าแรก (ตัดคลิกเปล่าไป /scan) */}
      {transfer && (
        <TransferModal
          open
          serial={transfer.serial}
          current={transfer.current}
          onClose={() => setTransfer(null)}
          onSuccess={() => setDataVersion((v) => v + 1)}
        />
      )}
      {/* Farm Detail Modal (Step 3) — กดการ์ดฟาร์มด้านบน → เห็น Asset Breakdown + Bundle Control ทันที */}
      {farmOpen && (
        <FarmDetailModal
          key={farmOpen}
          farm={openedFarm}
          assets={assets}
          bundles={bundles}
          loading={loading}
          onClose={() => setFarmOpen(null)}
          onTransfer={setTransfer}
          onHistory={setHistorySerial}
        />
      )}
      {historySerial && <AssetHistoryModal key={historySerial} serial={historySerial} onClose={() => setHistorySerial('')} />}
    </div>
  );
}

// ── Helpers ของหน้าแรก ──
const TONE_BOX = {
  green: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]',
  emerald: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]',
  blue: 'bg-[var(--blue-l)] text-[var(--blue)]',
  amber: 'bg-[var(--amber-l)] text-[var(--amber-d)]',
  purple: 'bg-[var(--purple-l)] text-[var(--purple)]',
};

// chip สี/ไอคอนตามประเภทการโอนย้าย (Asset_History.action)
function actionChip(action = '') {
  const a = String(action);
  if (a.includes('ลงทะเบียน')) return { icon: 'add_circle', cls: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]' };
  if (a.includes('เข้า Bundle')) return { icon: 'grid_on', cls: 'bg-[var(--blue-l)] text-[var(--blue)]' };
  if (a.includes('ออกจาก Bundle')) return { icon: 'remove', cls: 'bg-[var(--amber-l)] text-[var(--amber-d)]' };
  if (a.includes('คืนคลัง')) return { icon: 'undo', cls: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]' };
  if (a.includes('ย้าย') || a.includes('โอน')) return { icon: 'sync_alt', cls: 'bg-[var(--blue-l)] text-[var(--blue)]' };
  return { icon: 'sync_alt', cls: 'bg-[var(--g200)] text-[var(--tsub)]' };
}

function Card({ icon, tone, title, children }) {
  return (
    <div className="rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] p-4 sm:p-5 flex flex-col">
      <div className="flex items-center gap-2 text-[13px] font-semibold text-[var(--text)] mb-3">
        <span className={`w-7 h-7 flex items-center justify-center rounded-lg ${TONE_BOX[tone] || TONE_BOX.blue}`}><Icon name={icon} size="sm" /></span>
        {title}
      </div>
      <div className="flex-1">{children}</div>
    </div>
  );
}

function HeroKpi({ icon, glow, label, value, unit }) {
  return (
    <div className="relative rounded-xl bg-white/[0.06] border border-white/10 p-3.5 flex items-center gap-3 backdrop-blur-sm overflow-hidden">
      <div className="absolute -top-4 -right-4 w-16 h-16 rounded-full blur-2xl pointer-events-none" style={{ background: glow }} />
      <span
        className="relative flex items-center justify-center w-10 h-10 rounded-xl bg-white/[0.08] border border-white/10 text-white shrink-0"
        style={{ boxShadow: `inset 0 0 0 1px ${glow}, inset 0 0 20px ${glow}` }}
      >
        <Icon name={icon} size="md" />
      </span>
      <div className="relative min-w-0">
        <div className="text-[15px] font-bold text-white tabular-nums leading-tight">
          {value} <span className="text-[11px] font-medium text-white/50">{unit}</span>
        </div>
        <div className="text-[11px] text-white/55 truncate">{label}</div>
      </div>
    </div>
  );
}

// ── Farm Health & Status (Step 3) — tile สรุปรวม + การ์ด Quick Farm Access ──
const SUM_TONES = {
  green: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]',
  blue: 'bg-[var(--blue-l)] text-[var(--blue)]',
  purple: 'bg-[var(--purple-l)] text-[var(--purple)]',
  red: 'bg-[var(--red-l)] text-[var(--red)]',
  amber: 'bg-[var(--amber-l)] text-[var(--amber-d)]',
};

function SumTile({ icon, tone, label, value, sub }) {
  return (
    <div className="rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] px-3.5 py-3 flex items-center gap-3 min-w-0">
      <span className={`w-10 h-10 rounded-xl flex items-center justify-center shrink-0 ${SUM_TONES[tone] || SUM_TONES.blue}`}>
        <Icon name={icon} size="md" />
      </span>
      <div className="min-w-0">
        <div className="text-[19px] font-bold text-[var(--text)] tabular-nums leading-none">{value}</div>
        <div className="text-[11px] font-semibold text-[var(--tsub)] mt-0.5 truncate">{label}</div>
        <div className="text-[10px] text-[var(--tmuted)] truncate">{sub}</div>
      </div>
    </div>
  );
}

const FARM_TYPE_ICONS = { 'สัตว์ปีก': 'egg', 'สัตว์บก': 'pets', 'สุกร': 'agriculture', 'อื่นๆ': 'category', 'ไม่ระบุ': 'help' };

// การ์ดสรุปรายฟาร์ม — กดครั้งเดียวเปิด Asset Breakdown / Bundle Control บนหน้าแรก (FarmDetailModal)
function FarmAccessCard({ farm, onOpen }) {
  const icon = FARM_TYPE_ICONS[farm.type] || 'help';
  const totalActive = farm.ok + farm.rep;
  const healthPct = totalActive > 0 ? Math.round((farm.ok / totalActive) * 100) : 100;
  return (
    <button
      onClick={onOpen}
      title={`เปิดรายละเอียดฟาร์ม ${farm.name} — Asset Breakdown + Bundle Control`}
      className="group text-left rounded-2xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] p-4 hover:border-[var(--blue-b)] hover:shadow-[var(--sh-md)] transition-all focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--blue-glow)]"
    >
      <div className="flex items-start gap-2.5">
        <span className="w-10 h-10 rounded-xl bg-[var(--emerald-l)] text-[var(--emerald-d)] flex items-center justify-center shrink-0 group-hover:scale-105 transition-transform">
          <Icon name={icon} size="md" />
        </span>
        <div className="flex-1 min-w-0">
          <div className="text-[14px] font-bold text-[var(--text)] truncate">{farm.name}</div>
          <div className="text-[11px] text-[var(--tmuted)] truncate">{farm.type}{farm.province ? ` · ${farm.province}` : ''}</div>
        </div>
        <span className="text-[var(--tmuted)] group-hover:text-[var(--blue)] group-hover:translate-x-0.5 transition-all shrink-0 mt-1">
          <Icon name="chevron_right" size="sm" />
        </span>
      </div>
      <div className="grid grid-cols-3 gap-2 mt-3">
        <MiniNum icon="inventory_2" label="อุปกรณ์" value={farm.assets} />
        <MiniNum icon="grid_on" label="ตู้ชุด" value={farm.bundles} />
        <MiniNum icon="home_work" label="โรงเรือน" value={farm.houses} />
      </div>
      <div className="mt-3">
        <div className="h-1.5 rounded-full bg-[var(--g100)] overflow-hidden">
          <div
            className={`h-full rounded-full transition-all duration-700 ${farm.rep > 0 ? 'bg-[var(--amber)]' : 'bg-[var(--emerald)]'}`}
            style={{ width: `${healthPct}%` }}
          />
        </div>
        <div className="flex items-center justify-between mt-1.5 text-[10.5px]">
          <span className="flex items-center gap-1 text-[var(--emerald-d)]">
            <span className="w-1.5 h-1.5 rounded-full bg-[var(--emerald)]" /> ใช้งานได้ {farm.ok}
          </span>
          {farm.rep > 0 ? (
            <span className="flex items-center gap-1 text-[var(--amber-d)] font-semibold">
              <span className="w-1.5 h-1.5 rounded-full bg-[var(--amber)]" /> ส่งซ่อม {farm.rep}
            </span>
          ) : (
            <span className="text-[var(--tmuted)]">สถานะปกติ</span>
          )}
        </div>
      </div>
    </button>
  );
}

function MiniNum({ icon, label, value }) {
  return (
    <div className="rounded-lg bg-[var(--surface2)] border border-[var(--g100)] px-2 py-1.5 min-w-0">
      <div className="flex items-center gap-1 text-[10px] text-[var(--tmuted)]">
        <Icon name={icon} size="xs" className="shrink-0" />
        {label}
      </div>
      <div className="text-[14px] font-bold text-[var(--text)] tabular-nums leading-tight">{value}</div>
    </div>
  );
}

function SkeletonRows({ rows = 3 }) {
  return (
    <div className="space-y-2.5 animate-pulse">
      {Array.from({ length: rows }).map((_, i) => (
        <div key={i} className="h-3 rounded bg-[var(--g100)]" style={{ width: `${90 - i * 18}%` }} />
      ))}
    </div>
  );
}
