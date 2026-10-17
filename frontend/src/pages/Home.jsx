// Home — หน้าแรกแนว B2 (Greeting & Quick Actions) — ธีม Slate/Blue เดิมของระบบ
// โครง (บนลงล่าง กึ่งกลาง ≤1100px): แถวเครื่องมือ (รีเฟรช — ไม่มีโลโก้/Avatar ซ้ำ
//   เพราะโลโก้+Nav Pills+Avatar อยู่ที่ Topbar ของ Layout แล้ว (B2): ปุ่มสามขีดเปิด drawer เมนูเสริม / bottom nav (มือถือ <md))
//   → ส่วนทักทาย + ช่องค้นหาทรงเม็ด (มือถือ ≤520px มีปุ่มสแกน QR เต็มความกว้างใต้ช่องค้นหา) → ปุ่มงานหลักวงกลม 5 ปุ่ม
//   → ผลลัพธ์ค้นหา / สถิติ / เตือนของใกล้หมด / โอนย้ายล่าสุด / โซนฟาร์ม (ฟังก์ชันเดิมครบทุกอย่าง)
// Clean Home Layout (ล่าสุด): การ์ด "ติดตามและจัดการอุปกรณ์ตาม PO" ถูกถอดออกจากหน้าแรกถาวร (ไม่ดันเนื้อหา/ไม่ต้อง scroll โดยไม่จำเป็น)
//   → กดปุ่มวงกลม "ย้ายอุปกรณ์" เปิด Transfer Center Modal ป๊อปอัปกลางหน้าจอแทน (2 แท็บ: โอนย้ายตาม PO / โอนย้ายอุปกรณ์ทั่วไป นอก PO)
// ฟังก์ชันหน้าแรกเดิมที่คงไว้: ค้นหา (ชื่อ/รหัส/Serial/PO/ล็อต/ผู้ขาย) + ย้าย (TransferModal) + ประวัติ
//   + ฉลาก + เบิก (/stock?q=) + เพิ่มอุปกรณ์ (AddDeviceModal) + รีเฟรช + เตือนของใกล้หมด
//   + Farm Health & Quick Farm Access (FarmDetailModal, deep-link ?farm=) + ลงทะเบียนล่าสุด
// ไม่แสดงกราฟ chart.js (กราฟเดิมอยู่ที่ /dashboard)

import { useEffect, useMemo, useState } from 'react';
import axios from 'axios';
import { useNavigate, useSearchParams } from 'react-router-dom';
import { isInStock } from '../utils/location.js';
import { isRepairStatus } from '../utils/status.js';
import Icon from '../components/ui/Icon.jsx';
import StatusBadge from '../components/ui/StatusBadge.jsx';
import LocationPath from '../components/ui/LocationPath.jsx';
import AddDeviceModal from '../components/AddDeviceModal.jsx';
import TransferModal from '../components/TransferModal.jsx';
import TransferCenterModal from '../components/TransferCenterModal.jsx';
import { ToastHost } from '../components/ui/Toast.jsx';
import { AssetHistoryModal } from '../components/AssetHistory.jsx';
import FarmDetailModal from '../components/FarmDetailModal.jsx';
import { buildFarmOverview } from '../utils/farmMonitor.js';

const IMAGE_URL = (code, ext) =>
  `https://cdn.jsdelivr.net/gh/Khemachat2003/stock-image@main/images/${code}.${ext || 'jpg'}?v=4`;

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
  const [transferCenterOpen, setTransferCenterOpen] = useState(false); // ปุ่มวงกลม "ย้ายอุปกรณ์" — เปิด Transfer Center Modal ป๊อปอัปกลางหน้าจอ (2 แท็บ: ตาม PO / อุปกรณ์ทั่วไป นอก PO)
  const [lowStockOpen, setLowStockOpen] = useState(false); // แถบเตือนของใกล้หมด — กดกาง/พับรายการ
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
    // ค้นได้ทั้ง: ชื่อ/รหัส/AssetID/Serial + PO / ล็อตรับเข้า (Batch) / Supplier / ฟาร์ม-โรงเรือน
    // → ช่างพิมพ์ชื่อฟาร์มแล้วเห็นอุปกรณ์ทั้งหมดที่ติดตั้งอยู่ + พิมพ์เลข PO เห็นทันทีว่าชิ้นไหนอยู่ไหน
    const hitAssets = assets
      .filter((a) => [a.name, a.code, a.assetId, a.serialNumber, a.poNumber, a.batchId, a.supplier, a.siteName, a.houseName].some((v) => (v || '').toLowerCase().includes(k)))
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
  // การ์ดสรุปสถานะ 3 ใบ "นับจากตำแหน่ง/สถานะจริงของทุกชิ้น" ไม่กรองเฉพาะคำสถานะ
  // (เดิมการ์ด "อยู่ในคลัง" นับเฉพาะสถานะ "สำรอง" → ขึ้น 0 ทั้งที่มีอุปกรณ์อยู่ในคลัง):
  //   • อยู่ในคลัง  = ทุกชิ้นที่ยังไม่ได้ผูกเข้ากับฟาร์ม/โรงเรือน — พร้อมใช้งานที่คลังกลาง
  //                   (Intranin / Stock / คลังกลาง), ไม่ระบุฟาร์มจัดเก็บ, หรืออยู่ในชุดที่ยังไม่ได้ deploy
  //                   (กฎตำแหน่งเดียวกับ LocationPath "คลังกลาง" — utils/location.js isInStock)
  //   • ติดตั้งแล้ว = ชิ้นที่ผูกเข้ากับฟาร์ม/โรงเรือนแล้ว (ตัวเดี่ยวมีฟาร์ม / สมาชิกชุดที่ Deployed)
  //   • ส่งซ่อม     = สถานะ ส่งซ่อม / รอซ่อม / ชำรุด·สูญหาย (Damaged) — utils/status.js isRepairStatus
  // แหล่งข้อมูล: /api/assets (ทะเบียนเต็มคอลัมน์ตำแหน่ง — ชุดข้อมูลเดียวกับที่ deep-link ไปเปิดที่ /asset)
  // ถ้า API ทะเบียนล้มเหลวค่อยใช้ assets ของ /api/dashboard-full แทน (กันตัวเลขติด 0)
  const derived = useMemo(() => {
    const list = assets.length ? assets : (stats?.assets || assets);
    const totalAssets = list.length;
    // (Step 3: ตัด derived.statuses ทิ้ง — การ์ด "สถานะอุปกรณ์" ถูกลบไปแล้ว ไม่มีผู้ใช้ข้อมูลชุดนี้)

    let inStockCount = 0;
    let installedCount = 0;
    let bundleAssets = 0;
    const bundleSet = new Set();
    list.forEach((a) => {
      if (isInStock(a)) inStockCount += 1; else installedCount += 1;
      if (a.bundleId) { bundleSet.add(a.bundleId); bundleAssets++; }
    });
    // ต้องดูแล (ซ่อม/ชำรุด) — นับจากสถานะจริงของทุกชิ้น (ทนทานกว่านับคำจาก statusCount ที่อาจไม่ครบ)
    const repairCount = list.filter((a) => isRepairStatus(a.status)).length;

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
      totalAssets, repairCount,
      readyCount: Math.max(0, totalAssets - repairCount),
      bundleCount: bundleSet.size, bundleAssets,
      installedPct, recentSerials, lowStock, outCount,
      inStockCount,
      installedCount,
      repairPendingCount: repairCount,
    };
  }, [stats, assets, stock]);

  // Serial → ชื่ออุปกรณ์จริง (จากทะเบียนที่โหลดแล้ว) — ใช้กับ "เพิ่งเกิดขึ้น" ให้ชื่ออุปกรณ์เป็นตัวหนาหลัก
  const nameBySerial = useMemo(() => {
    const m = new Map();
    (stats?.assets || assets).forEach((a) => {
      if (a?.serialNumber && !m.has(a.serialNumber)) {
        m.set(a.serialNumber, { name: a.name || a.code || '' });
      }
    });
    return m;
  }, [stats, assets]);

  // วันที่ภาษาไทยแบบสั้น: "วันศุกร์ที่ 9 ตุลาคม"
  const now = new Date();
  const todayStr = `${now.toLocaleDateString('th-TH', { weekday: 'long' })}ที่ ${now.toLocaleDateString('th-TH', { day: 'numeric', month: 'long' })}`;

  // ปุ่มงานหลักวงกลม 5 ปุ่ม (แนว B2) — งานจบใน 1 คลิก
  // มือถือ ≤520px: ปุ่มสแกน QR (primary) ถูกซ่อนจากแถววงกลม — ไปแสดงเป็นปุ่มเต็มความกว้างใต้ช่องค้นหาแทน
  const orbActions = [
    { key: 'scan', icon: 'qr_code_scanner', label: 'สแกน QR', title: 'สแกนบาร์โค้ด / QR', primary: true, onClick: () => navigate('/scan') },
    { key: 'transfer', icon: 'local_shipping', label: 'ย้ายอุปกรณ์', title: 'เปิด Transfer Center — โอนย้ายตาม PO หรืออุปกรณ์ทั่วไป (นอก PO)', onClick: () => setTransferCenterOpen(true) },
    { key: 'stock', icon: 'inventory_2', label: 'เบิก–คืนของ', title: 'เบิก / คืนของจากคลัง', onClick: () => navigate('/stock') },
    { key: 'add', icon: 'add', label: 'เพิ่มอุปกรณ์', title: 'เพิ่มอุปกรณ์ใหม่ — ระบบสร้าง Serial ให้อัตโนมัติ ไม่ซ้ำ', onClick: () => setAddOpen(true) },
    { key: 'bundle', icon: 'folder_open', label: 'ชุดติดตั้งฟาร์ม', title: 'จัดการชุดอุปกรณ์ติดตั้งที่ฟาร์ม', onClick: () => navigate('/bundle') },
  ];

  // ปุ่ม "ค้นหา" ในช่องค้นหา — มีคำค้น → เลื่อนไปดูผลลัพธ์ / ยังว่าง → โฟกัสช่องค้นหา
  function scrollToResults() {
    if (!q.trim()) {
      document.getElementById('home-search-input')?.focus();
      return;
    }
    document.getElementById('home-results')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // จำนวนฟาร์มที่มีอุปกรณ์ติดตั้งอยู่จริง (สำหรับ "ติดตั้งแล้ว · ใน N ฟาร์ม") — จาก Farm Overview เดิม
  const installedFarmCount = farmOverview.filter((f) => f.assets > 0).length;

  // การ์ดสรุปสถานะ 3 ใบ (แนว B2) — ใช้ทั้งสีและรูปร่าง (■ ● ▲) เสมอ
  // ตัวเลข/ชื่อสถานะเป็น slate เข้ม (contrast ≥ 4.5:1) — สีสถานะใช้กับรูปร่างเท่านั้น
  // ตัวเลขนับจากตำแหน่ง/สถานะจริงของทุกชิ้น (isInStock / isRepairStatus) — ไม่กรองเฉพาะคำสถานะ "สำรอง"
  // กดการ์ด → หน้าทะเบียนรายชิ้นพร้อมตัวกรองที่ "ตรงกับตัวเลขบนการ์ดเป๊ะ":
  //   • อยู่ในคลัง  → /asset?location=stock  (เฉพาะชิ้นในคลังกลาง — ไม่ใช่ ?status=สำรอง ที่หาไม่เจอ)
  //   • ติดตั้งแล้ว → /asset?location=farm   (เฉพาะชิ้นที่ผูกเข้ากับฟาร์ม/โรงเรือน)
  //   • ส่งซ่อม     → /asset?status=repair   (ส่งซ่อม / รอซ่อม / ชำรุด·สูญหาย)
  const statusCards = [
    {
      key: 'stock', shape: '■', shapeCls: 'text-[var(--blue)]', label: 'อยู่ในคลัง',
      sub: 'พร้อมใช้งานในคลังกลาง', value: derived.inStockCount,
      title: 'ดูอุปกรณ์ที่ยังอยู่ในคลังกลาง (ยังไม่ได้ติดตั้งที่ฟาร์ม)',
      onClick: () => navigate('/asset?location=stock'),
    },
    {
      key: 'installed', shape: '●', shapeCls: 'text-[var(--emerald-d)]', label: 'ติดตั้งแล้ว',
      sub: `ใน ${installedFarmCount} ฟาร์ม`, value: derived.installedCount,
      title: 'ดูอุปกรณ์ที่ผูกเข้ากับฟาร์ม/โรงเรือนแล้ว',
      onClick: () => navigate('/asset?location=farm'),
    },
    {
      key: 'repair', shape: '▲', shapeCls: 'text-[var(--amber-d)]', label: 'ส่งซ่อม',
      sub: 'รอตามผล', value: derived.repairPendingCount,
      title: 'ดูอุปกรณ์ที่ส่งซ่อม / ชำรุด',
      onClick: () => navigate('/asset?status=repair'),
    },
  ];

  return (
    <div className="w-full max-w-[1100px] mx-auto space-y-6">
      {/* ══ 1. แถวเครื่องมือ (ขวา) — รีเฟรช ══
          Avatar ผู้ใช้แสดงจุดเดียวที่ Topbar ของ Layout (B2 — ไม่ซ้ำหน้าละ)
          เมนู: Nav Pills กลาง Topbar (≥lg) / ปุ่มสามขีด (drawer เมนูเสริม) / bottom nav (มือถือ <md) */}
      <header className="w-full flex items-center justify-end gap-1.5">
        <button
          onClick={() => setDataVersion((v) => v + 1)}
          title="รีเฟรชข้อมูล"
          className="w-10 h-10 rounded-full border border-[var(--g200)] bg-[var(--surface)] text-[var(--tsub)] hover:text-[var(--blue)] hover:border-[var(--blue-b)] flex items-center justify-center transition-colors focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--blue-glow)]"
        >
          <Icon name="refresh" size="sm" />
        </button>
      </header>

      {/* ══ 2. ส่วนทักทาย + ช่องค้นหา (กึ่งกลาง) ══ */}
      <section className="text-center w-full max-w-[680px] mx-auto">
        <div className="text-[14px] text-[var(--tmuted)]">{todayStr}</div>
       
       
        {/* ช่องค้นหาทรงเม็ดยาว สูง 64px — ค้นหาชื่อ/รหัส/Serial/PO/ล็อต (logic เดิม พิมพ์แล้วเจอทันที) */}
        <div className="mt-5 flex items-center gap-2 h-16 min-h-16 rounded-full bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-md)] px-2 pl-6">
          <Icon name="search" size="md" className="text-[var(--tmuted)] shrink-0" />
          <input
            id="home-search-input"
            value={q}
            onChange={(e) => setQ(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Escape') setQ('');
              if (e.key === 'Enter') scrollToResults();
            }}
            placeholder="พิมพ์ชื่อ / รหัส / Serial / PO / ล็อตรับเข้า..."
            aria-label="ค้นหาอุปกรณ์"
            className="flex-1 min-w-0 h-full bg-transparent border-0 outline-none text-[15px] text-[var(--text)] placeholder:text-[var(--tmuted)]"
          />
          <button
            onClick={scrollToResults}
            title="ค้นหา — เลื่อนไปดูผลลัพธ์ด้านล่าง"
            className="h-12 px-6 rounded-full bg-[var(--blue)] text-white text-[15px] font-semibold hover:bg-[var(--blue-d)] active:scale-[.98] transition flex items-center gap-1.5 shrink-0 focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--blue-glow)]"
          >
            ค้นหา
          </button>
        </div>
        <div className="mt-2 text-[12px] text-[var(--tmuted)]">ค้นหาได้ทั้งชื่ออุปกรณ์ รหัส SN PO / ล็อตรับเข้า และชื่อฟาร์ม</div>

        {/* มือถือ ≤520px: ปุ่มสแกน QR เด่นเต็มความกว้างใต้ช่องค้นหา (งานหลักที่ใช้บ่อยที่สุดไปถึงง่าย)
            ในแถววงกลมด้านล่างถูกซ่อนบนมือถือ — ไม่ซ้ำสองที่ */}
        <button
          onClick={() => navigate('/scan')}
          title="สแกนบาร์โค้ด / QR"
          className="min-[520px]:hidden mt-4 w-full h-14 rounded-full bg-[var(--blue)] text-white text-[15px] font-semibold hover:bg-[var(--blue-d)] active:scale-[.98] transition flex items-center justify-center gap-2 shadow-[0_10px_24px_var(--blue-glow)] focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--blue-glow)]"
        >
          <Icon name="qr_code_scanner" size="md" /> สแกน QR
        </button>
      </section>

      {/* ══ 3. ปุ่มงานหลักวงกลม 5 ปุ่ม (จอแคบขึ้นบรรทัดใหม่อัตโนมัติ) ══ */}
      <section aria-label="งานหลัก">
        <div className="flex flex-wrap justify-center gap-x-7 gap-y-4">
          {orbActions.map((o) => (
            <button
              key={o.key}
              onClick={o.onClick}
              title={o.title}
              // มือถือ ≤520px: ปุ่มสแกน QR ซ่อนจากแถววงกลม (แสดงเป็นปุ่มเต็มความกว้างใต้ช่องค้นหาแทน)
              className={`group flex-col items-center gap-2 w-[92px] rounded-2xl focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--blue-glow)] ${o.primary ? 'hidden min-[520px]:flex' : 'flex'}`}
            >
              <span
                className={`w-[92px] h-[92px] rounded-full flex items-center justify-center transition-all group-hover:-translate-y-0.5 group-active:scale-95 ${
                  o.primary
                    ? 'bg-[var(--blue)] text-white shadow-[0_10px_24px_var(--blue-glow)]'
                    : 'bg-[var(--surface)] border border-[var(--g200)] text-[var(--blue)] shadow-[var(--sh-sm)] group-hover:border-[var(--blue-b)]'
                }`}
              >
                <Icon name={o.icon} size="2xl" weight="light" />
              </span>
              <span className={`text-[13px] ${o.primary ? 'font-semibold text-[var(--text)]' : 'font-medium text-[var(--tsub)]'}`}>{o.label}</span>
            </button>
          ))}
        </div>
      </section>

      {/* Clean Home Layout: การ์ด PO Tracking แผ่กว้างถูกถอดออกถาวร — การโอนย้ายทั้งหมดอยู่ใน Transfer Center Modal
          (เปิดด้วยปุ่มวงกลม "ย้ายอุปกรณ์" ด้านบน หรือปุ่มในการ์ด "โอนย้ายล่าสุด" ตอนยังไม่มีรายการ) */}

      

      {/* ══ ผลลัพธ์ค้นหา — ทำต่อได้จากที่นี่เลย ไม่ต้องเปลี่ยนหน้า ══ */}
      {results && (
        <div className="space-y-3" id="home-results">
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

      {/* ══ 4. แถบเตือนของใกล้หมด — แถบเล็กเหนือการ์ดสถานะ (เดิม: การ์ด "สต็อกใกล้หมด") กดกางดูรายการได้ ══ */}
      {derived.lowStock.length > 0 && (
        <div className="rounded-[var(--r5)] bg-[var(--amber-l)] border border-[var(--amber-b)] border-l-4 border-l-[var(--amber)] shadow-[var(--sh-sm)] overflow-hidden">
          <button
            type="button"
            onClick={() => setLowStockOpen((v) => !v)}
            title="กดเพื่อกาง/พับรายการของใกล้หมด"
            className="w-full flex items-center justify-between gap-2 px-4 py-3 min-h-[48px] text-left focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--amber-d)]"
          >
            <span className="flex items-center gap-2.5 min-w-0 text-[13px]">
              <span
                aria-hidden="true"
                className="flex items-center justify-center w-8 h-8 rounded-full bg-[var(--amber)] text-white shrink-0 shadow-[0_2px_10px_rgba(255,149,0,.35)]"
              >
                <Icon name="warning" size="sm" fill="half" />
              </span>
              <b className="text-[var(--amber-d)]">ของใกล้หมด {derived.lowStock.length} รายการ</b>
              {derived.outCount > 0 && (
                <span className="px-1.5 py-0.5 rounded-full bg-[var(--red-l)] text-[var(--red)] text-[10px] font-bold shrink-0">หมด {derived.outCount}</span>
              )}
              <span className="text-[var(--tsub)] truncate hidden sm:inline">— ควรเติมคลังเร็ว ๆ นี้</span>
            </span>
            <span className="flex items-center gap-0.5 text-[12px] font-semibold text-[var(--amber-d)] shrink-0">
              {lowStockOpen ? 'พับ' : 'ดูรายการ'}
              <Icon name={lowStockOpen ? 'expand_less' : 'expand_more'} size="sm" />
            </span>
          </button>
          {lowStockOpen && (
            <div className="border-t border-[var(--amber-b)] bg-[var(--surface)] p-3">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
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
              <button
                onClick={() => navigate('/stock')}
                title="ไปหน้าเบิก–คืนของ"
                className="mt-2 w-full min-h-[40px] rounded-xl border border-[var(--g200)] text-[12px] font-semibold text-[var(--blue)] hover:bg-[var(--blue-l)] focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--blue-glow)]"
              >
                เปิดหน้าเบิก–คืนของ
              </button>
            </div>
          )}
        </div>
      )}

      {/* ══ 5. การ์ดสรุปสถานะ 3 ใบ — ตัวเลขจริงจากข้อมูล กด → หน้ารายการกรองตามสถานะ ══
          Step 2: จัดชั้นข้อมูลใหม่ — label บน · ตัวเลขเด่นชิดซ้าย (sub อยู่ baseline เดียวกัน)
          → ไม่มี dead space จากตัวเลขชิดขวา + ขอบบาง var(--border) + soft drop shadow */}
      <section aria-label="สรุปสถานะอุปกรณ์">
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          {statusCards.map((c) => (
            <button
              key={c.key}
              onClick={c.onClick}
              title={c.title}
              className="group text-left rounded-[var(--r5)] bg-[var(--surface)] border border-[var(--border)] shadow-[var(--sh-sm)] px-4 py-3.5 flex flex-col gap-2 hover:shadow-[var(--sh-md)] hover:-translate-y-0.5 transition-all focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--blue-glow)]"
            >
              <div className="flex items-center gap-1.5 text-[13px] font-semibold text-[var(--tsub)]">
                <span aria-hidden="true" className={`text-[15px] leading-none ${c.shapeCls}`}>{c.shape}</span>
                {c.label}
              </div>
              <div className="flex items-baseline gap-2 min-w-0">
                <span className="text-[32px] leading-none font-extrabold tabular-nums text-[var(--text)]">
                  {loading && !stats ? '–' : c.value}
                </span>
                <span className="text-[12px] text-[var(--tmuted)] truncate">{c.sub}</span>
              </div>
            </button>
          ))}
        </div>
      </section>

      {/* ══ 6. เพิ่งเกิดขึ้น — การย้ายล่าสุด 5 รายการ (เดิม: การโอนย้ายล่าสุด) กดดูประวัติได้ ══ */}
      <div className="rounded-[var(--r5)] bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
        <div className="flex items-center justify-between px-4 py-3 border-b border-[var(--g100)]">
          <div className="flex items-center gap-2 text-[13px] font-semibold text-[var(--text)]">
            <span className="flex items-center justify-center w-7 h-7 rounded-lg bg-[var(--blue-l)] text-[var(--blue)]">
              <Icon name="local_shipping" size="sm" />
            </span>
            เพิ่งเกิดขึ้น
          </div>
          <button onClick={() => navigate('/asset')} title="ดูทั้งหมดในหน้าทะเบียนรายชิ้น" className="text-[12px] font-semibold text-[var(--blue)] hover:underline">
            ดูทั้งหมด ›
          </button>
        </div>
        {recent.length === 0 ? (
          <button
            onClick={() => setTransferCenterOpen(true)}
            title="เปิด Transfer Center — ย้ายอุปกรณ์"
            className="w-full px-4 py-5 flex items-center justify-center gap-2 text-[12px] text-[var(--tmuted)] hover:bg-[var(--surface2)] transition-colors focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--blue-glow)]"
          >
            {loading ? 'กำลังโหลด...' : (
              <>
                ยังไม่มีการย้ายวันนี้ — กดย้ายอุปกรณ์เพื่อเริ่ม
                <Icon name="local_shipping" size="sm" className="text-[var(--blue)]" />
              </>
            )}
          </button>
        ) : (
          <div className="divide-y divide-[var(--g100)]">
            {recent.slice(0, 5).map((r, idx) => {
              const chip = actionChip(r.action);
              const asset = nameBySerial.get(r.serialNumber);
              const deviceName = asset?.name || r.name || '';
              const hasSerial = r.serialNumber && r.serialNumber !== '-';
              const toText = r.to && r.to !== '-' ? ` · ไป ${r.to}` : '';
              return (
                <button
                  type="button"
                  key={`${r.serialNumber}-${idx}`}
                  onClick={() => r.serialNumber && setHistorySerial(r.serialNumber)}
                  title={`ประวัติของ ${r.serialNumber}${r.user && r.user !== '-' ? ` (โดย ${r.user})` : ''}`}
                  className="flex w-full items-center gap-3 px-4 py-2.5 text-left hover:bg-[var(--surface2)] transition-colors focus:outline-none focus-visible:ring-2 focus-visible:ring-inset focus-visible:ring-[var(--blue-glow)]"
                >
                  <span className={`flex items-center justify-center w-8 h-8 rounded-xl shrink-0 ${chip.cls}`}>
                    <Icon name={chip.icon} size="xs" />
                  </span>
                  <span className="flex-1 min-w-0">
                    {/* Visual Hierarchy (Step 2): ชื่ออุปกรณ์จริง = ตัวหนาหลัก · Serial ย่อยเป็น badge สีเทาเล็ก */}
                    <span className="block text-[13px] font-semibold text-[var(--text)] truncate">
                      {deviceName || r.serialNumber}
                    </span>
                    <span className="block text-[11px] text-[var(--tsub)] truncate" title={r.to ? `ไป ${r.to}` : undefined}>
                      {activityLabel(r.action)}{toText}
                    </span>
                  </span>
                  <span className="ml-auto shrink-0 flex flex-col items-end gap-0.5">
                    <span className="text-[11px] text-[var(--tmuted)] tabular-nums">{fmtWhen(r.date)}</span>
                    {deviceName && hasSerial && (
                      <span
                        className="max-w-[9rem] truncate rounded border border-[var(--g100)] bg-[var(--surface2)] px-1.5 text-[10px] font-mono text-[var(--tmuted)]"
                        title={`Serial ${r.serialNumber}`}
                      >
                        {r.serialNumber}
                      </span>
                    )}
                  </span>
                </button>
              );
            })}
          </div>
        )}
      </div>

      {/* ══ (เดิม: งานที่ทำบ่อย 4 การ์ด) — แทนด้วยปุ่มงานหลักวงกลม 5 ปุ่มด้านบนตามแนว B2 ══
          ทางเข้า /asset (ทะเบียนรายชิ้น) ยังครบ: เมนู sidebar "เพิ่มเติม" + ปุ่ม "ดูทะเบียนรายชิ้น"
          ในการ์ดโอนย้ายล่าสุด + ลิงก์ท้ายการ์ด "ลงทะเบียนล่าสุด" */}

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
            sub={`${farmTotalBundles} ตู้ติดตั้งที่ฟาร์ม · ${Math.max(0, bundles.length - farmTotalBundles)} ตู้ในคลัง · ${derived.bundleAssets} ชิ้นในชุด`}
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

      {/* ══ สถิติเชิงลึก — ลงทะเบียนล่าสุด (Step 3: ตัดการ์ด "สถานะอุปกรณ์" ทิ้งถาวร — ข้อมูลซ้ำซ้อนกับ
          การ์ดสรุปสถานะ 3 ใบด้านบน + Farm Monitor แล้ว และการ์ดเดิมมี dead space พื้นขาวขนาดใหญ่)
          การ์ด "ลงทะเบียนล่าสุด" จึงขยายเต็มความกว้าง (Full Width) ให้เลย์เอาต์สมดุลกับหน้าจอ ══ */}
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

      {/* (เดิม: สต็อกใกล้หมด) — ย้ายขึ้นเป็นแถบเตือนเหนือการ์ดสถานะ (กดกางดูรายการได้) ตามแนว B2 */}

      {/* (เดิม: การโอนย้ายล่าสุด) — ย้ายขึ้นเป็นการ์ด "เพิ่งเกิดขึ้น" ตามแนว B2
          ฟังก์ชันเดิมครบ: กดรายการ → ดูประวัติ (AssetHistoryModal) / ปุ่มดูทั้งหมด → /asset */}

      {/* เพิ่มอุปกรณ์ใหม่ (โฟลว์เดียว: เลือก/สร้าง Part → Serial อัตโนมัติ กันซ้ำ) */}
      {addOpen && (
        <AddDeviceModal
          open
          onClose={() => setAddOpen(false)}
          onDone={() => { setAddOpen(false); setDataVersion((v) => v + 1); }}
        />
      )}

      {/* Episode 3 → Clean Home: TransferPickModal เดิม → PODispatchPicker → ปัจจุบันคือ Transfer Center Modal (ด้านล่าง) */}

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

      {/* Transfer Center Modal — กดปุ่มวงกลม "ย้ายอุปกรณ์" → ป๊อปอัปกลางหน้าจอ 2 แท็บ
          (แท็บ 1: โอนย้ายตาม PO · แท็บ 2: โอนย้ายอุปกรณ์ทั่วไป นอก PO) — ข้อมูลใช้ preload ของหน้าแรก ไม่ดึง API ซ้ำ
          ย้ายสำเร็จ → ปิด Modal + toast แจ้งเตือน + bump dataVersion ให้หน้าแรกรีเฟรช */}
      {transferCenterOpen && (
        <TransferCenterModal
          open
          onClose={() => setTransferCenterOpen(false)}
          preloadAssets={assets}
          preloadBundles={bundles}
          onTransferSuccess={() => setDataVersion((v) => v + 1)}
        />
      )}

      {/* ToastHost ระดับหน้า — toast แจ้งผลโอนย้ายจาก Transfer Center ยังโชว์ต่อแม้ Modal ปิดแล้ว */}
      <ToastHost />
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

// chip สี/ไอคอนตามประเภทการโอนย้าย (Asset_History.action) — รองรับทั้งคำไทยและคำเทคนิคเดิม (Deploy/Recall)
function actionChip(action = '') {
  const a = String(action);
  const al = a.toLowerCase();
  if (a.includes('ลงทะเบียน')) return { icon: 'add_circle', cls: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]' };
  if (a.includes('ซ่อม')) return { icon: 'build', cls: 'bg-[var(--amber-l)] text-[var(--amber-d)]' };
  if (al.includes('deploy') || a.includes('ย้าย') || a.includes('โอน')) return { icon: 'local_shipping', cls: 'bg-[var(--blue-l)] text-[var(--blue)]' };
  if (al.includes('recall') || a.includes('คืนคลัง')) return { icon: 'undo', cls: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]' };
  if (a.includes('เข้า Bundle')) return { icon: 'grid_on', cls: 'bg-[var(--blue-l)] text-[var(--blue)]' };
  if (a.includes('ออกจาก Bundle')) return { icon: 'remove', cls: 'bg-[var(--amber-l)] text-[var(--amber-d)]' };
  return { icon: 'sync_alt', cls: 'bg-[var(--g200)] text-[var(--tsub)]' };
}

// ขจัดศัพท์เทคนิค (Step 2): แปลงคอลัมน์ action ดิบ เช่น "Deploy Bundle: BND-xxx" / "Recall Bundle: BND-xxx"
// ให้เป็นภาษาไทยหน้างานล้วนๆ เช่น "ย้ายทั้งชุดไปฟาร์ม" / "ดึงคืนเข้าคลังกลาง"
function activityLabel(action = '') {
  const a = String(action);
  if (/deploy/i.test(a)) return 'ย้ายทั้งชุดไปฟาร์ม';
  if (/recall/i.test(a)) return 'ดึงคืนเข้าคลังกลาง';
  if (a.includes('คืนคลัง')) return 'คืนเข้าคลังกลาง';
  if (a.includes('เพิ่มเข้า Bundle')) return 'จัดเข้าชุดติดตั้งฟาร์ม';
  if (a.includes('ลบออกจาก Bundle')) return 'นำออกจากชุดติดตั้งฟาร์ม';
  if (a.includes('ลงทะเบียน')) return 'เพิ่มอุปกรณ์ใหม่เข้าคลัง';
  if (a.includes('ซ่อม')) return 'ส่งซ่อม';
  if (a.includes('ย้าย') || a.includes('โอน')) return 'ย้ายตำแหน่งอุปกรณ์';
  return a; // action ที่ไม่รู้จัก → โชว์ข้อความเดิม (กันข้อมูลหาย)
}

// เวลากระชับ (Step 2): วันนี้ "18:44 น." · เมื่อวาน "เมื่อวาน 18:44" · เก่ากว่านั้น "9 ต.ค." (+ปี ถ้าต่างปี)
// (คอลัมน์ date ใน Asset_History เก็บแบบ toLocaleString('th-TH') เช่น "10/10/2568, 18:44:33")
function fmtWhen(raw = '') {
  const s = String(raw).trim();
  const m = /(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})[,\s]+(\d{1,2}):(\d{2})/.exec(s);
  if (!m) return s && s !== '-' ? s : '';
  let [, dd, mm, yy, hh, mi] = m.map(Number);
  if (yy < 100) yy += 2500;
  if (yy > 2400) yy -= 543; // พ.ศ. → ค.ศ.
  const when = new Date(yy, mm - 1, dd, hh, mi);
  if (Number.isNaN(when.getTime())) return s;
  const now = new Date();
  const dayStart = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  const diffDays = Math.round((dayStart(now) - dayStart(when)) / 86400000);
  const hm = `${String(hh).padStart(2, '0')}:${String(mi).padStart(2, '0')}`;
  if (diffDays === 0) return `${hm} น.`;
  if (diffDays === 1) return `เมื่อวาน ${hm}`;
  const md = when.toLocaleDateString('th-TH', { day: 'numeric', month: 'short' });
  return when.getFullYear() === now.getFullYear() ? md : `${md} ${String(when.getFullYear()).slice(2)}`;
}

// ── TransferPickModal (คอมเมนต์อดีต) — Episode 3: แทนด้วย PODispatchPicker แบบฝังหน้าแรก
//    → Clean Home (ล่าสุด): ถอดการ์ดฝังออก ใช้ TransferCenterModal ป๊อปอัปกลางจอจากปุ่ม "ย้ายอุปกรณ์" แทน ──

function Card({ icon, tone, title, children }) {
  return (
    <div className="rounded-[var(--r5)] bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-sm)] p-4 sm:p-5 flex flex-col">
      <div className="flex items-center gap-2 text-[13px] font-semibold text-[var(--text)] mb-3">
        <span className={`w-7 h-7 flex items-center justify-center rounded-lg ${TONE_BOX[tone] || TONE_BOX.blue}`}><Icon name={icon} size="sm" /></span>
        {title}
      </div>
      <div className="flex-1">{children}</div>
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
// Step 3 (Premium): ส่วนหัว Dark Slate (ชื่อฟาร์มขาวบนพื้น --ink) · ตัวเลขย่อยเป็น Glassmorphic Badges
// · Health Bar หนาขึ้น (h-2.5) พร้อม Glow Effect สีเขียว (เหลืองอำพันเมื่อมีของส่งซ่อม)
function FarmAccessCard({ farm, onOpen }) {
  const icon = FARM_TYPE_ICONS[farm.type] || 'help';
  const totalActive = farm.ok + farm.rep;
  const healthPct = totalActive > 0 ? Math.round((farm.ok / totalActive) * 100) : 100;
  return (
    <button
      onClick={onOpen}
      title={`เปิดรายละเอียดฟาร์ม ${farm.name} — Asset Breakdown + Bundle Control`}
      className="group text-left rounded-[var(--r5)] bg-[var(--surface)] border border-[var(--border)] shadow-[var(--sh-sm)] overflow-hidden hover:shadow-[var(--sh-md)] hover:-translate-y-0.5 transition-all focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--blue-glow)]"
    >
      {/* ส่วนหัว Dark Slate — ชื่อฟาร์มเด่น อ่านชัดทันที */}
      <div className="flex items-center gap-2.5 bg-[var(--ink)] px-4 py-3">
        <span
          aria-hidden="true"
          className="w-9 h-9 rounded-xl bg-white/10 ring-1 ring-white/15 text-white/90 flex items-center justify-center shrink-0 backdrop-blur-sm"
        >
          <Icon name={icon} size="sm" />
        </span>
        <div className="flex-1 min-w-0">
          <div className="text-[14px] font-bold text-white truncate">{farm.name}</div>
          <div className="text-[11px] text-white/55 truncate">{farm.type}{farm.province ? ` · ${farm.province}` : ''}</div>
        </div>
        <span className="text-white/50 group-hover:text-white group-hover:translate-x-0.5 transition-all shrink-0">
          <Icon name="chevron_right" size="sm" />
        </span>
      </div>

      <div className="p-4 pt-3">
        {/* ตัวเลขย่อย — Glassmorphic Badges: พื้นขาวโปร่งแสงเบลอบนแถบ slate อ่อน */}
        <div className="grid grid-cols-3 gap-2 rounded-xl bg-[var(--g100)]/80 ring-1 ring-[var(--g200)] p-1.5">
          <MiniNum icon="inventory_2" label="อุปกรณ์" value={farm.assets} />
          <MiniNum icon="grid_on" label="ตู้ชุด" value={farm.bundles} />
          <MiniNum icon="home_work" label="โรงเรือน" value={farm.houses} />
        </div>

        {/* Health Bar — เพิ่มความหนา (h-2.5) + Glow Effect */}
        <div className="mt-3">
          <div className="flex items-center justify-between text-[10.5px] mb-1.5">
            <span className="flex items-center gap-1 font-semibold text-[var(--emerald-d)]">
              <span className="w-1.5 h-1.5 rounded-full bg-[var(--emerald)] shadow-[0_0_6px_var(--emerald-b)]" /> ใช้งานได้ {farm.ok}
            </span>
            {farm.rep > 0 ? (
              <span className="flex items-center gap-1 text-[var(--amber-d)] font-semibold">
                <span className="w-1.5 h-1.5 rounded-full bg-[var(--amber)] shadow-[0_0_6px_var(--amber-b)]" /> ส่งซ่อม {farm.rep}
              </span>
            ) : (
              <span className="text-[var(--tmuted)]">สถานะปกติ</span>
            )}
          </div>
          <div className="h-2.5 rounded-full bg-[var(--g100)] ring-1 ring-[var(--g200)] overflow-hidden">
            <div
              className={`h-full rounded-full transition-all duration-700 ${
                farm.rep > 0
                  ? 'bg-gradient-to-r from-[var(--amber-d)] to-[var(--amber)] shadow-[0_0_12px_var(--amber-b)]'
                  : 'bg-gradient-to-r from-[var(--emerald-d)] to-[var(--emerald)] shadow-[0_0_14px_var(--emerald-b)]'
              }`}
              style={{ width: `${healthPct}%` }}
            />
          </div>
        </div>
      </div>
    </button>
  );
}

function MiniNum({ icon, label, value }) {
  // Glassmorphic Badge: พื้นขาวโปร่ง (white/60) + backdrop-blur + ขอบขาว + เงาใน/นอกบางๆ ให้มิติแบบแก้ว
  return (
    <div className="rounded-lg bg-white/60 backdrop-blur-md border border-white/90 shadow-[inset_0_1px_0_rgba(255,255,255,.85),0_2px_8px_rgba(15,23,42,.08)] px-2 py-1.5 min-w-0">
      <div className="flex items-center gap-1 text-[10px] text-[var(--tsub)]">
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
