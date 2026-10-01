// Home — หน้าแรกแบบ Workdesk (UX simplify)
// หลักการ: "1 หน้าจอ = งานที่ผู้ใช้อยากทำ" — ค้นหาได้ทุกอย่างจากที่เดียว
// ค้นหา → เห็นอุปกรณ์รายชิ้น/ของในคลัง → กด action ต่อได้ทันที (ย้าย/ประวัติ/ฉลาก/เบิก)
// ไม่แสดงกราฟ chart.js (กราฟเดิมอยู่ที่ /dashboard)

import { useEffect, useMemo, useState } from 'react';
import axios from 'axios';
import { useNavigate } from 'react-router-dom';
import Icon from '../components/ui/Icon.jsx';
import StatusBadge from '../components/ui/StatusBadge.jsx';
import LocationPath from '../components/ui/LocationPath.jsx';
import AddDeviceModal from '../components/AddDeviceModal.jsx';

const IMAGE_URL = (code, ext) =>
  `https://cdn.jsdelivr.net/gh/Khemachat2003/stock-image@main/images/${code}.${ext || 'jpg'}?v=4`;

const QUICK_ACTIONS = [
  { to: '/asset', icon: 'list_alt', box: 'bg-[var(--purple-l)] text-[var(--purple)]', title: 'ทะเบียนรายชิ้น', desc: 'ดู/ค้นหารายชิ้นทั้งหมด — สถานะ, ที่ตั้ง และประวัติ' },
  { to: '/scan', icon: 'document_scanner', box: 'bg-[var(--blue-l)] text-[var(--blue)]', title: 'ย้าย/โอนอุปกรณ์', desc: 'สแกนบาร์โค้ด (หรือพิมพ์รหัส) แล้วเลือกปลายทาง' },
  { to: '/stock', icon: 'inventory_2', box: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]', title: 'เบิก–คืนของ', desc: 'หยิบของจากคลังไปใช้ที่งาน หรือคืนกลับคลัง' },
  { to: '/qr', icon: 'qr_code_2', box: 'bg-[var(--amber-l)] text-[var(--amber-d)]', title: 'พิมพ์ฉลาก', desc: 'พิมพ์ฉลาก QR / Barcode ติดอุปกรณ์' },
];



export default function Home() {
  const navigate = useNavigate();
  const [q, setQ] = useState('');
  const [assets, setAssets] = useState([]);
  const [stock, setStock] = useState([]);
  const [stats, setStats] = useState(null);
  const [recent, setRecent] = useState([]);
  const [loading, setLoading] = useState(true);
  const [addOpen, setAddOpen] = useState(false);
  const [dataVersion, setDataVersion] = useState(0); // เพิ่มของเสร็จ → โหลดสถิติใหม่

  useEffect(() => {
    let alive = true;
    (async () => {
      // allSettled: API ตัวใดล้มเหลว (เช่น quota/เน็ต) → ส่วนอื่นยังแสดงได้
      const [a, s, d, h] = await Promise.allSettled([
        axios.get('/api/assets'),
        axios.get('/api/stock'),
        axios.get('/api/dashboard-full'),
        axios.get('/api/asset-history-recent'),
      ]);
      if (!alive) return;
      if (a.status === 'fulfilled') setAssets(a.value.data || []);
      if (s.status === 'fulfilled') setStock(s.value.data || []);
      if (d.status === 'fulfilled') setStats(d.value.data || null);
      if (h.status === 'fulfilled') setRecent((h.value.data || []).slice(0, 8));
      setLoading(false);
    })();
    return () => { alive = false; };
  }, [dataVersion]);

  const results = useMemo(() => {
    const k = q.trim().toLowerCase();
    if (!k) return null;
    const hitAssets = assets
      .filter((a) => [a.name, a.code, a.assetId, a.serialNumber].some((v) => (v || '').toLowerCase().includes(k)))
      .slice(0, 8);
    const hitStock = stock
      .filter((i) => [i.code, i.name].some((v) => (v || '').toLowerCase().includes(k)))
      .slice(0, 6);
    return { assets: hitAssets, stock: hitStock, total: hitAssets.length + hitStock.length };
  }, [q, assets, stock]);

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
                placeholder="พิมพ์ชื่อ / รหัส / Serial อุปกรณ์..."
                className="w-full h-12 pl-10 pr-3 rounded-xl border border-transparent bg-white text-[14px] text-[var(--text)] focus:outline-none focus:ring-2 focus:ring-white/50 shadow-[var(--sh-sm)] placeholder:text-[var(--tmuted)]"
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
                <HeroKpi icon="grid_on" glow="rgba(27,108,168,.35)" label="อยู่ใน Bundle" value={derived.bundleAssets} unit="ชิ้น" />
              </>
            )}
          </div>
        </div>
      </div>

      {/* ══ ผลลัพธ์ค้นหา — ทำต่อได้จากที่นี่เลย ไม่ต้องเปลี่ยนหน้า ══ */}
      {results && (
        <div className="space-y-3">
          {results.total === 0 ? (
            <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-6 text-center text-[13px] text-[var(--tmuted)]">
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
                      className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-3.5 flex flex-wrap items-center gap-3 hover:border-[var(--blue-b)] transition-colors"
                    >
                      <div className="flex-1 min-w-[220px]">
                        <div className="flex items-center gap-2">
                          <span className="text-[13px] font-semibold text-[var(--text)] truncate">{a.name || a.code}</span>
                          <StatusBadge status={a.status} />
                        </div>
                        <div className="text-[11px] font-mono text-[var(--blue)] mt-0.5 truncate">
                          {a.serialNumber || a.assetId || a.code}
                        </div>
                        <LocationPath
                          className="mt-1"
                          siteName={a.siteName}
                          houseName={a.houseName}
                          houseId={a.houseId}
                          location={a.location}
                          bundleId={a.bundleName}
                        />
                      </div>
                      <div className="flex items-center gap-1.5 flex-wrap">
                        <button
                          onClick={() => navigate(`/scan?serial=${encodeURIComponent(a.serialNumber || a.assetId || '')}`)}
                          title="ย้าย/โอนอุปกรณ์"
                          className="h-8 px-2.5 rounded-lg bg-[var(--blue)] text-white text-[12px] font-semibold hover:bg-[var(--blue-d)] flex items-center gap-1"
                        >
                          <Icon name="local_shipping" size="sm" /> ย้าย
                        </button>
                        <a
                          href={`/trace/${encodeURIComponent(a.serialNumber || '')}`}
                          target="_blank"
                          rel="noreferrer"
                          title="ดูประวัติการเคลื่อนไหว"
                          className="h-8 px-2.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)] hover:bg-[var(--surface2)] flex items-center gap-1"
                        >
                          <Icon name="history" size="sm" /> ประวัติ
                        </a>
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
                    <div key={i.code} className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-3.5 flex flex-wrap items-center gap-3 hover:border-[var(--emerald-b)] transition-colors">
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

      {/* ══ งานที่ทำบ่อย ══ */}
      <div>
        <div className="text-[11px] font-semibold uppercase tracking-wider text-[var(--tmuted)] mb-2">งานที่ทำบ่อย</div>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
          {QUICK_ACTIONS.map((a) => (
            <button
              key={a.title}
              onClick={() => navigate(a.to)}
              className="group text-left rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-4 hover:border-[var(--blue)] hover:shadow-[var(--sh-md)] hover:-translate-y-0.5 transition-all"
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

      {/* ══ สถิติเชิงลึก — ลงทะเบียนล่าสุด + สถานะอุปกรณ์ | อุปกรณ์และฟาร์ม ══ */}
      <div className="grid grid-cols-1 xl:grid-cols-[1fr_360px] gap-4 items-stretch">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {/* Serial ลงทะเบียนล่าสุด (parse วันที่จาก Serial) */}
          <Card icon="app_registration" tone="purple" title="ลงทะเบียนล่าสุด">
            {!loading && derived.recentSerials.length > 0 ? (
              <div className="space-y-2">
                {derived.recentSerials.map((a) => (
                  <a
                    key={a.serialNumber}
                    href={`/trace/${encodeURIComponent(a.serialNumber || '')}`}
                    target="_blank"
                    rel="noreferrer"
                    title="ดูประวัติการเคลื่อนไหว (Trace)"
                    className="flex items-center gap-2.5 rounded-xl border border-[var(--g200)] px-2.5 py-2 hover:border-[var(--purple)] hover:bg-[var(--purple-l)]/30 transition-colors"
                  >
                    <div className="flex-1 min-w-0">
                      <div className="text-[11px] font-mono text-[var(--purple)] truncate" title={a.serialNumber}>{a.serialNumber}</div>
                      <div className="text-[12px] font-semibold text-[var(--text)] truncate">{a.name || a.code}</div>
                    </div>
                    <StatusBadge status={a.status} />
                    <Icon name="chevron_right" size="xs" className="text-[var(--tmuted)] shrink-0" />
                  </a>
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

        {/* อุปกรณ์และฟาร์ม — รวมสถิติ + % ติดตั้ง + อันดับฟาร์ม */}
        <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-5 flex flex-col">
          <div className="flex items-center gap-2 text-[13px] font-semibold text-[var(--text)] mb-2">
            <span className="w-7 h-7 flex items-center justify-center rounded-lg bg-[var(--purple-l)] text-[var(--purple)]"><Icon name="devices" size="sm" /></span>
            อุปกรณ์และฟาร์ม
          </div>
          {loading ? (
            <SkeletonRows rows={4} />
          ) : (
            <>
              <div className="divide-y divide-[var(--g100)]">
                <MiniStat icon="agriculture" tone="green" label="ฟาร์มทั้งหมด" value={stats?.totalFarms ?? 0} />
                <MiniStat icon="precision_manufacturing" tone="blue" label="ติดตั้งที่ฟาร์ม" value={stats?.totalFarmAssets ?? 0} />
                <MiniStat icon="inventory_2" tone="amber" label="รอติดตั้ง (สต็อก)" value={stats?.totalStockAssets ?? 0} />
                <MiniStat icon="category" tone="purple" label="รายชิ้นทั้งหมด" value={derived.totalAssets} />
                <MiniStat icon="grid_on" tone="blue" label="ชุดอุปกรณ์ (Bundle)" value={derived.bundleCount} />
              </div>

              {/* % ติดตั้งแล้ว */}
              <div className="mt-3">
                <div className="flex items-center justify-between text-[11px] text-[var(--tsub)]">
                  <span>สัดส่วนติดตั้งแล้ว</span>
                  <span className="font-bold text-[var(--text)] tabular-nums">{derived.installedPct}%</span>
                </div>
                <div className="h-2 rounded-full bg-[var(--g100)] mt-1.5 overflow-hidden">
                  <div
                    className="h-full rounded-full transition-all duration-700"
                    style={{ width: `${derived.installedPct}%`, background: 'linear-gradient(90deg, var(--blue), var(--emerald))' }}
                  />
                </div>
              </div>

              {/* อันดับฟาร์ม */}
              {(stats?.topFarms || []).length > 0 && (
                <div className="mt-4 pt-3 border-t border-[var(--g100)]">
                  <div className="flex items-center gap-1.5 text-[11px] font-semibold uppercase tracking-wider text-[var(--tmuted)] mb-2">
                    <Icon name="workspace_premium" size="xs" /> อันดับฟาร์มที่ติดตั้งสูงสุด
                  </div>
                  <div className="space-y-2">
                    {stats.topFarms.slice(0, 3).map((f, i) => (
                      <div key={f.name} className="flex items-center gap-2.5">
                        <span className={`flex items-center justify-center w-5 h-5 rounded-md text-[10px] font-bold shrink-0 ${i === 0 ? 'bg-[var(--amber)] text-white' : 'bg-[var(--surface2)] text-[var(--tsub)] border border-[var(--g200)]'}`}>#{i + 1}</span>
                        <span className="flex-1 text-[12px] text-[var(--text)] truncate">{f.name}</span>
                        <span className="text-[12px] font-bold tabular-nums">{f.count}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </>
          )}
        </div>
      </div>

      {/* ══ สต็อกใกล้หมด — เตือนล่วงหน้าก่อนของหมดจริง ══ */}
      <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
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
      <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
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
                <div key={idx} className="flex items-center gap-2.5 px-4 py-2.5 text-[12px] hover:bg-[var(--surface2)] transition-colors">
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
                </div>
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
    <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-4 sm:p-5 flex flex-col">
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

function MiniStat({ icon, tone, label, value }) {
  return (
    <div className="flex items-center gap-3 py-2.5">
      <span className={`w-9 h-9 rounded-xl flex items-center justify-center shrink-0 ${TONE_BOX[tone] || TONE_BOX.blue}`}>
        <Icon name={icon} size="sm" />
      </span>
      <div className="flex-1 min-w-0 text-[12px] text-[var(--tsub)]">{label}</div>
      <div className="text-[17px] font-bold text-[var(--text)] tabular-nums">{value}</div>
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