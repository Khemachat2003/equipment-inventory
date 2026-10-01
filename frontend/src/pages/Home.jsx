// Home — หน้าแรกแบบ Workdesk (UX simplify)
// หลักการ: "1 หน้าจอ = งานที่ผู้ใช้อยากทำ" — ค้นหาได้ทุกอย่างจากที่เดียว
// ค้นหา → เห็นอุปกรณ์รายชิ้น/ของในคลัง → กด action ต่อได้ทันที (ย้าย/ประวัติ/ฉลาก/เบิก)
// ไม่แสดงกราฟ (กราฟเดิมอยู่ที่ /dashboard — ดู UI_FLAGS.charts ใน uiConfig.js)

import { useEffect, useMemo, useState } from 'react';
import axios from 'axios';
import { useNavigate } from 'react-router-dom';
import Icon from '../components/ui/Icon.jsx';
import StatPill from '../components/ui/StatPill.jsx';
import StatusBadge from '../components/ui/StatusBadge.jsx';
import LocationPath from '../components/ui/LocationPath.jsx';
import { useAuth } from '../context/AuthContext.jsx';

const IMAGE_URL = (code, ext) =>
  `https://cdn.jsdelivr.net/gh/Khemachat2003/stock-image@main/images/${code}.${ext || 'jpg'}?v=4`;

const QUICK_ACTIONS = [
  { to: '/stock', icon: 'inventory_2', box: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]', title: 'เบิก–คืนของ', desc: 'หยิบของจากคลังไปใช้ที่งาน หรือคืนกลับคลัง' },
  { to: '/scan', icon: 'document_scanner', box: 'bg-[var(--blue-l)] text-[var(--blue)]', title: 'ย้าย/โอนอุปกรณ์', desc: 'สแกนบาร์โค้ด (หรือพิมพ์รหัส) แล้วเลือกปลายทาง' },
  { to: '/qr', icon: 'qr_code_2', box: 'bg-[var(--amber-l)] text-[var(--amber-d)]', title: 'พิมพ์ฉลาก', desc: 'พิมพ์ฉลาก QR / Barcode ติดอุปกรณ์' },
];

export default function Home() {
  const navigate = useNavigate();
  const { user } = useAuth();
  const [q, setQ] = useState('');
  const [assets, setAssets] = useState([]);
  const [stock, setStock] = useState([]);
  const [stats, setStats] = useState(null);
  const [recent, setRecent] = useState([]);
  const [loading, setLoading] = useState(true);

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
      if (h.status === 'fulfilled') setRecent((h.value.data || []).slice(0, 5));
      setLoading(false);
    })();
    return () => { alive = false; };
  }, []);

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

  const hour = new Date().getHours();
  const greeting = hour < 12 ? 'สวัสดีตอนเช้า' : hour < 18 ? 'สวัสดีตอนบ่าย' : 'สวัสดีตอนเย็น';

  return (
    <div className="space-y-5 max-w-4xl">
      {/* ══ ทักทาย + ค้นหาอะไรก็ได้ ══ */}
      <div>
        <div className="text-[15px] font-semibold text-[var(--text)]">
          {greeting}{user?.username ? `, ${user.username}` : ''} 👋
        </div>
        <div className="text-[13px] text-[var(--tmuted)] mt-0.5">วันนี้ต้องทำอะไรครับ? ค้นหาอุปกรณ์ได้จากช่องด้านล่าง</div>
        <div className="mt-3 flex gap-2">
          <div className="relative flex-1">
            <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[var(--tmuted)] pointer-events-none">
              <Icon name="search" size="sm" />
            </span>
            <input
              value={q}
              onChange={(e) => setQ(e.target.value)}
              onKeyDown={(e) => { if (e.key === 'Escape') setQ(''); }}
              placeholder="พิมพ์ชื่อ / รหัส / Serial อุปกรณ์..."
              className="w-full h-11 pl-9 pr-3 rounded-xl border border-[var(--g200)] bg-white text-[14px] focus:outline-none focus:border-[var(--blue)] shadow-[var(--sh-sm)]"
            />
          </div>
          <button
            onClick={() => navigate('/scan')}
            title="สแกนบาร์โค้ด / QR"
            className="h-11 px-4 rounded-xl bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)] flex items-center gap-1.5 shrink-0"
          >
            <Icon name="qr_code_scanner" size="sm" /> <span className="hidden min-[430px]:inline">สแกน</span>
          </button>
        </div>
      </div>

      {/* ══ ผลลัพธ์ค้นหา — ทำต่อได้จากที่นี่เลย ไม่ต้องเปลี่ยนหน้า ══ */}
      {results && (
        <div className="space-y-3">
          {results.total === 0 ? (
            <div className="rounded-2xl bg-white border border-[var(--g200)] p-6 text-center text-[13px] text-[var(--tmuted)]">
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
                      className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-3.5 flex flex-wrap items-center gap-3"
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
                    <div key={i.code} className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-3.5 flex flex-wrap items-center gap-3">
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
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          {QUICK_ACTIONS.map((a) => (
            <button
              key={a.to}
              onClick={() => navigate(a.to)}
              className="text-left rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-4 hover:border-[var(--blue)] hover:shadow-[var(--sh-md)] transition-all"
            >
              <span className={`flex items-center justify-center w-10 h-10 rounded-xl ${a.box}`}>
                <Icon name={a.icon} size="md" />
              </span>
              <div className="text-[14px] font-bold text-[var(--text)] mt-2.5">{a.title}</div>
              <div className="text-[12px] text-[var(--tmuted)] mt-1 leading-snug">{a.desc}</div>
              <div className="text-[12px] font-semibold text-[var(--blue)] mt-2.5 flex items-center gap-1">
                เปิด <Icon name="chevron_right" size="xs" />
              </div>
            </button>
          ))}
        </div>
      </div>

      {/* ══ สรุปภาพรวม (ไม่มีกราฟ — กราฟเดิมอยู่ที่ /dashboard) ══ */}
      <div>
        <div className="text-[11px] font-semibold uppercase tracking-wider text-[var(--tmuted)] mb-2">สรุปภาพรวม</div>
        {/* แถว 1 — วันนี้ */}
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
          <StatPill icon="inventory_2" tone="blue" label="ในคลัง (Office)" value={loading ? '…' : (stats?.totalOffice ?? 0)} />
          <StatPill icon="factory" tone="amber" label="อยู่ที่ Site" value={loading ? '…' : (stats?.totalSite ?? 0)} />
          <StatPill icon="trending_up" tone="green" label="เบิกวันนี้" value={loading ? '…' : (stats?.todayBorrow ?? 0)} />
          <StatPill icon="trending_down" tone="red" label="คืนวันนี้" value={loading ? '…' : (stats?.todayReturn ?? 0)} />
        </div>
        {/* แถว 2 — ฟาร์มและอุปกรณ์รายชิ้น (ให้ผู้ใช้ใหม่เห็นภาพรวมทันที) */}
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mt-3">
          <StatPill icon="place" tone="green" label="ฟาร์มทั้งหมด" value={loading ? '…' : (stats?.totalFarms ?? 0)} />
          <StatPill icon="precision_manufacturing" tone="blue" label="ติดตั้งที่ฟาร์ม" value={loading ? '…' : (stats?.totalFarmAssets ?? 0)} />
          <StatPill icon="inventory_2" tone="amber" label="รอติดตั้ง (สต็อก)" value={loading ? '…' : (stats?.totalStockAssets ?? 0)} />
          <StatPill icon="devices" tone="blue" label="รายชิ้นทั้งหมด" value={loading ? '…' : (stats?.assets?.length ?? 0)} />
        </div>
      </div>

      {/* ══ การโอนย้ายล่าสุด (จาก Asset_History — เฉพาะการย้าย/โอนอุปกรณ์รายชิ้น) ══ */}
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
            {recent.map((r, idx) => (
              <div key={idx} className="flex items-center gap-2.5 px-4 py-2.5 text-[12px]">
                <span className="flex items-center justify-center w-6 h-6 rounded-lg bg-[var(--blue-l)] text-[var(--blue)] shrink-0">
                  <Icon name="sync_alt" size="xs" />
                </span>
                <span className="font-semibold text-[var(--text)] shrink-0 hidden min-[430px]:inline">{r.action}</span>
                <span className="font-mono text-[var(--blue)] truncate max-w-[180px]" title={r.serialNumber}>{r.serialNumber}</span>
                <span className="truncate text-[var(--tsub)] hidden sm:inline" title={r.to}>→ {r.to}</span>
                <span className="ml-auto shrink-0 text-[var(--tmuted)]">{r.date}</span>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}