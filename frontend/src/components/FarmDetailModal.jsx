import { useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import Icon from './ui/Icon.jsx';
import StatusBadge from './ui/StatusBadge.jsx';
import LocationPath from './ui/LocationPath.jsx';
import { buildBundleFarmMap, farmAssetsOf } from '../utils/farmMonitor.js';

const TYPE_ICONS = { 'สัตว์ปีก': 'egg', 'สัตว์บก': 'pets', 'สุกร': 'agriculture', 'อื่นๆ': 'category', 'ไม่ระบุ': 'help' };

/**
 * FarmDetailModal — เปิดรายละเอียดฟาร์มทันทีจากการ์ด Quick Farm Access บนหน้าแรก
 * (Asset Breakdown + Bundle Control) — ไม่ต้องเปลี่ยนหน้าไป Farm Monitor
 *
 * กฎการนับ/คัดรายการ = กฎเดียวกับ /farm (utils/farmMonitor.js)
 */
export default function FarmDetailModal({ farm, assets = [], bundles = [], loading = false, onClose, onTransfer, onHistory }) {
  const navigate = useNavigate();
  const [q, setQ] = useState('');

  const bundleFarm = useMemo(() => buildBundleFarmMap(bundles), [bundles]);

  const farmAssets = useMemo(() => farmAssetsOf(farm?.name, assets, bundleFarm), [farm, assets, bundleFarm]);
  const farmBundles = useMemo(
    () => (bundles || []).filter((b) => b && b.status === 'Deployed' && (b.location || '').trim() === farm?.name),
    [bundles, farm],
  );

  // ค้นหาภายในฟาร์ม: Serial / ชื่อ / รหัส / AssetID / PO / ล็อต / Supplier / โรงเรือน
  const filtered = useMemo(() => {
    const k = (q || '').toLowerCase().trim();
    if (!k) return farmAssets;
    return farmAssets.filter((a) =>
      [a.serialNumber, a.name, a.code, a.assetId, a.poNumber, a.batchId, a.supplier, a.houseName, a.houseId]
        .some((v) => (v || '').toLowerCase().includes(k)));
  }, [q, farmAssets]);

  const filteredBundles = useMemo(() => {
    const k = (q || '').toLowerCase().trim();
    if (!k) return farmBundles;
    return farmBundles.filter((b) =>
      [b.bundleId, b.bundleName].some((v) => (v || '').toLowerCase().includes(k)));
  }, [q, farmBundles]);

  // เรียงตามโรงเรือน → Serial เพื่อให้ช่างเดินตรวจหน้างานได้ตามลำดับอาคาร
  const sorted = useMemo(() => [...filtered].sort((a, b) => {
    const ha = (a.houseName || a.houseId || '').toString();
    const hb = (b.houseName || b.houseId || '').toString();
    return ha.localeCompare(hb, 'th') || (a.serialNumber || '').localeCompare(b.serialNumber || '');
  }), [filtered]);

  if (!farm) return null;
  const typeIcon = TYPE_ICONS[farm.type] || 'help';

  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto bg-black/40 p-3 sm:p-4 backdrop-blur-sm" onClick={onClose}>
      <section
        role="dialog"
        aria-modal="true"
        aria-labelledby="farm-detail-title"
        className="my-3 sm:my-8 w-full max-w-3xl overflow-hidden rounded-2xl bg-[var(--surface)] shadow-[var(--sh-lg)]"
        onClick={(e) => e.stopPropagation()}
      >
        {/* ── Header: ชื่อฟาร์ม + สรุปสั้น ── */}
        <header className="flex items-start justify-between gap-3 border-b border-[var(--g100)] px-4 py-3.5 sm:px-5">
          <div className="flex items-start gap-3 min-w-0">
            <span className="flex items-center justify-center w-10 h-10 rounded-xl bg-[var(--emerald-l)] text-[var(--emerald-d)] shrink-0">
              <Icon name={typeIcon} size="md" />
            </span>
            <div className="min-w-0">
              <h2 id="farm-detail-title" className="text-[16px] font-bold text-[var(--text)] truncate">{farm.name}</h2>
              <div className="text-[11px] text-[var(--tmuted)] truncate">
                {farm.type}{farm.province ? ` · ${farm.province}` : ''}
              </div>
            </div>
          </div>
          <button
            onClick={onClose}
            aria-label="ปิดรายละเอียดฟาร์ม"
            className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)] hover:text-[var(--text)]"
          >
            <Icon name="close" size="sm" />
          </button>
        </header>

        {/* ── สรุปตัวเลข + ช่องค้นหา ── */}
        <div className="px-4 pt-3.5 sm:px-5 space-y-3">
          <div className="grid grid-cols-4 gap-2">
            <MiniStat icon="inventory_2" tone="blue" label="อุปกรณ์" value={farmAssets.length} />
            <MiniStat icon="grid_on" tone="purple" label="ตู้ชุด" value={farmBundles.length} />
            <MiniStat icon="check_circle" tone="green" label="ใช้งานได้" value={farmAssets.filter((a) => (a.status || '').includes('ใช้งานได้')).length} />
            <MiniStat icon="build" tone="red" label="ส่งซ่อม" value={farmAssets.filter((a) => (a.status || '').includes('ซ่อม')).length} />
          </div>
          <div className="relative">
            <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[var(--tmuted)] pointer-events-none">
              <Icon name="search" size="sm" />
            </span>
            <input
              value={q}
              onChange={(e) => setQ(e.target.value)}
              placeholder="ค้นหาในฟาร์มนี้ — Serial / ชื่อ / รหัส / PO / โรงเรือน..."
              className="w-full h-10 pl-9 pr-3 rounded-xl border border-[var(--g200)] bg-[var(--surface2)] text-[13px] text-[var(--text)] placeholder:text-[var(--tmuted)] focus:outline-none focus:ring-2 focus:ring-[var(--blue-glow)]"
            />
          </div>
        </div>

        {/* ── Body ── */}
        <div className="max-h-[62vh] overflow-y-auto px-4 py-3.5 sm:px-5 space-y-4">
          {/* Bundle Control — ตู้ควบคุมที่ติดตั้งอยู่ที่ฟาร์มนี้ */}
          <section>
            <div className="flex items-center gap-2 mb-2 text-[12px] font-semibold uppercase tracking-wider text-[var(--tmuted)]">
              <Icon name="grid_on" size="xs" /> ตู้ควบคุม (Bundle) ที่ฟาร์มนี้
            </div>
            {filteredBundles.length === 0 ? (
              <div className="rounded-xl border border-dashed border-[var(--g200)] px-3 py-3 text-center text-[12px] text-[var(--tmuted)]">
                ไม่มีตู้ชุดติดตั้งอยู่ที่ฟาร์มนี้
              </div>
            ) : (
              <div className="space-y-2">
                {filteredBundles.map((b) => (
                  <div key={b.bundleId} className="flex flex-wrap items-center gap-x-3 gap-y-2 rounded-xl border border-[var(--g200)] bg-[var(--surface2)] px-3 py-2.5">
                    <span className="flex items-center justify-center w-8 h-8 rounded-lg bg-[var(--purple-l)] text-[var(--purple)] shrink-0">
                      <Icon name="grid_on" size="sm" />
                    </span>
                    <div className="flex-1 min-w-[160px]">
                      <div className="text-[13px] font-bold text-[var(--text)] truncate">{b.bundleName || b.bundleId}</div>
                      <div className="text-[11px] font-mono text-[var(--tmuted)]">{b.bundleId}</div>
                    </div>
                    <StatusBadge status={b.status} />
                    <span className="text-[11px] text-[var(--tsub)] whitespace-nowrap">
                      สมาชิก <strong className="tabular-nums text-[var(--text)]">{(b.assetIds || []).length}</strong> ชิ้น
                    </span>
                    <button
                      onClick={() => navigate(`/bundle?open=${encodeURIComponent(b.bundleId)}`)}
                      title="เปิดหน้าชุดเพื่อจัดการ — เพิ่ม/ถอนอุปกรณ์ ย้ายทั้งชุด หรือคืนเข้าคลัง"
                      className="h-8 px-2.5 rounded-lg bg-[var(--purple-l)] border border-[var(--purple)]/30 text-[var(--purple)] text-[12px] font-semibold hover:bg-[var(--purple)] hover:text-white flex items-center gap-1"
                    >
                      <Icon name="tune" size="xs" /> จัดการชุด
                    </button>
                  </div>
                ))}
              </div>
            )}
          </section>

          {/* Asset Breakdown — อุปกรณ์ที่ติดตั้งอยู่ที่ฟาร์มนี้ */}
          <section>
            <div className="flex items-center justify-between gap-2 mb-2">
              <div className="flex items-center gap-2 text-[12px] font-semibold uppercase tracking-wider text-[var(--tmuted)]">
                <Icon name="sensors" size="xs" /> อุปกรณ์ที่ติดตั้ง ({sorted.length})
              </div>
              {q.trim() && <span className="text-[11px] text-[var(--tmuted)]">จากทั้งหมด {farmAssets.length} ชิ้น</span>}
            </div>
            {loading ? (
              <div className="rounded-xl border border-[var(--g200)] px-3 py-6 text-center text-[12px] text-[var(--tmuted)]">กำลังโหลด...</div>
            ) : sorted.length === 0 ? (
              <div className="rounded-xl border border-dashed border-[var(--g200)] px-3 py-6 text-center text-[12px] text-[var(--tmuted)]">
                <div className="flex justify-center mb-1.5"><Icon name={farmAssets.length === 0 ? 'factory' : 'search_off'} size="lg" /></div>
                {farmAssets.length === 0
                  ? 'ยังไม่มีอุปกรณ์ติดตั้งที่ฟาร์มนี้'
                  : `ไม่พบอุปกรณ์ที่ตรงกับ "${q.trim()}" ในฟาร์มนี้`}
              </div>
            ) : (
              <div className="space-y-2">
                {sorted.map((a) => (
                  <div key={`${a.assetId}-${a.serialNumber || ''}`} className="rounded-xl border border-[var(--g200)] px-3 py-2.5 hover:border-[var(--blue-b)] transition-colors">
                    <div className="flex flex-wrap items-center gap-x-3 gap-y-1.5">
                      <div className="flex-1 min-w-[180px]">
                        <div className="flex items-center gap-2 flex-wrap">
                          <span className="text-[13px] font-semibold text-[var(--text)] truncate max-w-full">{a.name || a.code}</span>
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
                        <div className="text-[11px] font-mono text-[var(--blue)] mt-0.5 truncate">{a.serialNumber || a.assetId}</div>
                      </div>
                      <div className="flex items-center gap-1.5 flex-wrap">
                        {onTransfer && (
                          <button
                            onClick={() => onTransfer({ serial: a.serialNumber || a.assetId || '', current: { status: a.status, location: a.location, siteName: a.siteName, user: a.user } })}
                            disabled={!a.serialNumber && !a.assetId}
                            title="ย้าย/โอนอุปกรณ์ — เปิดฟอร์มทันที"
                            className="h-8 px-2.5 rounded-lg bg-[var(--blue)] text-white text-[12px] font-semibold hover:bg-[var(--blue-d)] disabled:opacity-40 flex items-center gap-1"
                          >
                            <Icon name="local_shipping" size="xs" /> ย้าย
                          </button>
                        )}
                        {onHistory && (
                          <button
                            onClick={() => a.serialNumber && onHistory(a.serialNumber)}
                            disabled={!a.serialNumber}
                            title="ดูประวัติการเคลื่อนไหว"
                            className="h-8 px-2.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)] hover:bg-[var(--surface2)] disabled:opacity-40 flex items-center gap-1"
                          >
                            <Icon name="history" size="xs" /> ประวัติ
                          </button>
                        )}
                        <a
                          href={`/qr?serial=${encodeURIComponent(a.serialNumber || '')}`}
                          target="_blank"
                          rel="noreferrer"
                          title="พิมพ์ฉลาก"
                          className="h-8 px-2.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)] hover:bg-[var(--surface2)] flex items-center gap-1"
                        >
                          <Icon name="qr_code" size="xs" /> ฉลาก
                        </a>
                      </div>
                    </div>
                    <LocationPath
                      className="mt-1.5"
                      variant="primary"
                      siteName={a.siteName}
                      houseName={a.houseName}
                      houseId={a.houseId}
                      location={a.location}
                      bundleId={a.bundleId}
                    />
                  </div>
                ))}
              </div>
            )}
          </section>
        </div>

        {/* ── Footer ── */}
        <footer className="flex items-center justify-between gap-3 border-t border-[var(--g100)] px-4 py-3 sm:px-5">
          <span className="text-[11px] text-[var(--tmuted)] hidden sm:block">ข้อมูลอัปเดตจากทะเบียนอุปกรณ์ล่าสุด</span>
          <button
            onClick={() => navigate('/farm')}
            className="h-8 px-3 rounded-lg border border-[var(--g300)] text-[12px] font-semibold text-[var(--tsub)] hover:border-[var(--blue)] hover:text-[var(--blue)] flex items-center gap-1 ml-auto"
          >
            <Icon name="monitoring" size="xs" /> เปิด Farm Monitor เต็มหน้า
          </button>
        </footer>
      </section>
    </div>
  );
}

function MiniStat({ icon, tone, label, value }) {
  const TONE = {
    green: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]',
    blue: 'bg-[var(--blue-l)] text-[var(--blue)]',
    amber: 'bg-[var(--amber-l)] text-[var(--amber-d)]',
    red: 'bg-[var(--red-l)] text-[var(--red)]',
    purple: 'bg-[var(--purple-l)] text-[var(--purple)]',
  };
  return (
    <div className="rounded-xl bg-[var(--surface2)] border border-[var(--g100)] px-2 py-2 text-center">
      <span className={`mx-auto mb-1 flex h-6 w-6 items-center justify-center rounded-md ${TONE[tone] || TONE.blue}`}>
        <Icon name={icon} size="xs" />
      </span>
      <div className="text-[15px] font-bold text-[var(--text)] tabular-nums leading-none">{value}</div>
      <div className="mt-1 text-[10px] text-[var(--tmuted)]">{label}</div>
    </div>
  );
}