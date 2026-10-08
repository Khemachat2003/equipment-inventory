import { useEffect, useState, useMemo, useCallback } from 'react';
import axios from 'axios';
import { useNavigate } from 'react-router-dom';
import Icon from '../components/ui/Icon.jsx';
import StatusBadge from '../components/ui/StatusBadge.jsx';
import StatPill from '../components/ui/StatPill.jsx';
import TransferModal from '../components/TransferModal.jsx';
import LocationPath from '../components/ui/LocationPath.jsx';
import { useBusy, BusyOverlay } from '../components/ui/Busy.jsx';
import { suggestSiteId, suggestHouseId } from '../utils/farmId.js';

const TYPE_ICONS = { 'สัตว์ปีก': 'egg', 'สัตว์บก': 'pets', 'สุกร': 'agriculture', 'อื่นๆ': 'category', 'ไม่ระบุ': 'help' };
const FARM_TYPES = ['สัตว์ปีก', 'สัตว์บก', 'สุกร', 'อื่นๆ'];

export default function Farm() {
  const navigate = useNavigate();

  const [assets, setAssets] = useState([]);
  const [bundles, setBundles] = useState([]);
  const [farmSearch, setFarmSearch] = useState('');
  const [currentFarm, setCurrentFarm] = useState('ALL');
  const [assetSearch, setAssetSearch] = useState('');
  const [transfer, setTransfer] = useState(null);
  const [addLocOpen, setAddLocOpen] = useState(false);
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

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-3">
        {/* เพิ่มฟาร์ม/โรงเรือน — ทุกคนที่ล็อกอินเพิ่มได้ (เดิม gate ไว้เฉพาะ admin ทำให้ user ติดขั้นตอนโอนย้าย) */}
        {/* เพิ่มฟาร์ม/โรงเรือน — ปุ่มเดียว แผงเดียวทำต่อเนื่อง (เดิม 2 ปุ่ม 2 modal + alert) */}
        <div className="flex gap-2">
          <button onClick={() => setAddLocOpen(true)} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)]">
            <Icon name="add" size="sm" /> เพิ่มฟาร์ม / โรงเรือน
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-[240px_1fr] gap-4">
        {/* Farm sidebar — จอใหญ่เท่านั้น */}
        <div className="hidden lg:block rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden lg:max-h-[78vh] lg:sticky lg:top-[var(--topbar-h)]">
          <div className="p-2 border-b border-[var(--g100)]">
            <input
              value={farmSearch}
              onChange={(e) => setFarmSearch(e.target.value)}
              placeholder="ค้นหาฟาร์ม..."
              className="w-full h-9 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px]"
            />
          </div>
          <div className="overflow-y-auto lg:max-h-[65vh] p-1.5">
            <FarmItem icon="public" label="ทุกฟาร์ม" count={totalsAll} active={currentFarm === 'ALL'} onClick={() => setCurrentFarm('ALL')} />
            {Object.keys(groups).sort().map((type) => (
              <div key={type}>
                <div className="px-2.5 pt-2.5 pb-1 flex items-center gap-1 text-[11px] font-bold text-[var(--tmuted)]"><Icon name={TYPE_ICONS[type] || 'grass'} size="xs" /> {type}</div>
                {groups[type].map((f) => (
                  <FarmItem
                    key={f}
                    icon={TYPE_ICONS[type] || 'grass'}
                    label={f}
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

        {/* Main */}
        <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
          <div className="p-4 border-b border-[var(--g100)]">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-1.5 text-[14px] font-bold text-[var(--text)]">{currentFarm === 'ALL' ? <><Icon name="public" size="sm" /> ทุกฟาร์ม</> : <><Icon name="grass" size="sm" /> {currentFarm}</>}</div>
              <div className="flex flex-wrap items-center gap-2">
                <StatPill label="ทั้งหมด" value={`${scopeTotal} ชิ้น`} icon="inventory_2" tone="blue" />
                <StatPill label="ใช้งาน" value={ok} icon="check_circle" tone="green" />
                {rep > 0 && <StatPill label="ซ่อม" value={rep} icon="build" tone="red" />}
                {shownBundles.length > 0 && <StatPill label="Bundle" value={`${shownBundles.length} ชุด`} icon="folder_open" tone="amber" />}
              </div>
            </div>
            {/* Mobile farm selector */}
            <div className="lg:hidden flex items-center gap-2 mt-3">
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
                    <option key={f} value={f}>{f} ({farmTotals[f] || 0})</option>
                  )))}
              </select>
            </div>
            <div className="relative max-w-xs mt-3">
              <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[var(--tmuted)]"><Icon name="search" size="sm" /></span>
              <input
                value={assetSearch}
                onChange={(e) => setAssetSearch(e.target.value)}
                placeholder="ค้นหาอุปกรณ์ / ชุด..."
                className="w-full h-9 pl-9 pr-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px]"
              />
            </div>
          </div>

          <div className="hidden md:block overflow-x-auto">
            <table className="w-full text-[12px]">
              <thead>
                <tr className="text-left text-[var(--tmuted)] bg-[var(--surface2)] border-b border-[var(--g100)]">
                  <th className="px-3 py-2.5 font-medium">รหัส</th>
                  <th className="px-3 py-2.5 font-medium">ชื่อ</th>
                  <th className="px-3 py-2.5 font-medium">Serial</th>
                  <th className="px-3 py-2.5 font-medium">Status</th>
                  <th className="px-3 py-2.5 font-medium">ตำแหน่ง</th>
                  <th className="px-3 py-2.5 font-medium">User</th>
                  <th className="px-3 py-2.5 font-medium text-center">จัดการ</th>
                </tr>
              </thead>
              <tbody>
                {loading && <tr><td colSpan={7} className="text-center py-10 text-[var(--tmuted)]">กำลังโหลด...</td></tr>}
                {!loading && filteredAssets.length === 0 && filteredBundles.length === 0 && (
                  <tr><td colSpan={7} className="text-center py-12 text-[var(--tmuted)]">
                    <div className="flex justify-center mb-2 text-[var(--tmuted)]"><Icon name="factory" size="2xl" /></div>ไม่พบอุปกรณ์ในไซต์งานนี้
                  </td></tr>
                )}
                {filteredBundles.map((b) => (
                  <tr key={'bundle-' + b.bundleId} className="border-b border-[var(--g100)] bg-[var(--blue-l)] cursor-pointer hover:bg-[var(--blue-b)]" onClick={() => openBundle(b.bundleId)}>
                    <td className="px-3 py-2"><span className="font-mono text-[11px]">{b.bundleId}</span></td>
                    <td className="px-3 py-2 font-bold"><Icon name="folder_open" size="xs" /> {b.bundleName}</td>
                    <td className="px-3 py-2"><span className="text-[var(--blue)]">{b.assetIds?.length || 0} ชิ้นในชุด</span></td>
                    <td className="px-3 py-2"><StatusBadge status={b.status} /></td>
                    <td className="px-3 py-2 text-[var(--tsub)]">{b.location || '-'}</td>
                    <td className="px-3 py-2">-</td>
                    <td className="px-3 py-2 text-center">
                      <button onClick={(e) => { e.stopPropagation(); openBundle(b.bundleId); }} className="flex items-center gap-1 px-2.5 py-1 rounded-lg bg-[var(--blue)] text-white text-[11px] font-semibold mx-auto"><Icon name="folder_open" size="xs" /> ดูชุด</button>
                    </td>
                  </tr>
                ))}
                {filteredAssets.map((a) => (
                  <tr key={a.serialNumber + a.assetId} className="border-b border-[var(--g100)] hover:bg-[var(--surface2)]">
                    <td className="px-3 py-2">
                      <div className="font-mono text-[11px] text-[var(--blue)]">{a.code}</div>
                      <div className="font-mono text-[10px] text-[var(--tmuted)]">{a.assetId}</div>
                    </td>
                    <td className="px-3 py-2 font-medium">{a.name}</td>
                    <td className="px-3 py-2 font-mono text-[11px] text-[var(--blue)]"><span className="block max-w-[150px] truncate" title={a.serialNumber}>{a.serialNumber}</span></td>
                    <td className="px-3 py-2"><StatusBadge status={a.status} /></td>
                    <td className="px-3 py-2">
                      <div className="min-w-0 max-w-[260px]">
                        {a.bundleId && (
                          <span className="mb-0.5 inline-flex max-w-full items-center gap-1 rounded bg-[var(--blue-l)] px-1.5 py-px text-[10px] font-medium text-[var(--blue)] border border-[var(--blue-b)]">
                            <Icon name="inventory_2" size="xs" className="flex-shrink-0" />
                            <span className="truncate">{a.bundleName || a.bundleId}</span>
                          </span>
                        )}
                        <LocationPath
                          siteName={a.siteName}
                          houseName={a.houseName}
                          houseId={a.houseId}
                          location={a.location}
                          bundleId={a.bundleId}
                        />
                      </div>
                    </td>
                    <td className="px-3 py-2 text-[var(--tsub)]">{a.user}</td>
                    <td className="px-3 py-2">
                      <div className="flex items-center justify-center gap-0.5">
                        <a href={`/trace/${encodeURIComponent(a.serialNumber)}`} target="_blank" title="ดู trace" className="h-8 w-8 flex items-center justify-center rounded-lg text-[var(--blue)] hover:bg-[var(--blue-l)]"><Icon name="description" size="sm" /></a>
                        <a href={`/qr?serial=${encodeURIComponent(a.serialNumber)}`} target="_blank" title="QR" className="h-8 w-8 flex items-center justify-center rounded-lg text-[var(--tsub)] hover:bg-[var(--surface2)]"><Icon name="qr_code" size="sm" /></a>
                        <button onClick={() => setTransfer({
                          serial: a.serialNumber,
                          current: { status: a.status, location: a.location, siteName: a.siteName, user: a.user },
                        })} title="โอนย้าย" className="h-8 w-8 flex items-center justify-center rounded-lg text-[var(--blue)] hover:bg-[var(--blue-l)]"><Icon name="local_shipping" size="sm" /></button>
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
            {!loading && filteredAssets.length === 0 && filteredBundles.length === 0 && (
              <div className="text-center py-12 text-[var(--tmuted)] text-[13px]">ไม่พบอุปกรณ์ในไซต์งานนี้</div>
            )}
            {filteredBundles.map((b) => (
              <div key={'bundle-' + b.bundleId} className="p-3.5 space-y-2.5 bg-[var(--blue-l)]" onClick={() => openBundle(b.bundleId)}>
                <div className="flex items-center gap-2.5">
                  <span className="flex items-center justify-center w-8 h-8 rounded-lg shrink-0 bg-white text-[var(--blue)] shadow-sm">
                    <Icon name="folder_open" size="sm" />
                  </span>
                  <div className="flex-1 min-w-0">
                    <div className="text-[13px] font-bold text-[var(--text)] leading-snug">{b.bundleName}</div>
                    <div className="text-[11px] font-mono text-[var(--blue)] mt-0.5">{b.bundleId}</div>
                  </div>
                  <StatusBadge status={b.status} />
                </div>
                <div className="pl-[42px] text-[12px] text-[var(--tsub)]">{b.assetIds?.length || 0} ชิ้นในชุด{b.location ? ` · ${b.location}` : ''}</div>
                <div className="pl-[42px]">
                  <button onClick={(e) => { e.stopPropagation(); openBundle(b.bundleId); }} className="flex items-center gap-1 px-3 py-1.5 rounded-lg bg-[var(--blue)] text-white text-[12px] font-semibold">
                    <Icon name="folder_open" size="xs" /> ดูชุด
                  </button>
                </div>
              </div>
            ))}
            {filteredAssets.map((a) => (
              <div key={a.serialNumber + a.assetId} className="p-3.5 space-y-2.5">
                <div className="flex items-start gap-2.5">
                  <span className="flex items-center justify-center w-8 h-8 rounded-lg shrink-0 bg-[var(--g100)] text-[var(--tsub)]">
                    <Icon name="devices" size="sm" />
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
                  {a.siteName && a.siteName !== a.location && <div className="text-[var(--tsub)]">Site: <span className="text-[var(--text)]">{a.siteName}</span></div>}
                  {a.user && <div className="text-[var(--tsub)]">ผู้ใช้: <span className="text-[var(--text)]">{a.user}</span></div>}
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
            ))}
          </div>
        </div>
      </div>

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

function FarmItem({ icon, label, count, active, onClick, bundle }) {
  return (
    <button onClick={onClick} className={`w-full flex items-center gap-2 px-2.5 py-2 rounded-lg text-left transition-colors ${active ? 'bg-[var(--blue-l)] text-[var(--blue)]' : 'hover:bg-[var(--surface2)]'}`}>
      <span className={`flex items-center justify-center w-6 h-6 rounded-md shrink-0 ${active ? 'bg-[var(--blue-l)] text-[var(--blue)]' : 'bg-[var(--g100)] text-[var(--tsub)]'}`}>
        <Icon name={icon} size="xs" />
      </span>
      <span className="flex items-center gap-1 flex-1 min-w-0 text-[12px] font-semibold truncate">{label}{bundle && <Icon name="folder_open" size="xs" />}</span>
      <span className={`text-[11px] font-bold ${active ? 'text-[var(--blue)]' : 'text-[var(--tmuted)]'}`}>{count}</span>
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
      <div className="w-full max-w-md bg-white rounded-2xl shadow-xl my-8" onClick={(e) => e.stopPropagation()}>
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