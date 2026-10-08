import { useEffect, useState } from 'react';
import axios from 'axios';
import Icon from './ui/Icon.jsx';
import { useBusy, BusyOverlay } from './ui/Busy.jsx';
import { FarmInlineAdd, HouseInlineAdd } from './InlineFarmAdd.jsx';

// Shared Transfer Modal — จำลอง openTransferModal จากระบบเดิม
// ใช้ได้ทั้งหน้า Asset / Bundle / Farm / Scan / Trace
//
// รีดีไซน์ "เลือกปลายทางก่อน" (2026-10): แทน dropdown "ประเภทการโอนย้าย" เดิมด้วย
// ปุ่มใหญ่ 3 ปุ่ม (ไปฟาร์ม / คืนคลังกลาง / ส่งซ่อม) → map เป็น action/status เดิมทั้งหมด
// → payload /api/transfer-asset ไม่เปลี่ยนแม้แต่บิตเดียว (ข้อมูลเดิม/History/Backup ไม่กระทบ)
// หลักการ "มีแต่ไม่แสดง" ของ uiConfig.js: ประเภทการโอนย้าย 4 แบบ / สถานะ / ประเภทฟาร์ม-สัตว์ /
// หมายเหตุ ย้ายเข้า "ตัวเลือกเพิ่มเติม" (collapse) — โค้ดเดิมไม่ลบ

const ACTIONS = ['ย้ายตำแหน่งอุปกรณ์', 'ส่งซ่อมภายนอก', 'โอนย้ายผู้รับผิดชอบ', 'คืนคลังสินค้า'];
const STATUSES = ['ใช้งานได้', 'สำรอง', 'ส่งซ่อม', 'ชำรุด/สูญหาย'];

// ปลายทาง → action/status "เดิม" ของ API (แค่ map ศัพท์ให้ user นึกถึงปลายทาง ไม่ใช่ศัพท์ระบบ)
const DESTINATIONS = [
  { id: 'farm', icon: 'agriculture', title: 'ไปฟาร์ม', sub: 'ย้ายตำแหน่งอุปกรณ์', action: 'ย้ายตำแหน่งอุปกรณ์', status: 'ใช้งานได้' },
  { id: 'stock', icon: 'inventory_2', title: 'คืนคลังกลาง', sub: 'กลับเข้าคลัง Intranin', action: 'คืนคลังสินค้า', status: 'สำรอง' },
  { id: 'repair', icon: 'build', title: 'ส่งซ่อม', sub: 'ส่งซ่อมภายนอก', action: 'ส่งซ่อมภายนอก', status: 'ส่งซ่อม' },
];

const ANIMALS = {
  'สัตว์ปีก': ['ไก่เนื้อ', 'ไก่ไข่', 'เป็ด', 'ห่าน', 'อื่นๆ'],
  'สัตว์บก': ['วัว', 'ควาย', 'แพะ', 'แกะ', 'ม้า', 'อื่นๆ'],
  'สุกร': ['สุกรพ่อพันธุ์', 'สุกรแม่พันธุ์', 'ลูกสุกร', 'สุกรขุน'],
  'อื่นๆ': ['อื่นๆ'],
};

// ── จำ "ฟาร์มที่ใช้ล่าสุด" ไว้ในเบราว์เซอร์ของผู้ใช้เอง (A3) ──
// fail-safe ตามข้อกำหนด: localStorage ใช้ไม่ได้/ไม่มีค่า → ฟอร์มทำงานปกติเหมือนเดิม
const LAST_FARM_KEY = 'ems_last_farm_v1';
function readLastFarm() {
  try {
    const s = JSON.parse(localStorage.getItem(LAST_FARM_KEY) || 'null');
    return s && s.siteId ? s : null;
  } catch { return null; }
}
function saveLastFarm(farm) {
  try { localStorage.setItem(LAST_FARM_KEY, JSON.stringify(farm)); } catch { /* ignore */ }
}

export default function TransferModal({ open, onClose, onSuccess, serial, current }) {
  const [action, setAction] = useState('ย้ายตำแหน่งอุปกรณ์');
  const [status, setStatus] = useState('ใช้งานได้');
  const [farmType, setFarmType] = useState('');
  const [animalType, setAnimalType] = useState('');
  const [sites, setSites] = useState([]);
  const [siteId, setSiteId] = useState('');
  const [siteFree, setSiteFree] = useState('');
  const [houses, setHouses] = useState([]);
  const [houseId, setHouseId] = useState('');
  const [houseFree, setHouseFree] = useState('');
  const [location, setLocation] = useState('');
  const [knownLocations, setKnownLocations] = useState([]);
  const [user, setUser] = useState('');
  const [remark, setRemark] = useState('');
  // โครงใหม่: "ตัวเลือกเพิ่มเติม" พับไว้ default + ฟอร์มย่อเพิ่มฟาร์ม/โรงเรือน
  const [moreOpen, setMoreOpen] = useState(false);
  const [farmAddOpen, setFarmAddOpen] = useState(false);
  const [houseAddOpen, setHouseAddOpen] = useState(false);
  // A3: แถบบอกว่าฟาร์มถูกเติมให้จาก "ครั้งล่าสุด" (พร้อมปุ่มเลือกฟาร์มอื่น)
  const [fromMemory, setFromMemory] = useState(false);
  // A1: ผู้รับผิดชอบโชว์เป็นข้อความ "ค่าเดิม — แก้ได้" กดแก้ไขถึงเป็นช่องกรอก
  const [userEditing, setUserEditing] = useState(false);
  const busy = useBusy();

  const cur = current || {};
  const showFarm =
    !action.includes('ซ่อม') && !action.includes('คืนคลัง');
  // "ตำแหน่งย่อย" มีความหมายเฉพาะตอนอยู่ที่ฟาร์มเท่านั้น
  // (คืนคลัง = ไปที่ Stock, ส่งซ่อม = อยู่คลังแต่ซ่อม — ไม่ต้องระบุจุดย่อย)
  const isStockZone = !showFarm || siteId === 'Intranin' || !siteId;

  // ปลายทาง active ไล่จาก action เสมอ → ปุ่มปลายทางกับ advanced select ใช้ state เดียวกัน
  const activeDest = action.includes('คืนคลัง') ? 'stock' : action.includes('ซ่อม') ? 'repair' : 'farm';

  // ── แถบสรุปก่อนยืนยัน ──
  const destSiteName = siteId === 'Intranin'
    ? 'คลังกลาง (Intranin)'
    : (sites.find((s) => s.siteId === siteId)?.siteName || siteFree || '');
  const houseLabel = houseId
    ? `${houseId} ${houses.find((h) => h.houseId === houseId)?.houseName || ''}`.trim()
    : '';
  const fromText = cur.siteName
    ? `${cur.siteName}${cur.location ? ` (${cur.location})` : ''}`
    : 'ตำแหน่งเดิม';
  const summaryText = activeDest === 'stock'
    ? 'คลังกลาง (Stock)'
    : activeDest === 'repair'
      ? 'ส่งซ่อม'
      : [destSiteName || 'ฟาร์มปลายทาง', houseLabel, location.trim()].filter(Boolean).join(' › ');

  useEffect(() => {
    if (!open) return;
    setAction('ย้ายตำแหน่งอุปกรณ์');
    setStatus('ใช้งานได้'); // สถานะ auto ตามปลายทาง — แก้ได้ใน "ตัวเลือกเพิ่มเติม"
    setFarmType('');
    setAnimalType('');
    setSiteId('');
    setSiteFree('');
    setHouseId('');
    setHouseFree('');
    setLocation('');
    setUser(cur.user || '');
    setRemark('');
    setMoreOpen(false);
    setFarmAddOpen(false);
    setHouseAddOpen(false);
    setFromMemory(false);
    setUserEditing(false);
    axios.get('/api/farm-sites').then(({ data }) => {
      const list = data || [];
      setSites(list);
      // A3: เติม "ฟาร์มที่ใช้ล่าสุด" ให้ก่อน (เฉพาะถ้าฟาร์มนั้นยังมีอยู่จริง) — ไม่มีค่าก็ทำงานปกติเหมือนเดิม
      const last = readLastFarm();
      if (last && list.some((s) => s.siteId === last.siteId)) {
        setSiteId(last.siteId);
        setFromMemory(true);
      }
    }).catch(() => {});
  }, [open, cur]);

  // โหลด "ตำแหน่งย่อยในไซต์" ที่เคยใช้มาแล้ว เพื่อเป็นตัวเดาในช่องกรอก
  useEffect(() => {
    if (!open) return;
    axios.get('/api/asset-locations').then(({ data }) => setKnownLocations(data || [])).catch(() => {});
  }, [open]);

  // โหลดโรงเรือนเมื่อเลือกไซต์
  useEffect(() => {
    // เปลี่ยนไซต์แล้ว houseId ของไซต์เดิมต้องถูกล้างเสมอ
    // ไม่งั้น dropdown จะดูเหมือน "ยังไม่ได้เลือก" แต่ค่าของฟาร์มเก่ายังถูกส่งไปกับฟอร์ม
    setHouseId('');
    setHouseAddOpen(false); // เปลี่ยนฟาร์ม → ปิดฟอร์มย่อเพิ่มโรงเรือน (รหัส auto ของฟาร์มเก่าไม่ valid แล้ว)
    if (!siteId) { setHouses([]); return; }
    axios.get(`/api/farm-houses/${encodeURIComponent(siteId)}`)
      .then(({ data }) => {
        setHouses(data || []);
        const s = sites.find((x) => x.siteId === siteId);
        if (s && s.farmType) setFarmType(s.farmType);
      })
      .catch(() => setHouses([]));
  }, [siteId, sites]);

  function onActionChange(nextAction) {
    setAction(nextAction);
    if (nextAction.includes('ซ่อม')) {
      setStatus('ส่งซ่อม');
    } else if (nextAction.includes('คืนคลัง')) {
      setStatus('สำรอง');
      setSiteId('Intranin');
      setLocation('');
      setHouseId('');
    }
  }

  // ปุ่มปลายทาง 3 ปุ่ม → map เป็น action/status "เดิม" ของ API (payload ไม่เปลี่ยน)
  function chooseDest(d) {
    if (d === 'farm') {
      setAction('ย้ายตำแหน่งอุปกรณ์');
      setStatus('ใช้งานได้');
      setSiteId('');
      setHouseId('');
      setLocation('');
    } else if (d === 'stock') {
      setAction('คืนคลังสินค้า');
      setStatus('สำรอง');
      setSiteId('Intranin');
      setHouseId('');
      setLocation('');
    } else if (d === 'repair') {
      setAction('ส่งซ่อมภายนอก');
      setStatus('ส่งซ่อม');
    }
  }

  // ฟอร์มย่อเพิ่มฟาร์มสำเร็จ → refresh dropdown แล้ว "เลือกฟาร์มใหม่ให้เลย" (ไม่ต้องออกจาก modal)
  async function handleFarmAdded(site) {
    try {
      const { data } = await axios.get('/api/farm-sites');
      setSites(data || []);
    } catch { /* ignore */ }
    if (site?.siteId) setSiteId(site.siteId);
  }

  // ฟอร์มย่อเพิ่มโรงเรือนสำเร็จ → เลือกโรงเรือนใหม่ให้เลย + refresh รายการโรงเรือนของฟาร์มนี้
  async function handleHouseAdded(house) {
    if (house?.houseId) setHouseId(house.houseId);
    setHouseAddOpen(false);
    try {
      const { data } = await axios.get(`/api/farm-houses/${encodeURIComponent(siteId)}`);
      setHouses(data || []);
    } catch { /* ignore */ }
  }

  async function submit() {
    if (!serial) return alert('ไม่พบ Serial Number');
    const selectedSite = sites.find((s) => s.siteId === siteId);
    // SiteName (คอลัมน์ H) = ตัวระบุว่าอยู่ฟาร์มไหน — มาจากช่อง "ไซต์งาน / ฟาร์ม" ช่องเดียว
    const destSite = siteId === 'Intranin' ? 'Intranin' : selectedSite ? selectedSite.siteName : siteFree;
    if (showFarm && !destSite) return alert('เลือกไซต์งาน / ฟาร์มปลายทาง');

    const isReturn = action.includes('คืนคลัง');

    const body = {
      serialNumber: serial,
      action,
      status,
      // Location (คอลัมน์ G) = ตำแหน่งย่อยภายในไซต์ เช่น "ชั้น 2" / "ใกล้ประตู" — ไม่ซ้ำกับชื่อฟาร์มอีก
      location: isReturn ? 'Stock' : location.trim(),
      siteName: destSite,
      user,
      remark,
      fromLocation: `${cur.siteName || ''} (${cur.location || ''})`.trim(),
      farmType: showFarm ? farmType || '-' : '-',
      animalType: showFarm ? animalType || '-' : '-',
      houseId: isReturn ? '-' : houseId || houseFree || '-',
      houseName: isReturn ? '-' : houseId ? (houses.find((h) => h.houseId === houseId)?.houseName || houseFree) : houseFree || '-',
    };

    await busy.run('กำลังโอนย้าย...', async () => {
      try {
        const { data } = await axios.post('/api/transfer-asset', body);
        if (data.success) {
          // A3: จำฟาร์มที่เพิ่งใช้ (เฉพาะปลายทางเป็นฟาร์ม) ไว้เติมให้ครั้งหน้า
          if (showFarm && siteId && siteId !== 'Intranin') {
            saveLastFarm({ siteId, siteName: destSite });
          }
          onSuccess && onSuccess();
          onClose();
        } else {
          alert(data.error || 'เกิดข้อผิดพลาด');
        }
      } catch (e) {
        alert(e.response?.data?.error || 'เกิดข้อผิดพลาด');
      }
    });
  }

  if (!open) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center p-4 bg-black/40 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
      <div className="w-full max-w-2xl bg-white rounded-2xl shadow-xl my-8" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-[var(--g100)]">
          <div className="flex items-center gap-2">
            <span className="text-[var(--blue)]"><Icon name="sync_alt" /></span>
            <span className="text-[15px] font-bold text-[var(--text)]">โอนย้ายอุปกรณ์</span>
            <span className="px-2 py-0.5 rounded-full bg-[var(--blue-l)] text-[var(--blue)] font-mono text-[11px]">{serial}</span>
          </div>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)]">
            <Icon name="close" size="sm" />
          </button>
        </div>

        <div className="p-5 space-y-4">
          {/* ขั้น 1 — เลือกปลายทาง (ปุ่มใหญ่ 3 ปุ่ม แทน dropdown "ประเภทการโอนย้าย" เดิม) */}
          <div className="text-[12px] font-semibold text-[var(--tsub)]">อุปกรณ์นี้จะไปไหน?</div>
          <div className="grid grid-cols-3 gap-2">
            {DESTINATIONS.map((d) => (
              <button
                key={d.id}
                type="button"
                onClick={() => chooseDest(d.id)}
                className={`flex flex-col items-center gap-1 px-2 py-3 rounded-xl border-2 text-center transition-colors ${
                  activeDest === d.id
                    ? 'bg-[var(--blue-l)] border-[var(--blue)] text-[var(--blue)]'
                    : 'bg-white border-[var(--g200)] text-[var(--tsub)] hover:border-[var(--g300)] hover:bg-[var(--surface2)]'
                }`}
              >
                <Icon name={d.icon} />
                <span className="text-[13px] font-bold">{d.title}</span>
                <span className="text-[10.5px] text-[var(--tmuted)] leading-tight">{d.sub}</span>
              </button>
            ))}
          </div>
          {/* A1: สถานะ auto ตามปลายทาง — โชว์เป็นข้อความสั้นในหน้าหลัก (ไม่ซ่อนในกล่องพับ) */}
          <div className="text-[12px] text-[var(--tsub)]">
            สถานะที่จะบันทึก: <b className="text-[var(--text)]">{status}</b> <span className="text-[var(--tmuted)]">(แก้ได้ใน "รายละเอียดเพิ่ม")</span>
          </div>

          {/* ขั้น 2 — รายละเอียดปลายทาง (โชว์เฉพาะที่จำเป็นต่อปลายทางที่เลือก) */}
          {showFarm && (
            <>
              <div className="text-[12px] font-semibold text-[var(--tsub)]">เลือกฟาร์มปลายทาง</div>
              {/* A3: เติมฟาร์มล่าสุดให้ก่อน — บอกชัดพร้อมปุ่มเปลี่ยน กันย้ายผิดฟาร์ม */}
              {fromMemory && (
                <div className="flex items-center justify-between gap-2 px-3 py-2 rounded-lg bg-[var(--blue-l)] border border-[var(--blue-b)] text-[12px]">
                  <span className="text-[var(--blue-d)]">ย้ายไปที่: <b>{sites.find((s) => s.siteId === siteId)?.siteName || siteId}</b> — ฟาร์มที่ใช้ครั้งล่าสุด</span>
                  <button type="button" onClick={() => { setSiteId(''); setFromMemory(false); }} className="px-2 py-0.5 rounded-md bg-white border border-[var(--blue)] text-[11px] font-semibold text-[var(--blue)] whitespace-nowrap">เลือกฟาร์มอื่น</button>
                </div>
              )}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="ฟาร์มปลายทาง *">
                  <select value={siteId} onChange={(e) => { setSiteId(e.target.value); setFromMemory(false); }} className={inp}>
                    <option value="">— เลือกฟาร์ม —</option>
                    <option value="Intranin">บริษัท Intranin (คลังกลาง)</option>
                    {sites.filter((s) => s.siteId !== 'Intranin').map((s) => (
                      <option key={s.siteId} value={s.siteId}>{s.siteName} ({s.farmType})</option>
                    ))}
                  </select>
                  {/* เพิ่มฟาร์มใหม่ได้ทันที ไม่ต้องออกจาก modal — เพิ่มแล้ว dropdown refresh + เลือกให้เลย */}
                  <FarmInlineAdd mode="trigger" open={farmAddOpen} onOpenChange={setFarmAddOpen} />
                </Field>
                <Field label={`โรงเรือน${siteId && siteId !== 'Intranin' ? ' (ตามฟาร์มที่เลือก — ไม่บังคับ)' : ' (ไม่บังคับ)'}`}>
                  <select
                    value={houseId}
                    onChange={(e) => setHouseId(e.target.value)}
                    className={inp}
                    disabled={!siteId || siteId === 'Intranin'}
                  >
                    <option value="">— ไม่ระบุ —</option>
                    {houses.map((h) => (
                      <option key={h.houseId} value={h.houseId}>
                        {h.houseId} · {h.houseName}{h.houseType ? ` (${h.houseType})` : ''}
                      </option>
                    ))}
                  </select>
                  <HouseInlineAdd
                    mode="trigger"
                    open={houseAddOpen}
                    onOpenChange={setHouseAddOpen}
                    siteId={siteId === 'Intranin' ? '' : siteId}
                    houses={houses}
                  />
                </Field>
              </div>

              {/* ฟอร์มย่อ (แสดงเมื่อกดปุ่ม + ด้านบน) */}
              <FarmInlineAdd mode="panel" open={farmAddOpen} onOpenChange={setFarmAddOpen} onAdded={handleFarmAdded} />

              {/* ข้อความเดิมแบบ dead-end เปลี่ยนเป็นปุ่มที่กดได้ทันที */}
              {siteId && siteId !== 'Intranin' && houses.length === 0 && (
                <div className="flex items-start gap-1.5 text-[11px] text-[var(--tmuted)]">
                  <Icon name="info" size="sm" />
                  <span>ฟาร์มนี้ยังไม่มีโรงเรือนลงทะเบียน</span>
                  <button onClick={() => setHouseAddOpen(true)} className="font-semibold text-[var(--blue)] hover:underline whitespace-nowrap">
                    + เพิ่มโรงเรือนทันที
                  </button>
                </div>
              )}
              <HouseInlineAdd
                mode="panel"
                open={houseAddOpen}
                onOpenChange={setHouseAddOpen}
                siteId={siteId === 'Intranin' ? '' : siteId}
                houses={houses}
                onAdded={handleHouseAdded}
              />
            </>
          )}

          {/* ผู้รับผิดชอบ — โชว์เป็นข้อความ "ค่าเดิม — แก้ได้" ไม่ใช่ช่องว่าง (A1) */}
          <div className="flex items-center justify-between gap-2 px-3 py-2 rounded-lg bg-[var(--surface2)] border border-[var(--g100)]">
            {userEditing ? (
              <>
                <input
                  autoFocus
                  value={user}
                  onChange={(e) => setUser(e.target.value)}
                  placeholder={cur.user ? `เดิม: ${cur.user}` : 'ชื่อผู้ดูแล (ถ้ามี)'}
                  className="flex-1 h-8 px-2.5 rounded-md border border-[var(--g200)] bg-white text-[13px] focus:outline-none focus:border-[var(--blue)]"
                />
                <button type="button" onClick={() => setUserEditing(false)} className="px-2 py-1 rounded-md text-[12px] font-semibold text-[var(--blue)] whitespace-nowrap">เสร็จ</button>
              </>
            ) : (
              <>
                <span className="text-[12px] text-[var(--tsub)] min-w-0 truncate">ผู้รับผิดชอบ: <b className="text-[var(--text)]">{user || 'ไม่ระบุ'}</b>{cur.user ? ' (จากเดิม — แก้ได้)' : ''}</span>
                <button type="button" onClick={() => setUserEditing(true)} className="px-2 py-0.5 rounded-md border border-[var(--g300)] bg-white text-[11px] font-semibold text-[var(--tsub)] whitespace-nowrap">แก้ไข</button>
              </>
            )}
          </div>

          {/* แถบสรุปก่อนยืนยัน — อ่านเส้นทางเดียวจบก่อนกด */}
          <div className="px-3.5 py-3 rounded-xl bg-[var(--blue-l)] border border-[var(--blue-l)]">
            <div className="text-[11px] font-bold text-[var(--blue)] mb-0.5">สรุปก่อนยืนยัน</div>
            <div className="text-[13px] font-semibold text-[var(--text)] truncate" title={summaryText}>
              จาก {fromText} → {summaryText}
            </div>
          </div>

          {/* ตัวเลือกเพิ่มเติม — สิ่งที่ user ส่วนใหญ่ไม่ต้องแตะ (โค้ดเดิมทุกตัวเลือกยังอยู่ครบ) */}
          <div className="rounded-xl border border-[var(--g200)]">
            <button
              type="button"
              onClick={() => setMoreOpen((v) => !v)}
              className="w-full flex items-center gap-1.5 px-3 py-2.5 text-[12px] font-semibold text-[var(--tsub)] hover:bg-[var(--surface2)] rounded-t-xl"
            >
              <Icon name={moreOpen ? 'expand_less' : 'expand_more'} size="sm" />
              รายละเอียดเพิ่ม{moreOpen ? '' : ' (จุดติดตั้ง · สถานะ · ประเภทการโอนย้าย · หมายเหตุ)'}
            </button>
            {moreOpen && (
              <div className="p-3 pt-1 space-y-4 border-t border-[var(--g100)]">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <Field label="ประเภทการโอนย้าย (แก้เฉพาะกรณีพิเศษ)">
                    <select value={action} onChange={(e) => onActionChange(e.target.value)} className={inp}>
                      {ACTIONS.map((a) => <option key={a}>{a}</option>)}
                    </select>
                  </Field>
                  <Field label="สถานะใหม่">
                    <select value={status} onChange={(e) => setStatus(e.target.value)} className={inp}>
                      {STATUSES.map((s) => <option key={s}>{s}</option>)}
                    </select>
                  </Field>
                </div>
                {showFarm && (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <Field label="ประเภทฟาร์ม (เติมให้ตามฟาร์มที่เลือกแล้ว)">
                      <select value={farmType} onChange={(e) => { setFarmType(e.target.value); setAnimalType(''); }} className={inp}>
                        <option value="">— เลือก —</option>
                        {Object.keys(ANIMALS).map((t) => <option key={t}>{t}</option>)}
                      </select>
                    </Field>
                    {farmType && farmType !== 'อื่นๆ' && (
                      <Field label="ประเภทสัตว์">
                        <select value={animalType} onChange={(e) => setAnimalType(e.target.value)} className={inp}>
                          <option value="">— เลือก —</option>
                          {(ANIMALS[farmType] || []).map((a) => <option key={a}>{a}</option>)}
                        </select>
                      </Field>
                    )}
                  </div>
                )}
                {showFarm && (
                  <Field label="จุดติดตั้ง">
                    <input
                      list="knownLocations"
                      value={location}
                      onChange={(e) => setLocation(e.target.value)}
                      placeholder={isStockZone ? 'ไม่ระบุ' : 'เช่น ชั้น 2, ใกล้ประตู, โซนให้อาหาร'}
                      className={inp}
                      disabled={isStockZone}
                    />
                    <datalist id="knownLocations">
                      {knownLocations.map((l) => <option key={l} value={l} />)}
                    </datalist>
                    {!isStockZone && (
                      <div className="mt-1.5 text-[11px] text-[var(--tmuted)]">
                        พิมพ์เองได้ หรือเลือกจากจุดที่เคยใช้ในระบบ
                      </div>
                    )}
                  </Field>
                )}
                <Field label="หมายเหตุ">
                  <textarea value={remark} onChange={(e) => setRemark(e.target.value)} rows={2} placeholder="หมายเหตุ (ถ้ามี)" className={inp} />
                </Field>
              </div>
            )}
          </div>
        </div>

        <div className="flex justify-end gap-2 px-5 py-4 border-t border-[var(--g100)] bg-[var(--surface2)] rounded-b-2xl">
          <button onClick={onClose} className="px-4 py-2 rounded-lg border border-[var(--g300)] text-[13px] text-[var(--tsub)] hover:bg-white">ยกเลิก</button>
          <button
            onClick={submit}
            disabled={busy.busy}
            className="flex items-center gap-1.5 px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)] disabled:opacity-60"
          >
            <Icon name="check" size="sm" /> {busy.busy ? 'กำลังโอนย้าย...' : 'ยืนยันโอนย้าย'}
          </button>
        </div>
      </div>
      <BusyOverlay label={busy.busyLabel} />
    </div>
  );
}

const inp =
  'w-full h-9 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px] focus:outline-none focus:border-[var(--blue)]';

function Field({ label, children }) {
  return (
    <div className="space-y-1">
      <label className="block text-[12px] font-medium text-[var(--tsub)]">{label}</label>
      {children}
    </div>
  );
}

export { inp };