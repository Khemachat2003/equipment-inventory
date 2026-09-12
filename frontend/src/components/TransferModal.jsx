import { useEffect, useState } from 'react';
import axios from 'axios';
import Icon from './ui/Icon.jsx';
import { useBusy, BusyOverlay } from './ui/Busy.jsx';

// Shared Transfer Modal — จำลอง openTransferModal จากระบบเดิม
// ใช้ได้ทั้งหน้า Asset / Bundle / Farm Monitor

const ACTIONS = ['ย้ายตำแหน่งอุปกรณ์', 'ส่งซ่อมภายนอก', 'โอนย้ายผู้รับผิดชอบ', 'คืนคลังสินค้า'];
const STATUSES = ['ใช้งานได้', 'สำรอง', 'ส่งซ่อม', 'ชำรุด/สูญหาย'];

const ANIMALS = {
  'สัตว์ปีก': ['ไก่เนื้อ', 'ไก่ไข่', 'เป็ด', 'ห่าน', 'อื่นๆ'],
  'สัตว์บก': ['วัว', 'ควาย', 'แพะ', 'แกะ', 'ม้า', 'อื่นๆ'],
  'สุกร': ['สุกรพ่อพันธุ์', 'สุกรแม่พันธุ์', 'ลูกสุกร', 'สุกรขุน'],
  'อื่นๆ': ['อื่นๆ'],
};

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
  const [user, setUser] = useState('');
  const [remark, setRemark] = useState('');
  const busy = useBusy();

  const cur = current || {};
  const showFarm =
    !action.includes('ซ่อม') && !action.includes('คืนคลัง');

  useEffect(() => {
    if (!open) return;
    setAction('ย้ายตำแหน่งอุปกรณ์');
    setStatus(cur.status || 'ใช้งานได้');
    setFarmType('');
    setAnimalType('');
    setSiteId('');
    setSiteFree('');
    setHouseId('');
    setHouseFree('');
    setLocation('');
    setUser(cur.user || '');
    setRemark('');
    axios.get('/api/farm-sites').then(({ data }) => setSites(data || [])).catch(() => {});
  }, [open, cur]);

  // โหลดโรงเรือนเมื่อเลือกไซต์
  useEffect(() => {
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
      setLocation('Stock');
    }
  }

  async function submit() {
    if (!serial) return alert('ไม่พบ Serial Number');
    const selectedSite = sites.find((s) => s.siteId === siteId);
    // เดิม: transferSite เป็นชื่อฟาร์มแบบแสดงผล (siteName) — dropdown จับคู่ siteId→siteName
    const destSite = siteId === 'Intranin' ? 'Intranin' : selectedSite ? selectedSite.siteName : siteFree;
    if (showFarm && !destSite) return alert('เลือกไซต์งาน / ฟาร์มปลายทาง');

    const body = {
      serialNumber: serial,
      action,
      status,
      location: action.includes('คืนคลัง') ? 'Stock' : location,
      siteName: destSite,
      user,
      remark,
      fromLocation: `${cur.siteName || ''} (${cur.location || ''})`.trim(),
      farmType: showFarm ? farmType || '-' : '-',
      animalType: showFarm ? animalType || '-' : '-',
      houseId: houseId || houseFree || '-',
      houseName: houseId ? (houses.find((h) => h.houseId === houseId)?.houseName || houseFree) : houseFree || '-',
    };

    await busy.run('กำลังโอนย้าย...', async () => {
      try {
        const { data } = await axios.post('/api/transfer-asset', body);
        if (data.success) {
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
          {/* Action + Status */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <Field label="ประเภทการโอนย้าย">
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

          {/* Farm section (ซ่อนเมื่อ ซ่อม/คืนคลัง) */}
          {showFarm && (
            <>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="ประเภทฟาร์ม">
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

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="ไซต์งาน / ฟาร์ม">
                  <select value={siteId} onChange={(e) => setSiteId(e.target.value)} className={inp}>
                    <option value="">— เลือกฟาร์ม —</option>
                    <option value="Intranin">บริษัท Intranin (คลังกลาง)</option>
                    {sites.filter((s) => s.siteId !== 'Intranin').map((s) => (
                      <option key={s.siteId} value={s.siteId}>{s.siteName} ({s.farmType})</option>
                    ))}
                  </select>
                </Field>
                <Field label="โรงเรือน">
                  <select value={houseId} onChange={(e) => setHouseId(e.target.value)} className={inp}>
                    <option value="">— เลือกโรงเรือน —</option>
                    {houses.map((h) => <option key={h.houseId} value={h.houseId}>{h.houseName}</option>)}
                  </select>
                </Field>
              </div>
            </>
          )}

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <Field label="Location / ตำแหน่ง">
              <select value={location} onChange={(e) => setLocation(e.target.value)} className={inp} disabled={action.includes('คืนคลัง')}>
                <option value="Stock">Stock (คลังกลาง)</option>
                {sites.filter((s) => s.siteName && s.siteName !== 'Intranin').map((s) => <option key={s.siteId} value={s.siteName}>{s.siteName}</option>)}
              </select>
            </Field>
            <Field label="ผู้รับผิดชอบ">
              <input value={user} onChange={(e) => setUser(e.target.value)} placeholder="ชื่อผู้ดูแล" className={inp} />
            </Field>
          </div>

          <Field label="หมายเหตุ">
            <textarea value={remark} onChange={(e) => setRemark(e.target.value)} rows={2} placeholder="หมายเหตุ (ถ้ามี)" className={inp} />
          </Field>
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