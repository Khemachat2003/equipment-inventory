// components/InlineFarmAdd.jsx
// ฟอร์มย่อ "เพิ่มฟาร์ม / เพิ่มโรงเรือน" — กดเพิ่มได้ทันทีโดยไม่ต้องออกจากหน้า/modal ปัจจุบัน
// ใช้ร่วมกัน: TransferModal (โอนย้ายรายชิ้น) + Bundle DeployModal (ย้ายทั้งชุด)
//
// เพิ่มสำเร็จ → เรียก onAdded(ข้อมูลที่สร้าง) ให้ caller refresh dropdown และ "เลือกค่าใหม่ให้เลย"
// error จาก backend (เช่น 409 กันเพิ่มซ้ำ) แสดงในแผงเอง ไม่เด้ง alert ทำให้ฟอร์มกรอกต่อได้
//
// การคุม state 2 แบบ:
//   - ไม่ส่ง open → คุมเปิด/ปิดเอง (trigger + panel อยู่ด้วยกัน)
//   - ส่ง open/onOpenChange → controlled (ใช้กับ mode 'trigger'/'panel' แยกกันเพื่อวาง layout)
import { useEffect, useState } from 'react';
import axios from 'axios';
import Icon from './ui/Icon.jsx';
import { FARM_TYPES, suggestSiteId, suggestHouseId } from '../utils/farmId.js';

const inp =
  'w-full h-9 px-3 rounded-lg border border-[var(--g200)] bg-white text-[13px] focus:outline-none focus:border-[var(--blue)]';

function PanelShell({ title, err, onClose, onSave, saving, saveLabel, children }) {
  return (
    <div className="rounded-xl border border-[var(--g200)] bg-[var(--surface2)] p-3 space-y-2">
      <div className="flex items-center justify-between">
        <span className="text-[12px] font-bold text-[var(--text)]">{title}</span>
        <button type="button" onClick={onClose} className="w-7 h-7 flex items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-white">
          <Icon name="close" size="sm" />
        </button>
      </div>
      {err && (
        <div className="px-2.5 py-1.5 rounded-lg bg-[var(--red-l)] text-[11.5px] font-medium text-[var(--red)]">
          {err}
        </div>
      )}
      {children}
      <div className="flex justify-end gap-2">
        <button type="button" onClick={onClose} className="px-3 h-8 rounded-lg border border-[var(--g300)] bg-white text-[12px] text-[var(--tsub)]">
          ยกเลิก
        </button>
        <button
          type="button"
          onClick={onSave}
          disabled={saving}
          className="px-3 h-8 rounded-lg bg-[var(--blue)] text-white text-[12px] font-semibold hover:bg-[var(--blue-d)] disabled:opacity-60"
        >
          {saving ? 'กำลังบันทึก...' : saveLabel}
        </button>
      </div>
    </div>
  );
}

function TriggerBtn({ onClick, label, disabled }) {
  if (disabled) {
    return (
      <span
        title="เลือกฟาร์มก่อน แล้วจะเพิ่มโรงเรือนได้"
        className="inline-flex items-center gap-1 px-2.5 h-9 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[12px] text-[var(--tmuted)] cursor-not-allowed whitespace-nowrap"
      >
        <Icon name="add" size="sm" /> {label}
      </span>
    );
  }
  return (
    <button
      type="button"
      onClick={onClick}
      className="flex items-center gap-1 px-2.5 h-9 rounded-lg border border-[var(--g300)] bg-white text-[12px] font-semibold text-[var(--blue)] hover:bg-[var(--blue-l)] whitespace-nowrap"
    >
      <Icon name="add" size="sm" /> {label}
    </button>
  );
}

/**
 * ปุ่ม "+ เพิ่มฟาร์มใหม่" + ฟอร์มย่อ (ชื่อฟาร์ม + ประเภทฟาร์ม + รหัส auto จากชื่อ)
 * mode: 'trigger' (แค่ปุ่ม) | 'panel' (แค่ฟอร์ม — แสดงเมื่อ open) | 'all' (default: ปุ่ม+ฟอร์ม)
 */
export function FarmInlineAdd({ onAdded, open, onOpenChange, mode = 'all', label = 'เพิ่มฟาร์มใหม่' }) {
  const [innerOpen, setInnerOpen] = useState(false);
  const isOpen = open === undefined ? innerOpen : open;
  const setOpen = (v) => { setInnerOpen(v); if (onOpenChange) onOpenChange(v); };

  const [name, setName] = useState('');
  const [farmType, setFarmType] = useState('สัตว์ปีก');
  const [siteId, setSiteId] = useState('');
  const [idTouched, setIdTouched] = useState(false);
  const [err, setErr] = useState('');
  const [saving, setSaving] = useState(false);

  function onName(v) {
    setName(v);
    if (!idTouched) setSiteId(suggestSiteId(v));
  }

  function reset() {
    setName(''); setFarmType('สัตว์ปีก'); setSiteId(''); setIdTouched(false); setErr(''); setSaving(false);
    setOpen(false);
  }

  async function submit() {
    if (!name.trim()) { setErr('กรุณากรอกชื่อฟาร์ม'); return; }
    if (!siteId.trim()) { setErr('กรุณากรอกรหัสฟาร์ม (ตัวอังกฤษ/ตัวเลข)'); return; }
    setSaving(true); setErr('');
    try {
      const { data } = await axios.post('/api/add-farm-site', {
        siteId: siteId.trim().toUpperCase(),
        siteName: name.trim(),
        farmType,
      });
      if (data?.success) {
        onAdded && onAdded(data.site || { siteId: siteId.trim().toUpperCase(), siteName: name.trim(), farmType });
        reset();
      } else {
        setErr(data?.error || 'เพิ่มฟาร์มไม่สำเร็จ');
      }
    } catch (e) {
      setErr(e.response?.data?.error || 'ไม่สามารถเชื่อมต่อได้');
    } finally {
      setSaving(false);
    }
  }

  const showTrigger = mode !== 'panel';
  const showPanel = mode !== 'trigger' && isOpen;

  return (
    <div className="min-w-0">
      {showTrigger && <TriggerBtn onClick={() => setOpen(!isOpen)} label={label} />}
      {showPanel && (
        <PanelShell title="เพิ่มฟาร์มใหม่" err={err} onClose={reset} onSave={submit} saving={saving} saveLabel="บันทึกฟาร์ม">
          <input value={name} onChange={(e) => onName(e.target.value)} placeholder="ชื่อฟาร์ม * เช่น ฟาร์มโคนมบางแพ" className={inp} />
          <div className="grid grid-cols-2 gap-2">
            <input
              value={siteId}
              onChange={(e) => { setSiteId(e.target.value.toUpperCase()); setIdTouched(true); }}
              placeholder="รหัสฟาร์ม * (เดาให้แล้ว)"
              className={inp}
            />
            <select value={farmType} onChange={(e) => setFarmType(e.target.value)} className={inp}>
              {FARM_TYPES.map((t) => <option key={t}>{t}</option>)}
            </select>
          </div>
        </PanelShell>
      )}
    </div>
  );
}

/**
 * ปุ่ม "+ เพิ่มโรงเรือน" + ฟอร์มย่อ (ชื่อโรงเรือน + รหัส auto เช่น SITE-H01)
 * ต้องมี siteId (ฟาร์มที่เลือกอยู่) ถึงจะเพิ่มได้ — ยังไม่เลือกฟาร์ม → ปุ่ม disabled
 * @param {Array<{houseId:string}>} houses โรงเรือนที่มีอยู่ของฟาร์มนั้น (ใช้เดารหัสถัดไป)
 */
export function HouseInlineAdd({ siteId, houses = [], onAdded, open, onOpenChange, mode = 'all', label = 'เพิ่มโรงเรือน' }) {
  const [innerOpen, setInnerOpen] = useState(false);
  const isOpen = open === undefined ? innerOpen : open;
  const setOpen = (v) => { setInnerOpen(v); if (onOpenChange) onOpenChange(v); };

  const [houseName, setHouseName] = useState('');
  const [houseId, setHouseId] = useState('');
  const [idTouched, setIdTouched] = useState(false);
  const [err, setErr] = useState('');
  const [saving, setSaving] = useState(false);

  // เปิดแผง (หรือเปลี่ยนฟาร์ม) แล้วยังไม่ได้แก้รหัสเอง → เดารหัสถัดไปให้เลย เช่น SITE-H03
  useEffect(() => {
    if (isOpen && !idTouched && siteId) {
      setHouseId(suggestHouseId(siteId, houses.map((h) => h.houseId)));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen, siteId]);

  function onName(v) {
    setHouseName(v);
    if (!idTouched) setHouseId(suggestHouseId(siteId, houses.map((h) => h.houseId)));
  }

  function reset() {
    setHouseName(''); setHouseId(''); setIdTouched(false); setErr(''); setSaving(false);
    setOpen(false);
  }

  async function submit() {
    if (!houseName.trim()) { setErr('กรุณากรอกชื่อโรงเรือน'); return; }
    if (!houseId.trim()) { setErr('กรุณากรอกรหัสโรงเรือน'); return; }
    setSaving(true); setErr('');
    try {
      const { data } = await axios.post('/api/add-farm-house', {
        houseId: houseId.trim().toUpperCase(),
        siteId,
        houseName: houseName.trim(),
      });
      if (data?.success) {
        onAdded && onAdded(data.house || { houseId: houseId.trim().toUpperCase(), siteId, houseName: houseName.trim() });
        reset();
      } else {
        setErr(data?.error || 'เพิ่มโรงเรือนไม่สำเร็จ');
      }
    } catch (e) {
      setErr(e.response?.data?.error || 'ไม่สามารถเชื่อมต่อได้');
    } finally {
      setSaving(false);
    }
  }

  const showTrigger = mode !== 'panel';
  const showPanel = mode !== 'trigger' && isOpen;

  return (
    <div className="min-w-0">
      {showTrigger && <TriggerBtn onClick={() => setOpen(!isOpen)} label={label} disabled={!siteId} />}
      {showPanel && (
        <PanelShell title="เพิ่มโรงเรือน" err={err} onClose={reset} onSave={submit} saving={saving} saveLabel="บันทึกโรงเรือน">
          <input value={houseName} onChange={(e) => onName(e.target.value)} placeholder="ชื่อโรงเรือน * เช่น โรงเรือน 1" className={inp} />
          <input
            value={houseId}
            onChange={(e) => { setHouseId(e.target.value.toUpperCase()); setIdTouched(true); }}
            placeholder="รหัสโรงเรือน (auto เช่น SITE-H01)"
            className={inp}
          />
        </PanelShell>
      )}
    </div>
  );
}