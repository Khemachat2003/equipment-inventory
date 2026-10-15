// components/CreateBundleModal.jsx — สร้างชุดติดตั้งใหม่ (Bundle)
// Episode 2: แยกออกจากหน้า Bundle เป็น component กลาง เพื่อใช้ซ้ำได้ทั้งหน้า Bundle
// และ Fast-Track Intake Flow ใน AddDeviceModal (Step 2 — สร้างชุดใหม่ทันทีโดยไม่ต้องออกจาก Modal)
//
// Auto-Increment Bundle Code: เปิดฟอร์ม → ระบบดูรหัสชุด "ล่าสุด" ในระบบ (เช่น BDL-018)
// แล้ว Autofill รหัสถัดไปให้ทันที (BDL-019) — ผู้ใช้ยังพิมพ์แก้ไขได้ + กันกรอกรหัสซ้ำกับชุดเดิม
// (ใช้ suggestBundleId จาก utils/bundleGroups.js — เลขสูงสุด +1 เติมศูนย์ 3 หลัก แล้วไล่ต่อถ้าชน)
import { useState } from 'react';
import Icon from './ui/Icon.jsx';
import { suggestBundleId } from '../utils/bundleGroups.js';

const inp = 'w-full h-11 sm:h-9 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[14px] sm:text-[13px] focus:outline-none focus:border-[var(--blue)]';

function Field({ label, children }) {
  return <div className="space-y-1"><label className="block text-[13px] sm:text-[12px] font-medium text-[var(--tsub)]">{label}</label>{children}</div>;
}

function Modal({ onClose, title, children }) {
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center p-3 sm:p-4 bg-black/40 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
      <div className="w-full max-w-md bg-[var(--surface)] rounded-2xl shadow-xl my-3 sm:my-8" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-4 sm:px-5 py-4 border-b border-[var(--g100)]">
          <span className="text-[15px] font-bold text-[var(--text)]">{title}</span>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)]"><Icon name="close" size="sm" /></button>
        </div>
        <div className="p-4 sm:p-5 space-y-4">{children}</div>
      </div>
    </div>
  );
}

function ModalFooter({ onClose, onSubmit, submitLabel, submitDisabled }) {
  return (
    <div className="flex flex-col-reverse sm:flex-row sm:justify-end gap-2 pt-3 border-t border-[var(--g100)]">
      <button onClick={onClose} className="w-full sm:w-auto min-h-11 sm:min-h-0 px-4 py-2 rounded-lg border border-[var(--g300)] text-[14px] sm:text-[13px] text-[var(--tsub)]">ยกเลิก</button>
      <button onClick={onSubmit} disabled={submitDisabled} className="w-full sm:w-auto min-h-11 sm:min-h-0 px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[14px] sm:text-[13px] font-semibold disabled:opacity-50">✓ {submitLabel}</button>
    </div>
  );
}

/**
 * ฟอร์มสร้างชุดติดตั้งใหม่ — รหัสชุด Autofill รันเลขต่อจากรหัสล่าสุดในระบบ (แก้ไขได้)
 * @param {string[]} existingIds รหัสชุดทั้งหมดที่มีอยู่ (ใช้เดารหัสถัดไป + กันซ้ำ)
 * @param {boolean} saving กำลังบันทึกอยู่ (parent ทำงาน async — ปิดปุ่มกันกดซ้ำ)
 * @param {(data: {bundleId:string, bundleName:string, description:string, status:string}) => void} onSubmit
 */
export default function CreateBundleModal({ onClose, onSubmit, existingIds = [], saving = false }) {
  // Episode 2 — เปิดฟอร์มครั้งแรก: ตรวจรหัสชุดล่าสุดในระบบ (เช่น BDL-018) → Autofill รหัสถัดไป (BDL-019) ให้ทันที
  const [bundleId, setBundleId] = useState(() => suggestBundleId(existingIds));
  const [bundleName, setBundleName] = useState('');
  const [description, setDescription] = useState('');
  const [status, setStatus] = useState('In Stock');
  const dup = bundleId.trim()
    ? existingIds.some((x) => (x || '').trim().toUpperCase() === bundleId.trim().toUpperCase())
    : false;
  return (
    <Modal onClose={onClose} title="เพิ่มชุดติดตั้งใหม่">
      <Field label="รหัสชุด *">
        <input value={bundleId} onChange={(e) => setBundleId(e.target.value.toUpperCase())} placeholder="เช่น BDL-001" className={inp} />
        <div className="mt-1 flex items-center gap-1 text-[11px] flex-wrap">
          {dup
            ? <span className="flex items-center gap-1 font-semibold text-[var(--red)]"><Icon name="warning" size="sm" /> รหัสนี้ถูกใช้แล้ว — ต้องไม่ซ้ำกับชุดเดิม</span>
            : <span className="flex items-center gap-1 text-[var(--tmuted)]"><Icon name="info" size="sm" /> ระบบรันเลขต่อจากรหัสชุดล่าสุด (BDL-…) ให้แล้ว — แก้ไขได้ตามต้องการ</span>}
        </div>
      </Field>
      <Field label="ชื่อชุดอุปกรณ์ *"><input value={bundleName} onChange={(e) => setBundleName(e.target.value)} placeholder="เช่น ชุดตู้ควบคุมฟาร์ม 1" className={inp} /></Field>
      <Field label="คำอธิบาย"><input value={description} onChange={(e) => setDescription(e.target.value)} className={inp} /></Field>
      <Field label="สถานะเริ่มต้น">
        <select value={status} onChange={(e) => setStatus(e.target.value)} className={inp}>
          <option value="In Stock">อยู่ในคลัง</option>
          <option value="Maintenance">ซ่อมบำรุง</option>
        </select>
      </Field>
      <ModalFooter
        onClose={onClose}
        onSubmit={() => onSubmit({ bundleId: bundleId.trim(), bundleName, description, status })}
        submitLabel="บันทึก"
        submitDisabled={!bundleId.trim() || !bundleName.trim() || dup || saving}
      />
    </Modal>
  );
}
