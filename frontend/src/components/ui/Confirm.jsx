// components/ui/Confirm.jsx
// กล่องยืนยันเล็ก (แทน window.confirm) — ใช้เหมือน Toast: showConfirm() คืน Promise<boolean>
// การใช้งาน:
//   import { showConfirm, ConfirmHost } from '../components/ui/Confirm.jsx';
//   if (!(await showConfirm({ title: 'นำ SN-X ออกจากชุด?', confirmLabel: 'นำออกจากชุด', danger: true }))) return;
//   ...และ render <ConfirmHost /> ครั้งเดียวในหน้านั้น (เทียบเท่า ToastHost)
//   options: { title, message?, confirmLabel?, cancelLabel?, danger? } — danger = ปุ่มยืนยันสีแดง (งานถอน/ลบ)
import { useEffect, useRef, useState } from 'react';
import Icon from './Icon.jsx';

let pushFn = null; // ผูกกับ ConfirmHost ที่กำลัง mount อยู่ (หน้าหนึ่ง render ครั้งเดียวพอ)

// คืน Promise<boolean> — true = กดยืนยัน · false = ยกเลิก (ปุ่มยกเลิก / Esc / คลิกพื้นหลัง)
export function showConfirm(opts = {}) {
  return new Promise((resolve) => {
    if (!pushFn) { resolve(false); return; } // host ยังไม่ mount → ถือเป็นยกเลิก (กัน promise ค้าง)
    pushFn(opts, resolve);
  });
}

// ใช้เฉพาะไอคอนที่มีใน font subset แล้ว: warning
export function ConfirmHost() {
  const [item, setItem] = useState(null); // { opts }
  const pendingRef = useRef(null); // resolve ของ promise ที่รอคำตอบอยู่

  // ปิดกล่องพร้อมตอบ promise — เรียกจากปุ่ม/พื้นหลัง
  function answer(value) {
    const pending = pendingRef.current;
    pendingRef.current = null;
    setItem(null);
    if (pending) pending.resolve(value);
  }

  useEffect(() => {
    pushFn = (opts, resolve) => {
      if (pendingRef.current) pendingRef.current.resolve(false); // มีกล่องค้างอยู่ → ปิดตัวเก่าเป็นยกเลิก
      pendingRef.current = { resolve };
      setItem({ opts });
    };
    return () => {
      // หน้าถูก unmount ระหว่างรอคำตอบ → ปิด promise เป็น false กันค้าง
      if (pendingRef.current) { pendingRef.current.resolve(false); pendingRef.current = null; }
      pushFn = null;
    };
  }, []);

  // Esc = ยกเลิก (logic เดียวกับ answer แต่เขียน inline เพื่อไม่ติด exhaustive-deps)
  useEffect(() => {
    if (!item) return undefined;
    const onKey = (e) => {
      if (e.key !== 'Escape') return;
      const pending = pendingRef.current;
      pendingRef.current = null;
      setItem(null);
      if (pending) pending.resolve(false);
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [item]);

  if (!item) return null;
  const { opts } = item;
  const danger = !!opts.danger;

  return (
    <div
      className="fixed inset-0 z-[70] flex items-center justify-center p-4 bg-black/40"
      role="dialog"
      aria-modal="true"
      onClick={(e) => { if (e.target === e.currentTarget) answer(false); }}
    >
      <div className="w-full max-w-sm rounded-2xl bg-white shadow-[var(--sh-lg)] p-5">
        <div className="flex items-start gap-3">
          <span className={`mt-0.5 shrink-0 ${danger ? 'text-[var(--red)]' : 'text-[var(--blue)]'}`}>
            <Icon name="warning" />
          </span>
          <div className="min-w-0">
            <h3 className="text-[15px] font-bold text-[var(--text)] leading-snug">{opts.title || 'ยืนยัน'}</h3>
            {opts.message && <p className="mt-1.5 text-[13px] text-[var(--tsub)] leading-relaxed">{opts.message}</p>}
          </div>
        </div>
        <div className="mt-4 flex justify-end gap-2">
          <button
            onClick={() => answer(false)}
            className="px-3 py-2 rounded-lg border border-[var(--g300)] text-[13px] text-[var(--tsub)] hover:bg-[var(--surface2)]"
          >
            {opts.cancelLabel || 'ยกเลิก'}
          </button>
          <button
            autoFocus
            onClick={() => answer(true)}
            className={`px-4 py-2 rounded-lg text-white text-[13px] font-semibold ${danger ? 'bg-[var(--red)] hover:opacity-90' : 'bg-[var(--blue)] hover:bg-[var(--blue-d)]'}`}
          >
            {opts.confirmLabel || 'ยืนยัน'}
          </button>
        </div>
      </div>
    </div>
  );
}