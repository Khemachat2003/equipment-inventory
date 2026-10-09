// components/ui/Toast.jsx
// แถบแจ้งผลมุมล่างของจอ (แทน alert) — ไม่บังคับกดปิด (หายเองใน 5 วิ) และใส่ปุ่ม "ขั้นต่อไป" ได้
// การใช้งาน:
//   import { showToast, ToastHost } from '../components/ui/Toast.jsx';
//   showToast('ย้ายชุดไปฟาร์มแล้ว', { actionLabel: 'เปิดดูชุดนี้', onAction: () => ... });
//   showToast('เกิดข้อผิดพลาด', { type: 'err' });   // type: 'ok' | 'err' | 'warn'
//   ...และ render <ToastHost /> ครั้งเดียวในหน้านั้น
import { useEffect, useState } from 'react';
import Icon from './Icon.jsx';

const pushFns = new Set(); // host ล่าสุดรับ toast; ถ้า modal unmount ให้ host ของหน้ารับต่อ

export function showToast(msg, opts = {}) {
  const push = [...pushFns].at(-1);
  if (push) push({ msg, type: opts.type || 'ok', actionLabel: opts.actionLabel || '', onAction: opts.onAction || null });
}

// ใช้เฉพาะไอคอนที่มีใน font subset แล้ว: check_circle / warning
const ICON = { ok: 'check_circle', err: 'warning', warn: 'warning' };
const COLOR = { ok: 'var(--emerald)', err: '#f87171', warn: '#fbbf24' };

export function ToastHost() {
  const [item, setItem] = useState(null);

  useEffect(() => {
    const push = (next) => setItem({ ...next, id: Date.now() });
    pushFns.add(push);
    return () => { pushFns.delete(push); };
  }, []);

  // หายเองใน 5 วิ — เปลี่ยนข้อความใหม่ก็เริ่มนับใหม่
  useEffect(() => {
    if (!item) return undefined;
    const t = setTimeout(() => setItem(null), 5000);
    return () => clearTimeout(t);
  }, [item]);

  if (!item) return null;

  function act() {
    if (item.onAction) item.onAction();
    setItem(null);
  }

  return (
    <div className="fixed left-1/2 bottom-7 -translate-x-1/2 w-full px-4 sm:w-auto" style={{ zIndex: 999 }}>
      <div className="flex items-center gap-2.5 px-4 py-3 rounded-xl shadow-xl bg-[#0f172a] text-white mx-auto sm:mx-0">
        <span className="flex" style={{ color: COLOR[item.type] }}><Icon name={ICON[item.type]} size="sm" /></span>
        <span className="text-[13px] font-medium">{item.msg}</span>
        {item.actionLabel && (
          <button
            onClick={act}
            className="px-2.5 py-1 rounded-lg bg-white/15 border border-white/25 text-[12px] font-semibold hover:bg-white/25 whitespace-nowrap"
          >
            {item.actionLabel}
          </button>
        )}
        <button
          onClick={() => setItem(null)}
          title="ปิด"
          className="w-6 h-6 flex items-center justify-center rounded-lg text-white/60 hover:text-white hover:bg-white/10"
        >
          <Icon name="close" size="xs" />
        </button>
      </div>
    </div>
  );
}
