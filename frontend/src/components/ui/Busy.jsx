// components/ui/Busy.jsx
// ชุดเครื่องมือ "กำลังทำงาน" — กัน user กดซ้ำ + แจ้งว่าระบบกำลังประมวลผล
import { useState } from 'react';
import Icon from './Icon.jsx';

export function useBusy() {
  const [label, setLabel] = useState('');

  // ครอบ async action: ระหว่างรอจะแสดง overlay + ป้องกันการเรียกซ้ำ
  async function run(loadingLabel, fn) {
    if (label) return; // กำลังมีงานทำอยู่แล้ว
    setLabel(loadingLabel);
    try {
      return await fn();
    } finally {
      setLabel('');
    }
  }

  return { busy: !!label, busyLabel: label, run };
}

export function BusyOverlay({ label, subText = 'กำลังประมวลผล กรุณารอสักครู่ ห้ามกดปุ่มซ้ำ' }) {
  if (!label) return null;
  return (
    <div
      className="fixed inset-0 z-[70] flex items-center justify-center bg-[var(--bg)]/60 backdrop-blur-[2px]"
      role="status"
      aria-live="polite"
    >
      <div className="flex flex-col items-center gap-2.5 px-6 py-5 rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-lg)]">
        <span className="text-[var(--blue)]">
          <Icon name="progress_activity" size="xl" className="animate-spin" />
        </span>
        <span className="text-[14px] font-bold text-[var(--text)]">{label}</span>
        <span className="text-[12px] text-[var(--tmuted)]">{subText}</span>
      </div>
    </div>
  );
}