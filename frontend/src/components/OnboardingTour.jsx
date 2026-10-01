// OnboardingTour — บทแนะนำ 4 สไลด์ สำหรับผู้ใช้ใหม่ (ไม่ต้องมีคนสอน)
// แสดงครั้งแรกครั้งเดียว (จำใน localStorage) — ปิดได้ทุกเมื่อ
// เปิด/ปิดทั้งฟีเจอร์ที่ UI_FLAGS.onboardingTour (src/uiConfig.js)

import { useState } from 'react';
import Icon from './ui/Icon.jsx';
import { UI_FLAGS } from '../uiConfig.js';

const STORAGE_KEY = 'intranin-onboarded-v1';

const STEPS = [
  {
    icon: 'home',
    title: 'ยินดีต้อนรับ 👋',
    text: 'ระบบนี้ใช้ 3 งานหลักเท่านั้น — เบิก–คืนของ, ย้าย/โอนอุปกรณ์ และค้นหาว่าของอยู่ไหน เราเล่าให้ฟังสั้น ๆ 20 วินาที',
  },
  {
    icon: 'inventory_2',
    title: 'จะเบิกของ / คืนของ',
    text: 'เข้าเมนู "เบิก–คืนของ" พิมพ์ค้นหาแล้วกดปุ่ม "เบิก" ตรงแถวนั้นได้เลย จบในไม่กี่คลิก',
  },
  {
    icon: 'document_scanner',
    title: 'จะย้าย / โอนอุปกรณ์',
    text: 'เข้าเมนู "ย้าย/โอนอุปกรณ์" สแกนบาร์โค้ด (หรือพิมพ์รหัส) แล้วเลือกปลายทาง — ตำแหน่งเดิมขึ้นให้เองแล้ว',
  },
  {
    icon: 'help',
    title: 'ต้องการความช่วยเหลือ',
    text: 'กดปุ่ม ❓ ที่แถบบนได้ทุกหน้า เพื่อเปิดคู่มือ "วิธีใช้งาน" — ไม่ต้องรอใครมาสอน',
  },
];

export default function OnboardingTour() {
  // -1 = ปิด (เคยดูแล้ว หรือผู้ใช้กดปิด)
  const [step, setStep] = useState(() => {
    try {
      return localStorage.getItem(STORAGE_KEY) ? -1 : 0;
    } catch {
      return -1; // private mode ไม่มี storage → ไม่กวนผู้ใช้
    }
  });

  if (!UI_FLAGS.onboardingTour || step < 0) return null;

  const s = STEPS[step];
  const last = step === STEPS.length - 1;

  function finish() {
    try { localStorage.setItem(STORAGE_KEY, '1'); } catch { /* ignore */ }
    setStep(-1);
  }

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-4 bg-black/45 backdrop-blur-sm">
      <div className="w-full max-w-md rounded-2xl bg-white shadow-[var(--sh-lg)] p-6">
        <div className="flex items-center gap-3 mb-3">
          <span className="flex items-center justify-center w-12 h-12 rounded-xl bg-[var(--blue-l)] text-[var(--blue)] shrink-0">
            <Icon name={s.icon} size="lg" />
          </span>
          <div className="min-w-0">
            <div className="text-[11px] font-semibold uppercase tracking-widest text-[var(--tmuted)]">
              รู้จักระบบ {step + 1}/{STEPS.length}
            </div>
            <div className="text-[16px] font-bold text-[var(--text)] truncate">{s.title}</div>
          </div>
          <button onClick={finish} title="ปิด" className="ml-auto w-8 h-8 rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)] shrink-0">
            <Icon name="close" size="sm" />
          </button>
        </div>
        <p className="text-[13px] text-[var(--tsub)] leading-relaxed mb-4">{s.text}</p>
        <div className="flex items-center gap-2">
          <div className="flex gap-1.5 mr-auto">
            {STEPS.map((_, i) => (
              <span
                key={i}
                className={`h-1.5 rounded-full transition-all ${i === step ? 'w-5 bg-[var(--blue)]' : 'w-1.5 bg-[var(--g300)]'}`}
              />
            ))}
          </div>
          <button onClick={finish} className="h-9 px-3 rounded-lg text-[13px] text-[var(--tsub)] hover:bg-[var(--surface2)]">
            ข้าม
          </button>
          <button
            onClick={() => (last ? finish() : setStep(step + 1))}
            className="h-9 px-4 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)]"
          >
            {last ? 'เริ่มใช้งาน' : 'ถัดไป'}
          </button>
        </div>
      </div>
    </div>
  );
}