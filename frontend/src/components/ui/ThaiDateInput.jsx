// ThaiDateInput — ช่องวันที่รูปแบบไทย "วว/ดด/ปปปป" (B2 Strict Terminology)
// ปัญหาเดิม: <input type="date"> ของ browser โชว์ mm/dd/yyyy ตาม locale เครื่อง (อ่านไม่รู้เรื่องสำหรับหน้างานไทย)
// → แทนด้วยช่องข้อความรูปแบบไทย + ปุ่มปฏิทิน (เปิด date picker ของ browser ผ่าน showPicker)
// พิมพ์ได้ทั้ง ค.ศ. (2025) และ พ.ศ. (2568) — ค่าที่ส่งออก (onChange) เป็น ISO "YYYY-MM-DD" เสมอ
// (ตรงกับ receivedDateKey / matchesInboundFilters ที่ใช้กรองช่วงวันรับเข้า)

import { useEffect, useRef, useState } from 'react';
import Icon from './Icon.jsx';

const NON_DIGITS = /\D/g;

// ISO 'YYYY-MM-DD' → ข้อความไทย 'วว/ดด/ปปปป' (ปี พ.ศ.)
export function thaiDisplay(iso) {
  if (!iso) return '';
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso));
  if (!m) return '';
  return `${m[3]}/${m[2]}/${parseInt(m[1], 10) + 543}`;
}

// ข้อความไทย 'วว/ดด/ปปปป' → ISO 'YYYY-MM-DD'
//   ''    = ยังพิมพ์ไม่ครบ (ยังไม่เปลี่ยนตัวกรอง)
//   null  = รูปแบบไม่ถูกต้อง
//   รับปีได้ทั้ง ค.ศ. (1900-2200) และ พ.ศ. (≥2400 → ลบ 543)
export function thaiToIso(text) {
  const digits = String(text || '').replace(NON_DIGITS, '');
  if (!digits) return '';
  if (digits.length < 8) return '';
  const dd = digits.slice(0, 2);
  const mm = digits.slice(2, 4);
  let year = parseInt(digits.slice(4, 8), 10);
  if (mm < '01' || mm > '12') return null;
  if (dd < '01' || dd > '31') return null;
  if (year >= 2400) year -= 543;
  if (year < 1900 || year > 2200) return null;
  return `${year}-${mm}-${dd}`;
}

// จัดขีด "/" ขณะพิมพ์: '01102026' → '01/10/2026'
function maskThaiDigits(digits) {
  const d = digits.slice(0, 8);
  const parts = [d.slice(0, 2)];
  if (d.length > 2) parts.push(d.slice(2, 4));
  if (d.length > 4) parts.push(d.slice(4, 8));
  return parts.filter(Boolean).join('/');
}

export default function ThaiDateInput({ label, ariaLabel, value = '', onChange, className = '' }) {
  const [text, setText] = useState(() => thaiDisplay(value));
  const lastSynced = useRef(value);
  const pickerRef = useRef(null);

  // sync เมื่อค่าถูกเปลี่ยนจากภายนอก (เช่น เลือก PO แล้วเคลียร์ตัวกรอง) — ไม่แตะขณะผู้ใช้กำลังพิมพ์
  useEffect(() => {
    if (lastSynced.current === value) return;
    lastSynced.current = value;
    setText(thaiDisplay(value));
  }, [value]);

  function handleChange(e) {
    const digits = e.target.value.replace(NON_DIGITS, '').slice(0, 8);
    const masked = maskThaiDigits(digits);
    setText(masked);
    if (!digits.length) { lastSynced.current = ''; onChange(''); return; }
    const iso = thaiToIso(masked);
    if (iso) { lastSynced.current = iso; onChange(iso); }
  }

  // พิมพ์ค้างไม่ครบ/ไม่ถูกต้อง → คืนข้อความตามค่าที่ใช้กรองจริงตอนออกจากช่อง
  function handleBlur() {
    const iso = thaiToIso(text);
    if (!iso) setText(thaiDisplay(lastSynced.current));
  }

  function openPicker() {
    const el = pickerRef.current;
    if (!el) return;
    try {
      if (typeof el.showPicker === 'function') el.showPicker();
      else el.focus();
    } catch {
      el.focus();
    }
  }

  return (
    <label
      className={`relative flex items-center gap-1.5 h-10 rounded-xl border border-[var(--g200)] bg-[var(--surface2)] px-2.5 transition-colors focus-within:border-[var(--blue)] focus-within:ring-2 focus-within:ring-[var(--blue-glow)] ${className}`}
    >
      <span className="shrink-0 text-[11px] font-semibold text-[var(--tsub)] whitespace-nowrap">{label}</span>
      <input
        type="text"
        inputMode="numeric"
        value={text}
        onChange={handleChange}
        onBlur={handleBlur}
        placeholder="วว/ดด/ปปปป"
        aria-label={ariaLabel || label}
        className="h-full min-w-0 w-24 flex-1 bg-transparent text-[12.5px] font-medium tabular-nums text-[var(--ink)] placeholder:text-[var(--tmuted)] placeholder:font-normal focus:outline-none"
      />
      {value && (
        <button
          type="button"
          onClick={() => { lastSynced.current = ''; setText(''); onChange(''); }}
          aria-label={`ล้างวันที่${label ? ' ' + label : ''}`}
          title="ล้างวันที่"
          className="shrink-0 flex items-center justify-center w-6 h-6 rounded-md text-[var(--tmuted)] hover:bg-[var(--g100)] hover:text-[var(--ink)]"
        >
          <Icon name="close" size="xs" />
        </button>
      )}
      <button
        type="button"
        onClick={openPicker}
        aria-label={`เปิดปฏิทินเลือกวันที่${label ? ' ' + label : ''}`}
        title="เลือกจากปฏิทิน"
        className="shrink-0 flex items-center justify-center w-7 h-7 rounded-lg text-[var(--blue)] hover:bg-[var(--blue-l)]"
      >
        <Icon name="event" size="xs" />
      </button>
      {/* ปฏิทิน native (ซ่อนจากการมองเห็น) — เปิดผ่านปุ่มด้านบน (showPicker) ค่าที่ได้เป็น ISO ตรง ๆ */}
      <input
        ref={pickerRef}
        type="date"
        value={value || ''}
        onChange={(e) => {
          const iso = e.target.value;
          lastSynced.current = iso;
          setText(thaiDisplay(iso));
          onChange(iso);
        }}
        tabIndex={-1}
        aria-hidden="true"
        className="absolute left-1/2 bottom-0 h-px w-px opacity-0 pointer-events-none"
      />
    </label>
  );
}
