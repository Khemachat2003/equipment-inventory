// Shared status badge — Modern Pill Badges (B2 Strict Terminology)
// ข้อความที่แสดงผล: ไทยล้วน — ค่าข้อมูลจริง (อังกฤษจาก backend/Sheets) ถูกแปลงที่นี่จุดเดียว
//   • รายชิ้น: ใช้งานได้ → เขียว / สำรอง·ส่งซ่อม → ส้ม / ชำรุด,สูญ → เทา / In Bundle → ฟ้า "อยู่ในชุดติดตั้ง"
//     (ค่า "In Bundle" คือสิ่งที่ backend เขียนลงชีตตอนจัดอุปกรณ์เข้าชุด — routes/bundle.js)
//   • ชุดติดตั้ง: In Stock → ฟ้า "อยู่ในคลัง" / Deployed → เขียว "ติดตั้งแล้ว" / Maintenance → ส้ม "ซ่อมบำรุง"

const STATUS_LABELS = {
  'In Stock': 'อยู่ในคลัง',
  'Deployed': 'ติดตั้งแล้ว',
  'Maintenance': 'ซ่อมบำรุง',
  'In Bundle': 'อยู่ในชุดติดตั้ง',
};

// แปลงค่าสถานะ (ข้อมูลจริง) → ข้อความภาษาไทยที่แสดงบนหน้าจอ
export function statusLabel(status) {
  const s = String(status || '');
  if (STATUS_LABELS[s]) return STATUS_LABELS[s];
  if (s.includes('In Bundle')) return 'อยู่ในชุดติดตั้ง';
  return s;
}

export function statusTone(status) {
  const s = String(status || '');
  if (s.includes('In Bundle')) return 'blue';
  if (s === 'In Stock') return 'blue';
  if (s === 'Deployed') return 'green';
  if (s === 'Maintenance') return 'amber';
  if (s.includes('ซ่อม')) return 'amber';
  if (s.includes('สำรอง')) return 'amber';
  if (s.includes('ชำรุด') || s.includes('สูญ')) return 'gray';
  return 'green';
}

// Modern Pill — พื้นโทนสว่างสดใส + ขอบสีตรงกันคมชัด + จุดสถานะนำหน้า (อ่านชัดจากระยะ)
const TONES = {
  green: { cls: 'bg-[var(--emerald-l)] text-[var(--emerald-d)] border-[var(--emerald-b)]', dot: 'var(--emerald)' },
  amber: { cls: 'bg-[var(--amber-l)] text-[var(--amber-d)] border-[var(--amber-b)]', dot: 'var(--amber)' },
  red: { cls: 'bg-[var(--red-l)] text-[var(--red)] border-[var(--red-b)]', dot: 'var(--red)' },
  gray: { cls: 'bg-[var(--g100)] text-[var(--g600)] border-[var(--g200)]', dot: 'var(--g400)' },
  blue: { cls: 'bg-[var(--blue-l)] text-[var(--blue)] border-[var(--blue-b)]', dot: 'var(--blue)' },
};

export default function StatusBadge({ status, tone }) {
  const t = tone || statusTone(status);
  const T = TONES[t] || TONES.green;
  return (
    <span className={`inline-flex items-center gap-1.5 rounded-full border px-2.5 py-1 text-[11px] font-semibold whitespace-nowrap ${T.cls}`}>
      <span aria-hidden="true" className="h-1.5 w-1.5 rounded-full shrink-0" style={{ background: T.dot }} />
      {status ? statusLabel(status) : '-'}
    </span>
  );
}

export function historyTypeTone(type) {
  if (type === 'เบิก') return 'blue';
  if (type === 'คืน') return 'green';
  return 'gray';
}
