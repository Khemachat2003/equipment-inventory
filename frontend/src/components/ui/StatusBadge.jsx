// Shared status badge — จำลอง getStatusBadge จากระบบเดิม
// ใช้งานได้ → green / สำรอง → amber / ซ่อม → red / ชำรุด/สูญ → gray

// สถานะของชุดติดตั้ง (Bundle) เก็บค่าข้อมูลจริงเป็นภาษาอังกฤษตามเดิม (ไม่แตะข้อมูล/Sheets)
// → แสดงผลเป็นภาษาไทย พร้อมชื่อเดิมในวงเล็บ ให้คนที่คุ้นศัพท์เก่ายังหาเจอ
const BUNDLE_STATUS_LABELS = {
  'In Stock': 'อยู่ในคลัง (In Stock)',
  'Deployed': 'ติดตั้งที่ฟาร์มแล้ว (Deployed)',
  'Maintenance': 'ซ่อมบำรุง (Maintenance)',
};

export function statusTone(status) {
  const s = String(status || '');
  if (s.includes('ซ่อม')) return 'red';
  if (s.includes('สำรอง')) return 'amber';
  if (s.includes('ชำรุด') || s.includes('สูญ')) return 'gray';
  return 'green';
}

const TONES = {
  green: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]',
  amber: 'bg-[var(--amber-l)] text-[var(--amber-d)]',
  red: 'bg-[var(--red-l)] text-[var(--red)]',
  gray: 'bg-[var(--g200)] text-[var(--tsub)]',
  blue: 'bg-[var(--blue-l)] text-[var(--blue)]',
};

export default function StatusBadge({ status, tone }) {
  const t = tone || statusTone(status);
  return (
    <span className={`inline-block px-2 py-0.5 rounded-full text-[11px] font-semibold whitespace-nowrap ${TONES[t] || TONES.green}`}>
      {BUNDLE_STATUS_LABELS[status] || status || '-'}
    </span>
  );
}

export function historyTypeTone(type) {
  if (type === 'เบิก') return 'blue';
  if (type === 'คืน') return 'green';
  return 'gray';
}