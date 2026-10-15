// CodeBadge — รหัสรอง (รหัสอุปกรณ์ / Asset ID) เป็น Badge เทารองขนาดเล็ก
// B2: ชื่ออุปกรณ์ไทยเป็นตัวหนาหลัก — รหัสย่อเป็นรอง (ภาษาเดียวกับหน้า Farm Monitor)
export default function CodeBadge({ children }) {
  return (
    <span className="inline-flex items-center whitespace-nowrap rounded-md border border-[var(--border)] bg-[var(--g100)] px-1.5 py-0.5 font-mono text-[10px] font-semibold text-[var(--tmuted)]">
      {children}
    </span>
  );
}
