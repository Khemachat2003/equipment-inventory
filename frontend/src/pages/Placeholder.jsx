// Placeholder page shown while migrating each section of the original app.
export default function Placeholder({ title, note }) {
  return (
    <div className="flex flex-col items-center justify-center py-24 text-center">
      <div className="text-[var(--tmuted)] text-sm mb-2">อยู่ระหว่างย้ายหน้า</div>
      <div className="text-lg font-semibold text-[var(--text)]">{title || 'ใช้งานของเดิมผ่านหน้า HTML เดิม'}</div>
      {note && <div className="text-[12px] text-[var(--tmuted)] mt-1">{note}</div>}
    </div>
  );
}
