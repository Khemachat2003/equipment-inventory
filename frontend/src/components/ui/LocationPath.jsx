import { buildLocation } from '../../utils/location.js';

/**
 * แสดง "ตำแหน่งเต็ม" ของอุปกรณ์/ชุดอุปกรณ์
 *   บรรทัดบน = เส้นทางเต็มอ่านตามลำดับ  ฟาร์ม › โรงเรือน › จุดติดตั้ง
 *   บรรทัดล่าง = ป้ายย่อย 3 ชั้น (แสดงเฉพาะชั้นที่มีข้อมูล)
 */
export default function LocationPath({ siteName, houseName, houseId, location, size = 'sm', showChips = false, className = '' }) {
  const loc = buildLocation({ siteName, houseName, houseId, location });

  const isMain = size === 'md';
  const textCls = isMain
    ? 'text-[14px] font-semibold text-[var(--text)]'
    : 'text-[12px] text-[var(--tsub)]';
  const empty = loc.kind === 'stock' && !loc.chips.length;

  return (
    <div className={`min-w-0 ${className}`}>
      <div className={`${textCls} truncate`} title={loc.full}>
        {loc.full || '—'}
      </div>
      {showChips && loc.chips.length > 0 && (
        <div className="flex flex-wrap items-center gap-x-3 gap-y-1 mt-1">
          {loc.chips.map((c) => (
            <span key={c.key} className="inline-flex items-center gap-1 text-[10px] text-[var(--tmuted)]">
              <span className="rounded bg-[var(--surface2)] px-1.5 py-px font-medium text-[var(--tsub)] border border-[var(--g100)]">
                {c.label}
              </span>
              <span className="truncate max-w-[180px]">{c.value}</span>
            </span>
          ))}
        </div>
      )}
      {empty && !showChips && (
        <div className="text-[11px] text-[var(--tmuted)] mt-0.5">ยังไม่ได้ระบุตำแหน่ง</div>
      )}
    </div>
  );
}
