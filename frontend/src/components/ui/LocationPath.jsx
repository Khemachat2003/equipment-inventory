import Icon from './Icon.jsx';
import { buildLocation } from '../../utils/location.js';

/**
 * แสดง "ตำแหน่งเต็ม" ของอุปกรณ์/ชุดอุปกรณ์
 *   บรรทัดบน = เส้นทางเต็มอ่านตามลำดับ  ฟาร์ม › โรงเรือน › จุดติดตั้ง
 *   บรรทัดล่าง = ป้ายย่อย 3 ชั้น (แสดงเฉพาะชั้นที่มีข้อมูล)
 *
 * variant="primary" (Step 2) — Primary Visual Indicator:
 *   pill สี + ไอคอนพินให้ช่างเห็นตำแหน่งฟาร์ม/โรงเรือนทันทีบนการ์ด/ตาราง
 *   ไม่ต้องกดเปิด modal — ใช้ร่วมกันทุกหน้า (Asset / Scan / การ์ดมือถือ)
 *   • อยู่ที่ฟาร์ม → pill เขียว (emerald)
 *   • อยู่คลังกลาง → pill ฟ้า (blue)
 */
export default function LocationPath({ siteName, houseName, houseId, location, bundleId, size = 'sm', showChips = false, className = '', variant = 'default' }) {
  const loc = buildLocation({ siteName, houseName, houseId, location, bundleId });

  if (variant === 'primary') {
    const isMain = size === 'md';
    const tone = loc.kind === 'stock'
      ? 'border-[var(--blue-b)] bg-[var(--blue-l)] text-[var(--blue)]'
      : 'border-[var(--emerald-b)] bg-[var(--emerald-l)] text-[var(--emerald-d)]';
    const textCls = isMain ? 'text-[13px] font-bold' : 'text-[12px] font-semibold';
    const label = loc.full || 'ยังไม่ได้ระบุตำแหน่ง';
    return (
      <div className={`min-w-0 ${className}`}>
        <span title={label} className={`inline-flex max-w-full items-center gap-1 rounded-lg border px-2 py-1 ${tone} ${textCls}`}>
          <Icon name="place" size="xs" className="flex-shrink-0" />
          <span className="truncate">{label}</span>
        </span>
        {showChips && loc.chips.length > 0 && (
          <div className="mt-1 flex flex-wrap items-center gap-x-3 gap-y-1">
            {loc.chips.map((c) => (
              <span key={c.key} className="inline-flex items-center gap-1 text-[10px] text-[var(--tmuted)]">
                <span className="rounded border border-[var(--g100)] bg-[var(--surface2)] px-1.5 py-px font-medium text-[var(--tsub)]">
                  {c.label}
                </span>
                <span className="max-w-[180px] truncate">{c.value}</span>
              </span>
            ))}
          </div>
        )}
      </div>
    );
  }

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
