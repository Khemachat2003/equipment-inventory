// ActionIcon — Clean Action Group: ปุ่มไอคอนโปร่งตัวเบา + Hover Tooltip นุ่มนวล
// (แถบเม็ดเข้ม var(--ink) + เงานุ่ม — ตามภาษาเดียวกับตารางรายชิ้น Farm Monitor)
// tooltip วาง "แนวนอนซ้ายของไอคอน" ภายในแถวเสมอ → ไม่โดน clip จาก overflow-x-auto ของ scroll container
import Icon from './Icon.jsx';

export default function ActionIcon({ icon, label, href, target, onClick, tone = 'ink' }) {
  const btnCls = tone === 'blue'
    ? 'text-[var(--blue)] hover:bg-[var(--blue-l)]'
    : 'text-[var(--tsub)] hover:bg-[var(--g100)] hover:text-[var(--ink)]';
  const cls = `relative z-10 flex h-8 w-8 items-center justify-center rounded-lg transition-colors ${btnCls}`;
  return (
    <span className="group/action relative flex items-center justify-center">
      {href ? (
        <a href={href} target={target} rel={target ? 'noreferrer' : undefined} aria-label={label} className={cls}>
          <Icon name={icon} size="sm" />
        </a>
      ) : (
        <button type="button" aria-label={label} onClick={onClick} className={cls}>
          <Icon name={icon} size="sm" />
        </button>
      )}
      <span
        aria-hidden="true"
        className="pointer-events-none absolute right-[calc(100%+6px)] top-1/2 z-20 -translate-y-1/2 whitespace-nowrap rounded-lg bg-[var(--ink)] px-2 py-1 text-[10px] font-semibold text-white opacity-0 shadow-[var(--sh-md)] transition-opacity duration-150 group-hover/action:opacity-100"
      >
        {label}
      </span>
    </span>
  );
}
