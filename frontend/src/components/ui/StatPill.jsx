// StatPill — mini KPI pill สำหรับแถวสถิติ (รวม/ใช้งาน/ซ่อม/...)
// ดีไซน์เดียวกับ StatCard ของ Dashboard/Bundle (icon box + label + ตัวเลข) แต่ compact กว่า
// ใช้ shared tokens: --blue-l/--blue, --emerald-l/--emerald-d, --amber-l/--amber-d, --red-l/--red

import Icon from './Icon.jsx';

const TONES = {
  blue: {
    box: 'bg-[var(--blue-l)] text-[var(--blue)] border-[var(--blue-l)]',
    dot: 'var(--blue)',
  },
  green: {
    box: 'bg-[var(--emerald-l)] text-[var(--emerald-d)] border-[var(--emerald-l)]',
    dot: 'var(--emerald)',
  },
  amber: {
    box: 'bg-[var(--amber-l)] text-[var(--amber-d)] border-[var(--amber-l)]',
    dot: 'var(--amber)',
  },
  red: {
    box: 'bg-[var(--red-l)] text-[var(--red)] border-[var(--red-l)]',
    dot: 'var(--red)',
  },
};

export default function StatPill({ label, value, icon, tone = 'blue', dot = false }) {
  const t = TONES[tone] || TONES.blue;
  return (
    <div className="flex items-center gap-2.5 px-3 py-2 rounded-xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] select-none">
      {icon && (
        <span className={`flex items-center justify-center w-8 h-8 rounded-lg border ${t.box}`}>
          <Icon name={icon} size="sm" />
        </span>
      )}
      {dot && !icon && <span className="w-2 h-2 rounded-full shrink-0" style={{ background: t.dot }} />}
      <div className="leading-none">
        <div className="text-[10px] font-semibold uppercase tracking-wider text-[var(--tmuted)]">{label}</div>
        <div className="text-[17px] font-bold text-[var(--text)] mt-0.5 tabular-nums">{value}</div>
      </div>
    </div>
  );
}