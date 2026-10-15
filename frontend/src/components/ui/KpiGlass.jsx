// KpiGlass — KPI Summary Card แบบ Glassmorphic (ภาษาเดียวกับ KPI Cards หน้า Farm/Bundle):
// พื้นโปร่งแสงเบลอ --surface/80 + backdrop-blur-md + ขอบบาง var(--border) คมชัด
// + แถบไล่สี tone ด้านบน 3px + เงานุ่มนวล (hover lift เบา ๆ) + ตัวเลข Hero tabular-nums
import Icon from './Icon.jsx';

const KPI_ACCENTS = {
  slate: 'from-[var(--ink3)]',
  blue: 'from-[var(--blue)]',
  green: 'from-[var(--emerald)]',
  amber: 'from-[var(--amber)]',
  red: 'from-[var(--red)]',
};

const KPI_BOX = {
  slate: 'bg-[var(--g100)] text-[var(--g700)]',
  blue: 'bg-[var(--blue-l)] text-[var(--blue)]',
  green: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]',
  amber: 'bg-[var(--amber-l)] text-[var(--amber-d)]',
  red: 'bg-[var(--red-l)] text-[var(--red)]',
};

export default function KpiGlass({ icon, tone = 'blue', label, value, unit, testId }) {
  return (
    <div
      data-testid={testId}
      className="relative flex select-none items-center gap-2.5 min-w-0 overflow-hidden rounded-xl border border-[var(--border)] bg-[var(--surface)]/80 px-3 py-2 backdrop-blur-md shadow-[var(--sh-sm)] hover:shadow-[var(--sh-md)] hover:-translate-y-px transition-all"
    >
      <span aria-hidden="true" className={`absolute inset-x-0 top-0 h-[3px] bg-gradient-to-r ${KPI_ACCENTS[tone] || KPI_ACCENTS.blue} to-transparent`} />
      <span className={`flex items-center justify-center w-8 h-8 rounded-lg shrink-0 ${KPI_BOX[tone] || KPI_BOX.blue}`}>
        <Icon name={icon} size="sm" />
      </span>
      <div className="leading-none min-w-0">
        <div className="text-[10px] font-semibold text-[var(--tmuted)] whitespace-nowrap">{label}</div>
        <div className="mt-1 text-[20px] font-extrabold text-[var(--text)] tabular-nums leading-none">
          {value} <span className="text-[11px] font-bold text-[var(--tsub)] whitespace-nowrap">{unit}</span>
        </div>
      </div>
    </div>
  );
}
