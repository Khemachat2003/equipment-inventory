import { useEffect, useState } from 'react';
import axios from 'axios';
import { Bar } from 'react-chartjs-2';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Filler,
  Tooltip,
  Legend,
} from 'chart.js';
import Icon from '../components/ui/Icon.jsx';
import StatPill from '../components/ui/StatPill.jsx';
import { useBusy, BusyOverlay } from '../components/ui/Busy.jsx';

ChartJS.register(CategoryScale, LinearScale, BarElement, Filler, Tooltip, Legend);

const RANGES = [
  { days: 7, label: '7 วัน' },
  { days: 14, label: '14 วัน' },
  { days: 30, label: '30 วัน' },
];

const STATUS_TONES = {
  'ใช้งานได้': { tone: 'green', icon: 'check_circle' },
  'สำรอง': { tone: 'blue', icon: 'inventory_2' },
  'ส่งซ่อม': { tone: 'red', icon: 'build' },
  'ชำรุด/สูญหาย': { tone: 'red', icon: 'report' },
};

export default function Dashboard() {
  const [data, setData] = useState(null);
  const [range, setRange] = useState(30);
  const [loading, setLoading] = useState(true);
  const busy = useBusy();

  async function fetchData() {
    setLoading(true);
    try {
      const { data } = await axios.get('/api/dashboard-full');
      setData(data);
    } catch (e) {
      console.error('Dashboard load error', e);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    fetchData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  if (loading && !data) {
    return (
      <div className="flex items-center justify-center py-24 text-[var(--tmuted)]">
        กำลังโหลด...
      </div>
    );
  }

  const chart = buildChartData(data?.chartData, range);
  const totalAssets = data?.assets?.length ?? 0;
  const statusEntries = Object.entries(data?.statusCount || {})
    .map(([label, value]) => ({ label, value, ...(STATUS_TONES[label] || { tone: 'blue', icon: 'inventory_2' }) }))
    .sort((a, b) => b.value - a.value);
  const maxStatus = Math.max(1, ...statusEntries.map((s) => s.value));

  return (
    <div className="space-y-4">
      <BusyOverlay label={busy.busyLabel} />

      {/* ══ Hero — ภาพรวมตัวเลขหลัก ══ */}
      <div className="relative overflow-hidden rounded-2xl bg-[var(--ink)] text-white shadow-[var(--sh-md)]">
        {/* decorative glow */}
        <div className="absolute -top-24 -right-16 w-72 h-72 rounded-full bg-[var(--blue)]/25 blur-3xl pointer-events-none" />
        <div className="absolute -bottom-28 right-40 w-64 h-64 rounded-full bg-[var(--emerald)]/15 blur-3xl pointer-events-none" />

        <div className="relative px-6 py-6 sm:px-8 sm:py-7">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <div className="text-[11px] font-semibold uppercase tracking-widest text-white/45">ภาพรวมระบบ</div>
              <div className="flex items-baseline gap-2 mt-1">
                <span className="text-4xl sm:text-5xl font-bold tabular-nums">{data?.totalItems ?? 0}</span>
                <span className="text-[13px] text-white/60">รายการอุปกรณ์ใน Stock Master</span>
              </div>
            </div>
            <button
              onClick={() => busy.run('กำลังโหลดข้อมูล...', fetchData)}
              title="รีเฟรชข้อมูล"
              className="flex items-center gap-1.5 h-9 px-3.5 rounded-lg bg-white/10 border border-white/15 text-[13px] font-medium text-white/80 hover:bg-white/20 hover:text-white transition-colors"
            >
              <Icon name="refresh" size="sm" /> รีเฟรช
            </button>
          </div>

          {/* KPI sub-grid — ยอดวันนี้ + แยกคลัง */}
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mt-6">
            <HeroKpi icon="business_center" label="คงอยู่ Office" value={data?.totalOffice ?? 0} unit="ชิ้น" glow="rgba(27,108,168,.35)" />
            <HeroKpi icon="factory" label="อยู่ที่ Site" value={data?.totalSite ?? 0} unit="ชิ้น" glow="rgba(255,149,0,.28)" />
            <HeroKpi icon="trending_up" label="เบิกวันนี้" value={data?.todayBorrow ?? 0} unit="ครั้ง" glow="rgba(0,200,150,.28)" />
            <HeroKpi icon="trending_down" label="คืนวันนี้" value={data?.todayReturn ?? 0} unit="ครั้ง" glow="rgba(224,49,49,.28)" />
          </div>
        </div>
      </div>

      {/* ══ KPI cards — ขอบเขตอุปกรณ์/ฟาร์ม/สต็อก ══ */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <KpiCard icon="agriculture" tone="green" label="ฟาร์มทั้งหมด" value={data?.totalFarms ?? 0} desc="ตามตาราง Farm Sites" />
        <KpiCard icon="precision_manufacturing" tone="blue" label="อุปกรณ์ติดตั้งที่ฟาร์ม" value={data?.totalFarmAssets ?? 0} desc={`แยกตาม ${data?.topFarms?.length || 0} อันดับฟาร์มสูงสุด`} />
        <KpiCard icon="inventory_2" tone="amber" label="อุปกรณ์ในสต็อก" value={data?.totalStockAssets ?? 0} desc="ยังไม่ได้ติดตั้ง/อยู่ที่ Intranin" />
      </div>

      {/* ══ Main split: chart + สถานะ/อันดับฟาร์ม ══ */}
      <div className="grid grid-cols-1 xl:grid-cols-[1fr_320px] gap-4 items-start">
        {/* Chart */}
        <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-5">
          <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
            <div className="flex items-center gap-2 text-[13px] font-semibold text-[var(--text)]">
              <span className="flex items-center justify-center w-8 h-8 rounded-lg bg-[var(--blue-l)] text-[var(--blue)]">
                <Icon name="monitoring" size="sm" />
              </span>
              สถิติการเบิก–คืน
            </div>
            <div className="flex gap-1 rounded-lg bg-[var(--surface2)] border border-[var(--g100)] p-0.5">
              {RANGES.map((r) => (
                <button
                  key={r.days}
                  onClick={() => setRange(r.days)}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    range === r.days
                      ? 'bg-[var(--blue)] text-white shadow-[var(--sh-sm)]'
                      : 'text-[var(--tsub)] hover:bg-white'
                  }`}
                >
                  {r.label}
                </button>
              ))}
            </div>
          </div>
          <div className="h-[280px]">
            <Bar data={chart.data} options={chart.options} />
          </div>
        </div>

        {/* Right rail — สถานะ + อันดับฟาร์ม */}
        <div className="space-y-4">
          {/* สถานะอุปกรณ์ */}
          <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-5">
            <div className="flex items-center gap-2 text-[13px] font-semibold text-[var(--text)] mb-4">
              <span className="flex items-center justify-center w-8 h-8 rounded-lg bg-[var(--emerald-l)] text-[var(--emerald-d)]">
                <Icon name="fact_check" size="sm" />
              </span>
              สถานะอุปกรณ์
            </div>
            {statusEntries.length === 0 && <div className="text-[12px] text-[var(--tmuted)]">ไม่มีข้อมูลสถานะ</div>}
            {statusEntries.map((s) => (
              <div key={s.label} className="mb-3 last:mb-0">
                <div className="flex items-center justify-between mb-1">
                  <span className="text-[12px] text-[var(--tsub)]">{s.label}</span>
                  <span className="text-[13px] font-bold text-[var(--text)] tabular-nums">{s.value}</span>
                </div>
                <div className="h-1.5 rounded-full bg-[var(--g100)]">
                  <div
                    className="h-full rounded-full transition-all duration-700"
                    style={{ width: `${(s.value / maxStatus) * 100}%`, background: varForTone(s.tone) }}
                  />
                </div>
              </div>
            ))}
            <div className="mt-4 pt-3 border-t border-[var(--g100)] flex items-center justify-between text-[11px] text-[var(--tmuted)]">
              <span>รวมอุปกรณ์ทั้งหมด</span>
              <span className="font-bold text-[var(--text)] tabular-nums">{totalAssets} ชิ้น</span>
            </div>
          </div>

          {/* อันดับฟาร์ม */}
          <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-5">
            <div className="flex items-center gap-2 text-[13px] font-semibold text-[var(--text)] mb-4">
              <span className="flex items-center justify-center w-8 h-8 rounded-lg bg-[var(--amber-l)] text-[var(--amber-d)]">
                <Icon name="workspace_premium" size="sm" />
              </span>
              อันดับฟาร์ม (ติดตั้งสูงสุด)
            </div>
            {(!data?.topFarms || data.topFarms.length === 0) && (
              <div className="text-[12px] text-[var(--tmuted)]">ยังไม่มีฟาร์มที่ติดตั้งอุปกรณ์</div>
            )}
            {(data?.topFarms || []).map((f, i) => {
              const max = Math.max(1, ...((data?.topFarms || []).map((x) => x.count)));
              const bar = (f.count / max) * 100;
              return (
                <div key={f.name} className="flex items-center gap-3 py-2 border-b border-[var(--g100)] last:border-b-0">
                  <span className={`flex items-center justify-center w-7 h-7 rounded-lg text-[11px] font-bold shrink-0 ${i === 0 ? 'bg-[var(--amber)] text-white' : 'bg-[var(--surface2)] text-[var(--tsub)] border border-[var(--g200)]'}`}>
                    #{i + 1}
                  </span>
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center justify-between gap-2">
                      <span className="text-[12px] font-medium text-[var(--text)] truncate">{f.name}</span>
                      <span className="text-[12px] font-bold text-[var(--text)] tabular-nums shrink-0">{f.count}</span>
                    </div>
                    <div className="h-1 rounded-full bg-[var(--g100)] mt-1.5">
                      <div className="h-full rounded-full" style={{ width: `${bar}%`, background: i === 0 ? 'var(--amber)' : 'var(--blue)' }} />
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}

function varForTone(tone) {
  return {
    green: 'var(--emerald)',
    blue: 'var(--blue)',
    amber: 'var(--amber)',
    red: 'var(--red)',
  }[tone] || 'var(--blue)';
}

function HeroKpi({ icon, label, value, unit, glow }) {
  return (
    <div className="rounded-xl bg-white/[0.06] border border-white/10 p-3.5 flex items-center gap-3 backdrop-blur-sm">
      <span
        className="flex items-center justify-center w-10 h-10 rounded-xl bg-white/[0.08] border border-white/10 text-white"
        style={{ boxShadow: `inset 0 0 0 1px ${glow}, inset 0 0 20px ${glow}` }}
      >
        <Icon name={icon} size="md" />
      </span>
      <div className="min-w-0">
        <div className="text-[15px] font-bold text-white tabular-nums leading-tight">{value} <span className="text-[11px] font-medium text-white/50">{unit}</span></div>
        <div className="text-[11px] text-white/55 truncate">{label}</div>
      </div>
    </div>
  );
}

function KpiCard({ icon, tone, label, value, desc }) {
  const cls = {
    blue: 'bg-[var(--blue-l)] text-[var(--blue)]',
    green: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]',
    amber: 'bg-[var(--amber-l)] text-[var(--amber-d)]',
    red: 'bg-[var(--red-l)] text-[var(--red)]',
  }[tone] || 'bg-[var(--blue-l)] text-[var(--blue)]';
  return (
    <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-5 flex items-center gap-4">
      <span className={`flex items-center justify-center w-11 h-11 rounded-xl shrink-0 ${cls}`}>
        <Icon name={icon} size="md" />
      </span>
      <div className="min-w-0">
        <div className="text-2xl font-bold text-[var(--text)] tabular-nums leading-tight">{value}</div>
        <div className="text-[13px] font-medium text-[var(--tsub)] truncate">{label}</div>
        <div className="text-[11px] text-[var(--tmuted)] truncate">{desc}</div>
      </div>
    </div>
  );
}

function buildChartData(chartData, days) {
  const labels = chartData ? Object.keys(chartData).sort().slice(-days) : [];
  const borrowData = labels.map((d) => chartData?.[d]?.borrow || 0);
  const returnData = labels.map((d) => chartData?.[d]?.return || 0);

  return {
    data: {
      labels,
      datasets: [
        {
          label: 'เบิก',
          data: borrowData,
          backgroundColor: 'rgba(27,108,168,.85)',
          hoverBackgroundColor: '#1B6CA8',
          borderRadius: 6,
          borderSkipped: false,
          barPercentage: 0.7,
          categoryPercentage: 0.55,
          maxBarThickness: 18,
        },
        {
          label: 'คืน',
          data: returnData,
          backgroundColor: 'rgba(0,200,150,.75)',
          hoverBackgroundColor: '#00C896',
          borderRadius: 6,
          borderSkipped: false,
          barPercentage: 0.7,
          categoryPercentage: 0.55,
          maxBarThickness: 18,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { intersect: false, mode: 'index' },
      plugins: {
        legend: {
          position: 'top',
          align: 'end',
          labels: { usePointStyle: true, pointStyle: 'circle', boxWidth: 8, boxHeight: 8, color: '#475569', font: { size: 12, weight: '600' } },
        },
        tooltip: {
          backgroundColor: '#0F172A',
          titleColor: '#F8FAFC',
          bodyColor: '#E2E8F0',
          borderColor: 'rgba(255,255,255,.08)',
          borderWidth: 1,
          cornerRadius: 12,
          padding: 14,
          displayColors: true,
          usePointStyle: true,
        },
      },
      scales: {
        x: {
          grid: { display: false },
          border: { display: false },
          ticks: { color: '#94A3B8', maxTicksLimit: 12, maxRotation: 0 },
        },
        y: {
          beginAtZero: true,
          border: { display: false },
          grid: { color: 'rgba(100,116,139,.10)', drawTicks: false },
          ticks: { precision: 0, color: '#94A3B8' },
        },
      },
    },
  };
}