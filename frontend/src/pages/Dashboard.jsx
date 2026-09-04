import { useEffect, useState } from 'react';
import axios from 'axios';
import { Line } from 'react-chartjs-2';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  Filler,
  Tooltip,
  Legend,
} from 'chart.js';
import Icon from '../components/ui/Icon.jsx';

ChartJS.register(CategoryScale, LinearScale, PointElement, LineElement, Filler, Tooltip, Legend);

const RANGES = [
  { days: 7, label: '7 วัน' },
  { days: 14, label: '14 วัน' },
  { days: 30, label: '30 วัน' },
];

export default function Dashboard() {
  const [data, setData] = useState(null);
  const [range, setRange] = useState(30);
  const [loading, setLoading] = useState(true);

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

  if (loading && !data) return <div className="text-[var(--tmuted)]">กำลังโหลด...</div>;

  const chart = buildChartData(data?.chartData, range);

  return (
    <div className="space-y-4">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <div className="text-lg font-bold text-[var(--text)]">Dashboard</div>
          <div className="text-[12px] text-[var(--tmuted)]">ภาพรวมระบบ ณ วันนี้</div>
        </div>
        <button
          onClick={fetchData}
          className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[13px] font-medium text-[var(--tsub)] hover:bg-[var(--surface2)] transition-colors"
        >
          <Icon name="refresh" size="sm" /> รีเฟรช
        </button>
      </div>

      {/* Hero: primary metric + KPI + today */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        {/* Primary */}
        <div className="rounded-2xl bg-[var(--ink)] text-white p-6 shadow-[var(--sh-md)]">
          <div className="text-[11px] font-semibold tracking-widest uppercase text-white/40">ภาพรวมอุปกรณ์</div>
          <div className="text-5xl font-bold mt-2">{data?.totalItems ?? 0}</div>
          <div className="text-[12px] text-white/50 mt-1">ประเภทอุปกรณ์ทั้งหมดในระบบ</div>
        </div>

        {/* KPI */}
        <div className="grid grid-cols-2 gap-4">
          <StatCard icon="business_center" iconColor="text-[var(--blue)]" label="คงอยู่ Office (ชิ้น)" value={data?.totalOffice ?? 0} />
          <StatCard icon="factory" iconColor="text-[var(--amber)]" label="อยู่ที่ Site (ชิ้น)" value={data?.totalSite ?? 0} />
        </div>

        {/* Today chips */}
        <div className="grid grid-cols-2 gap-4">
          <ChipCard label="เบิกวันนี้" value={data?.todayBorrow ?? 0} icon="trending_up" tone="up" />
          <ChipCard label="คืนวันนี้" value={data?.todayReturn ?? 0} icon="trending_down" tone="dn" />
        </div>
      </div>

      {/* Main split: chart + inventory rail */}
      <div className="grid grid-cols-1 xl:grid-cols-[1fr_300px] gap-4">
        {/* Chart */}
        <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-5">
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-2 text-[13px] font-semibold text-[var(--text)]">
              <Icon name="monitoring" size="sm" /> สถิติการเบิก–คืน
            </div>
            <div className="flex gap-1">
              {RANGES.map((r) => (
                <button
                  key={r.days}
                  onClick={() => setRange(r.days)}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    range === r.days
                      ? 'bg-[var(--blue)] text-white'
                      : 'text-[var(--tsub)] hover:bg-[var(--surface2)]'
                  }`}
                >
                  {r.label}
                </button>
              ))}
            </div>
          </div>
          <div className="h-[280px]">
            <Line data={chart.data} options={chart.options} />
          </div>
        </div>

        {/* Inventory rail */}
        <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-5">
          <div className="flex items-center gap-2 text-[13px] font-semibold text-[var(--text)] mb-4">
            <Icon name="account_tree" size="sm" /> ขอบเขตระบบ
          </div>
          <RailRow icon="agriculture" label="ฟาร์มทั้งหมด" value={data?.totalFarms ?? 0} />
          <RailRow icon="precision_manufacturing" tone="teal" label="อุปกรณ์ในฟาร์ม" value={data?.totalFarmAssets ?? 0} />
          <RailRow icon="inventory_2" tone="violet" label="อุปกรณ์ในสต็อก" value={data?.totalStockAssets ?? 0} />
          <div className="mt-4 rounded-xl bg-[var(--amber-l)] p-3.5 flex items-center gap-3">
            <span className="flex items-center justify-center w-9 h-9 rounded-lg bg-[var(--amber)] text-white">
              <Icon name="workspace_premium" size="sm" />
            </span>
            <div>
              <div className="text-[10px] font-semibold uppercase tracking-wide text-[var(--amber)]">ฟาร์มติดตั้งมากสุด</div>
              <div className="text-[13px] font-bold text-[var(--text)]">{data?.topFarms?.[0] ? `${data.topFarms[0].name} (${data.topFarms[0].count})` : '—'}</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function StatCard({ icon, iconColor, label, value }) {
  return (
    <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-4 flex flex-col justify-between">
      <span className={`mb-2`}>
        <Icon name={icon} size="md" color={iconColor} />
      </span>
      <div>
        <div className="text-2xl font-bold text-[var(--text)]">{value}</div>
        <div className="text-[11px] text-[var(--tmuted)]">{label}</div>
      </div>
    </div>
  );
}

function ChipCard({ label, value, icon, tone }) {
  const toneStyle =
    tone === 'up'
      ? 'bg-[var(--emerald-l)] text-[var(--emerald-d)]'
      : 'bg-[var(--red-l)] text-[var(--red)]';
  return (
    <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-4">
      <div className={`flex items-center justify-center w-9 h-9 rounded-lg mb-2 ${toneStyle}`}>
        <Icon name={icon} size="sm" />
      </div>
      <div className="text-2xl font-bold text-[var(--text)]">{value}</div>
      <div className="text-[11px] text-[var(--tmuted)]">{label}</div>
    </div>
  );
}

function RailRow({ icon, label, value, tone }) {
  const toneMap = {
    teal: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]',
    violet: 'bg-[var(--purple-l)] text-[var(--purple)]',
    undefined: 'bg-[var(--blue-l)] text-[var(--blue)]',
  };
  const cls = toneMap[tone] ?? toneMap[undefined];
  return (
    <div className="flex items-center gap-3 py-2.5 border-b border-[var(--g100)] last:border-b-0">
      <span className={`flex items-center justify-center w-9 h-9 rounded-lg ${cls}`}>
        <Icon name={icon} size="sm" />
      </span>
      <div>
        <div className="text-[11px] text-[var(--tmuted)]">{label}</div>
        <div className="text-[15px] font-bold text-[var(--text)]">{value}</div>
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
          borderColor: '#1B6CA8',
          backgroundColor: 'rgba(27,108,168,.14)',
          fill: true,
          tension: 0.45,
          pointRadius: 3,
          pointBackgroundColor: '#fff',
          pointBorderColor: '#1B6CA8',
          pointBorderWidth: 2,
          borderWidth: 2.5,
          borderCapStyle: 'round',
          borderJoinStyle: 'round',
        },
        {
          label: 'คืน',
          data: returnData,
          borderColor: '#00C896',
          backgroundColor: 'rgba(0,200,150,.12)',
          fill: true,
          tension: 0.45,
          pointRadius: 3,
          pointBackgroundColor: '#fff',
          pointBorderColor: '#00C896',
          pointBorderWidth: 2,
          borderWidth: 2.5,
          borderCapStyle: 'round',
          borderJoinStyle: 'round',
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
