import { NavLink, Outlet, useNavigate } from 'react-router-dom';
import Icon from '../ui/Icon.jsx';
import { useAuth } from '../../context/AuthContext.jsx';

const NAV_GROUPS = [
  {
    label: 'หลัก',
    items: [
      { to: '/', label: 'Dashboard', icon: 'dashboard', end: true },
      { to: '/stock', label: 'Stock', icon: 'inventory_2' },
      { to: '/asset', label: 'Asset', icon: 'devices' },
      { to: '/bundle', label: 'Bundle', icon: 'folder_open' },
    ],
  },
  {
    label: 'จัดการ',
    items: [
      { to: '/farm', label: 'ฟาร์ม', icon: 'agriculture' },
      { to: '/history', label: 'ประวัติ', icon: 'history' },
      { to: '/report', label: 'รายงาน', icon: 'bar_chart' },
      { to: '/settings', label: 'ตั้งค่า', icon: 'settings' },
    ],
  },
];

export default function Layout() {
  const { user, logout } = useAuth();
  const navigate = useNavigate();

  async function handleLogout() {
    await logout();
    navigate('/login');
  }

  return (
    <div className="flex min-h-screen bg-[var(--bg)]">
      {/* Sidebar */}
      <aside className="fixed inset-y-0 left-0 z-40 w-[var(--sb-w)] flex flex-col bg-[var(--panel)] text-white shadow-[var(--sh-nav)]">
        {/* Header */}
        <div className="flex items-center gap-3 px-4 h-[var(--topbar-h)] border-b border-white/10">
          <span className="flex items-center justify-center w-8 h-8 rounded-lg bg-[var(--blue)] text-white">
            <Icon name="inventory_2" size="sm" />
          </span>
          <div className="leading-tight">
            <div className="text-sm font-bold tracking-wide">INTRANIN</div>
            <div className="text-[11px] text-white/50">Equipment Mgmt</div>
          </div>
        </div>

        {/* Nav */}
        <nav className="flex-1 overflow-y-auto px-2 py-3 space-y-4">
          {NAV_GROUPS.map((group) => (
            <div key={group.label}>
              <div className="px-3 mb-1 text-[10px] font-semibold uppercase tracking-wider text-white/35">
                {group.label}
              </div>
              <div className="space-y-0.5">
                {group.items.map((item) => (
                  <NavLink
                    key={item.to}
                    to={item.to}
                    end={item.end}
                    className={({ isActive }) =>
                      `flex items-center gap-2.5 px-3 py-2 rounded-lg text-[13px] font-medium transition-colors ${
                        isActive
                          ? 'bg-[var(--blue)] text-white'
                          : 'text-white/70 hover:bg-white/5 hover:text-white'
                      }`
                    }
                  >
                    <Icon name={item.icon} size="sm" />
                    <span>{item.label}</span>
                  </NavLink>
                ))}
              </div>
            </div>
          ))}
        </nav>

        {/* Footer */}
        <div className="px-4 py-3 border-t border-white/10 text-[11px] text-white/40">
          INTRANIN EMS v1
        </div>
      </aside>

      {/* Main */}
      <div className="flex-1 ml-[var(--sb-w)] flex flex-col min-h-screen">
        {/* Topbar */}
        <header className="sticky top-0 z-30 h-[var(--topbar-h)] flex items-center justify-between px-6 bg-white/80 backdrop-blur border-b border-[var(--g200)]">
          <div>
            <div className="text-[15px] font-semibold text-[var(--text)]">Dashboard</div>
            <div className="text-[11px] text-[var(--tmuted)]">ภาพรวมระบบ</div>
          </div>
          <div className="flex items-center gap-3">
            <span className="text-[11px] text-[var(--tmuted)]">09:00</span>
            <UserChip name={user?.username} onLogout={handleLogout} />
          </div>
        </header>

        {/* Content */}
        <main className="flex-1 p-6">
          <Outlet />
        </main>
      </div>
    </div>
  );
}

function UserChip({ name, onLogout }) {
  return (
    <div className="flex items-center gap-2">
      <div className="flex items-center gap-2 px-2.5 py-1.5 rounded-lg bg-[var(--surface2)] border border-[var(--g200)]">
        <span className="flex items-center justify-center w-6 h-6 rounded-full bg-[var(--blue)] text-white text-[10px] font-bold">
          {(name || 'U')[0]?.toUpperCase()}
        </span>
        <span className="text-[12px] font-medium text-[var(--tsub)]">{name || 'ผู้ใช้'}</span>
      </div>
      <button
        onClick={onLogout}
        title="ออกจากระบบ"
        className="flex items-center justify-center w-8 h-8 rounded-lg text-[var(--tmuted)] hover:text-[var(--red)] hover:bg-[var(--red-l)] transition-colors"
      >
        <Icon name="logout" size="sm" />
      </button>
    </div>
  );
}
