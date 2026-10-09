import { useEffect, useState } from 'react';
import { NavLink, Outlet, useLocation, useNavigate } from 'react-router-dom';
import axios from 'axios';
import Icon from '../ui/Icon.jsx';
import { useAuth } from '../../context/AuthContext.jsx';
import { NAV_GROUPS, getRouteMeta, MOBILE_NAV } from '../../navigation.js';
import { UI_FLAGS } from '../../uiConfig.js';

export default function Layout() {
  const { user, logout, portalHomeUrl, portalMode } = useAuth();
  const navigate = useNavigate();
  const location = useLocation();
  const meta = getRouteMeta(location.pathname);
  const isAdmin = user?.role === 'admin';
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [now, setNow] = useState(() => new Date());
  // กลุ่ม "เพิ่มเติม" พับเก็บไว้เริ่มต้น (UX simplify — คุมด้วย UI_FLAGS.simpleMenu)
  const [expanded, setExpanded] = useState(() => {
    const m = {};
    NAV_GROUPS.forEach((g) => { m[g.label] = !(g.collapsed && UI_FLAGS.simpleMenu); });
    return m;
  });

  useEffect(() => {
    const t = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(t);
  }, []);

  const timeStr = now.toLocaleTimeString('th-TH', {
    hour: '2-digit',
    minute: '2-digit',
  });

  async function handleLogout() {
    // อยู่ใน PORTAL_MODE (portal user หรือ admin ที่ล็อกอินต่อจาก portal):
    // destroy session แล้ว hard navigation กลับหน้า portal รวม — โดยไม่แตะ state
    // (เลี่ยง React rerender ระหว่างที่ state ว่าง → ไม่มีแว็บหน้า login)
    if (portalMode) {
      try { await axios.post('/api/logout'); } catch (e) {}
      const target = portalHomeUrl && /^https?:/i.test(portalHomeUrl) ? portalHomeUrl : '/';
      window.location.href = target;
      return;
    }
    await logout();
    navigate('/login');
  }

  const visibleGroups = NAV_GROUPS.filter(
    (g) => !g.adminOnly || isAdmin
  );

  return (
    <div className="flex min-h-screen bg-[var(--bg)]">
      {/* Overlay (mobile) */}
      {sidebarOpen && (
        <div
          className="fixed inset-0 z-30 bg-black/40 lg:hidden"
          onClick={() => setSidebarOpen(false)}
        />
      )}

      {/* Sidebar — drawer บนมือถือ, ปักอยู่กับที่บน lg */}
      <aside
        className={`fixed inset-y-0 left-0 z-40 w-[var(--sb-w)] flex flex-col bg-[var(--panel)] text-white shadow-[var(--sh-nav)] transition-transform duration-200 ease-in-out ${
          sidebarOpen ? 'translate-x-0' : '-translate-x-full'
        } lg:translate-x-0`}
      >
        {/* Header */}
        <div className="flex items-center gap-3 px-4 h-[var(--topbar-h)] border-b border-white/10">
          <span className="flex items-center justify-center w-8 h-8 rounded-lg bg-[var(--blue)] text-white">
            <Icon name="inventory_2" size="sm" />
          </span>
          <div className="leading-tight">
            <div className="text-sm font-bold tracking-wide">INTRANIN</div>
            <div className="text-[11px] text-white/50">Equipment Mgmt</div>
          </div>
          <button
            onClick={() => setSidebarOpen(false)}
            className="ml-auto lg:hidden flex items-center justify-center w-8 h-8 rounded-lg text-white/60 hover:bg-white/10 hover:text-white"
          >
            <Icon name="close" size="sm" />
          </button>
        </div>

        {/* Nav */}
        <nav className="flex-1 overflow-y-auto px-2 py-3 space-y-4">
          {visibleGroups
            .map((group) => ({
              ...group,
              items: group.items.filter((it) => !it.adminOnly || isAdmin),
            }))
            .filter((g) => g.items.length > 0)
            .map((group) => {
              const collapsible = group.collapsed && UI_FLAGS.simpleMenu;
              const open = expanded[group.label] ?? !collapsible;
              return (
                <div key={group.label}>
                  {collapsible ? (
                    <button
                      onClick={() => setExpanded((p) => ({ ...p, [group.label]: !open }))}
                      title={open ? 'พับเก็บ' : 'กางเมนู'}
                      className="w-full flex items-center gap-1.5 px-3 mb-1 py-1 rounded-lg text-[10px] font-semibold uppercase tracking-wider text-white/35 hover:bg-white/5 hover:text-white/60"
                    >
                      <Icon
                        name="chevron_right"
                        size="xs"
                        className={`transition-transform duration-150 ${open ? 'rotate-90' : ''}`}
                      />
                      <span>{group.label}</span>
                      {!open && (
                        <span className="ml-auto normal-case tracking-normal text-[10px] font-medium text-white/25">
                          {group.items.length} รายการ
                        </span>
                      )}
                    </button>
                  ) : (
                    <div className="px-3 mb-1 text-[10px] font-semibold uppercase tracking-wider text-white/35">
                      {group.label}
                    </div>
                  )}
                  {open && (
                    <div className="space-y-0.5">
                      {group.items.map((item) => (
                        <NavLink
                          key={item.to}
                          to={item.to}
                          end={item.end}
                          onClick={() => setSidebarOpen(false)}
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
                  )}
                </div>
              );
            })}
        </nav>

        {/* Footer */}
        <div className="px-4 py-3 border-t border-white/10 text-[11px] text-white/40">
          <div>INTRANIN EMS v1</div>
          {!isAdmin && (
            <button
              onClick={() => navigate('/login')}
              className="mt-1 text-[var(--blue-b)] hover:text-white text-[11px]"
            >
              เข้าสู่ระบบ (ผู้ดูแล)
            </button>
          )}
        </div>
      </aside>

      {/* Main */}
      <div className="min-w-0 flex-1 lg:ml-[var(--sb-w)] flex flex-col min-h-screen">
        {/* Topbar */}
        <header className="sticky top-0 z-20 h-[var(--topbar-h)] flex items-center justify-between gap-3 px-4 sm:px-6 bg-[var(--surface)]/85 backdrop-blur border-b border-[var(--g200)]">
          <div className="flex items-center gap-3 min-w-0">
            <button
              onClick={() => setSidebarOpen(true)}
              className="lg:hidden flex items-center justify-center w-9 h-9 rounded-lg text-[var(--tsub)] hover:bg-[var(--g100)]"
              title="เปิดเมนู"
            >
              <Icon name="menu" size="sm" />
            </button>
            <div className="min-w-0">
              <div className="text-[15px] font-semibold text-[var(--text)] truncate">{meta.title}</div>
              <div className="text-[11px] text-[var(--tmuted)] hidden sm:block">{meta.subtitle}</div>
            </div>
          </div>
          <div className="flex items-center gap-2 sm:gap-3 shrink-0">
            <span className="text-[11px] text-[var(--tmuted)] hidden sm:block tabular-nums">{timeStr}</span>
            <UserChip name={user?.username} role={user?.role} onLogout={handleLogout} />
          </div>
        </header>

        {/* Content — pb เผื่อที่ให้ bottom nav มือถือ */}
        <main className="min-w-0 w-full flex-1 p-4 sm:p-6 pb-24 md:pb-6">
          <Outlet />
        </main>
      </div>

      {/* Bottom nav (มือถือ) — 3 งานหลัก ไม่ต้องเปิด drawer */}
      <nav className="md:hidden fixed bottom-0 inset-x-0 z-30 flex bg-[var(--panel)] text-white shadow-[var(--sh-nav)] pb-[env(safe-area-inset-bottom)]">
        {MOBILE_NAV.map((item) => (
          <NavLink
            key={item.to}
            to={item.to}
            end={item.end}
            onClick={() => setSidebarOpen(false)}
            className={({ isActive }) =>
              `flex-1 flex flex-col items-center gap-0.5 py-2 text-[10px] font-medium transition-colors ${
                isActive ? 'text-white' : 'text-white/55'
              }`
            }
          >
            {({ isActive }) => (
              <>
                <span className={`flex items-center justify-center w-8 h-8 rounded-lg ${isActive ? 'bg-[var(--blue)]' : ''}`}>
                  <Icon name={item.icon} size="sm" />
                </span>
                {item.label}
              </>
            )}
          </NavLink>
        ))}
      </nav>

      {/* ปุ่มกลับขึ้นด้านบน — โผล่เฉพาะเวลาสกอลลงพอสมควร (ทุกอุปกรณ์) */}
      <BackToTop />
    </div>
  );
}

function BackToTop() {
  const [show, setShow] = useState(false);
  useEffect(() => {
    const onScroll = () => setShow(window.scrollY > 400);
    window.addEventListener('scroll', onScroll, { passive: true });
    return () => window.removeEventListener('scroll', onScroll);
  }, []);
  return (
    <button
      onClick={() => window.scrollTo({ top: 0, behavior: 'smooth' })}
      title="กลับขึ้นด้านบน"
      aria-label="กลับขึ้นด้านบน"
      className={`fixed right-4 bottom-20 md:bottom-6 z-30 w-10 h-10 rounded-full bg-[var(--ink)] text-white/90 shadow-[var(--sh-md)] flex items-center justify-center hover:bg-[var(--blue)] hover:text-white transition-all ${
        show ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-3 pointer-events-none'
      }`}
    >
      <Icon name="keyboard_arrow_up" size="md" />
    </button>
  );
}

function UserChip({ name, role, onLogout }) {
  const isAdmin = role === 'admin';
  return (
    <div className="flex items-center gap-2">
      <div className="flex items-center gap-2 px-2 sm:px-2.5 py-1.5 rounded-lg bg-[var(--surface2)] border border-[var(--g200)]">
        <span className="flex items-center justify-center w-6 h-6 rounded-full bg-[var(--blue)] text-white text-[10px] font-bold">
          {(name || 'U')[0]?.toUpperCase()}
        </span>
        <span className="hidden min-[400px]:inline text-[12px] font-medium text-[var(--tsub)]">{name || 'ผู้ใช้'}</span>
        {isAdmin && (
          <span className="px-1.5 py-0.5 rounded bg-[var(--amber-l)] text-[var(--amber-d)] text-[10px] font-bold">
            ADMIN
          </span>
        )}
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
