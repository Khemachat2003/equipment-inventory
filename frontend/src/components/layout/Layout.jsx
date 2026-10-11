import { useEffect, useRef, useState } from 'react';
import { NavLink, Outlet, useNavigate } from 'react-router-dom';
import axios from 'axios';
import Icon from '../ui/Icon.jsx';
import { useAuth } from '../../context/AuthContext.jsx';
import { TOPBAR_PILLS, DRAWER_GROUPS, MOBILE_NAV } from '../../navigation.js';

// ── B2 Concept Layout ────────────────────────────────────────────────────────
// Topbar ชุดเดียวด้านบนสุด (3 โซน): ซ้าย = โลโก้ + INTRANIN · กลาง = Nav Pills เมนูหลัก 5 รายการ
// · ขวา = Avatar ผู้ใช้ (จุดเดียวของระบบ) + ปุ่มสามขีดเปิด Drawer "เมนูเสริม" (เพิ่มเติม/ผู้ดูแลระบบ)
// - Sidebar ถาวรถูกถอดออก: เมนูหลักอยู่บน Topbar (pills) ส่วนเสริมอยู่ drawer — ไม่มี Header ซ้ำซ้อน
// - Active pill: พื้นเข้มเด่นชัด --ink (#0F172A Slate-900) + ตัวหนังสือ/ไอคอนขาวล้วน (contrast สูง)
// - Inactive pill: ตัวหนังสือเทา --g500 (#64748B) บนพื้นโปร่ง/สว่าง
// - Drawer เมนูเสริม = Clean List: ทุกกลุ่มกางพร้อมกันทันที (accordion "เพิ่มเติม" ถูกถอนออกถาวรตาม spec)
//   + ช่องไฟกว้างขึ้น + ไอคอนนำหน้าชัดเจน + hover ไฮไลท์น้ำเงินโปร่งแสงนุ่มนวล + เส้นคั่นบางๆ ใต้หัวข้อ
// - Footer drawer: ล็อกอินแล้ว = ตัวตนผู้ใช้ปัจจุบัน (Avatar เล็ก + ชื่อ + ปุ่มออกจากระบบ)
//   · admin: ไม่มีปุ่ม login ซ้ำซ้อน (/login ก็ redirect admin กลับหน้าแรกอยู่แล้ว)
//   · portal user (role:user): มีปุ่มรอง "เข้าสู่ระบบ (ผู้ดูแล)" สลับไปล็อกอิน admin ได้
//     — App.jsx อนุญาติให้ portal user เปิด /login ได้ = ทางเข้า admin login จุดเดียวของระบบ
//   · ยังไม่ล็อกอิน: ปุ่ม "เข้าสู่ระบบ (ผู้ดูแล)" เป็นปุ่มหลัก
export default function Layout() {
  const { user, logout, portalHomeUrl, portalMode } = useAuth();
  const navigate = useNavigate();
  const isAdmin = user?.role === 'admin';
  const [sidebarOpen, setSidebarOpen] = useState(false);

  // Drawer: ปิดด้วยปุ่ม Esc + ล็อก scroll พื้นหลังขณะเปิด (ทุกขนาดจอ — Modern layout)
  useEffect(() => {
    if (!sidebarOpen) return;
    const onKey = (e) => { if (e.key === 'Escape') setSidebarOpen(false); };
    window.addEventListener('keydown', onKey);
    const prevOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => {
      window.removeEventListener('keydown', onKey);
      document.body.style.overflow = prevOverflow;
    };
  }, [sidebarOpen]);

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

  const visibleGroups = DRAWER_GROUPS.filter(
    (g) => !g.adminOnly || isAdmin
  );

  return (
    <div className="flex min-h-screen bg-[var(--bg)]">
      {/* Overlay — คลุมทั้งจอทุกขนาดขณะเปิด drawer (คลิกนอกเมนู = ปิด) */}
      {sidebarOpen && (
        <div
          className="fixed inset-0 z-30 bg-black/40"
          onClick={() => setSidebarOpen(false)}
          aria-hidden="true"
        />
      )}

      {/* Drawer "เมนูเสริม" — ซ่อนตลอด เปิดด้วยปุ่มสามขีดขวาบน Topbar (เมนูหลัก 5 รายการอยู่ที่ Nav Pills กลางแถบบนแล้ว) */}
      <aside
        id="app-sidebar"
        aria-hidden={!sidebarOpen}
        inert={!sidebarOpen}
        className={`fixed inset-y-0 left-0 z-40 w-[var(--sb-w)] flex flex-col bg-[var(--panel)] text-white shadow-[var(--sh-nav)] transition-transform duration-200 ease-in-out ${
          sidebarOpen ? 'translate-x-0' : '-translate-x-full'
        }`}
      >
        {/* Header — ปุ่มปิด + ชื่อโซนเท่านั้น (ไม่ซ้ำโลโก้ INTRANIN กับ Topbar) */}
        <div className="flex items-center gap-3 px-4 h-[var(--topbar-h)] border-b border-white/10">
          <div className="text-sm font-bold text-white/90">เมนูเสริม</div>
          <div className="hidden sm:block text-[11px] text-white/50">เครื่องมือเพิ่มเติม · ผู้ดูแลระบบ</div>
          <button
            onClick={() => setSidebarOpen(false)}
            aria-label="ปิดเมนู"
            className="ml-auto flex items-center justify-center w-8 h-8 rounded-lg text-white/60 hover:bg-white/10 hover:text-white"
          >
            <Icon name="close" size="sm" />
          </button>
        </div>

        {/* Nav — Clean List: ทุกกลุ่มกางพร้อมกันทันที (ตัด accordion "เพิ่มเติม" ออกถาวรตาม spec)
            ช่องไฟกว้างขึ้น ไม่ชิดขอบซ้าย + ไอคอนนำหน้าชัดเจน + hover ไฮไลท์น้ำเงินโปร่งแสงนุ่มนวล
            + เส้นคั่นบางๆ (border-bottom จาก var(--border) จางลงให้กลมกลืนบนพื้นเข้ม) คั่นหัวข้อกับรายการ */}
        <nav className="flex-1 overflow-y-auto px-3.5 py-4 space-y-5">
          {visibleGroups
            .map((group) => ({
              ...group,
              items: group.items.filter((it) => !it.adminOnly || isAdmin),
            }))
            .filter((g) => g.items.length > 0)
            .map((group) => (
              <div key={group.label}>
                {/* หัวข้อกลุ่ม = label คั่นโซน (ไม่ใช่ปุ่มพับ) + เส้นคั่นบางๆ */}
                <div className="px-3.5 mb-2 pb-1.5 border-b border-[var(--border)]/25 text-[10px] font-semibold uppercase tracking-wider text-white/40">
                  {group.label}
                </div>
                <div className="space-y-1">
                  {group.items.map((item) => (
                    <NavLink
                      key={item.to}
                      to={item.to}
                      end={item.end}
                      onClick={() => setSidebarOpen(false)}
                      className={({ isActive }) =>
                        `group relative flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-[13px] font-medium transition-colors duration-150 ${
                          isActive
                            ? 'bg-[var(--blue)]/20 text-white before:absolute before:left-0 before:top-1/2 before:h-4 before:w-[3px] before:-translate-y-1/2 before:rounded-full before:bg-[var(--blue-b)]'
                            : 'text-white/80 hover:bg-[var(--blue)]/15 hover:text-white'
                        }`
                      }
                    >
                      {({ isActive }) => (
                        <>
                          <Icon
                            name={item.icon}
                            size="sm"
                            className={
                              isActive
                                ? 'text-[var(--blue-b)]'
                                : 'text-white/65 group-hover:text-white'
                            }
                          />
                          <span>{item.label}</span>
                        </>
                      )}
                    </NavLink>
                  ))}
                </div>
              </div>
            ))}
        </nav>

        {/* Footer — ล็อกอินแล้ว: ตัวตนผู้ใช้ปัจจุบัน (Avatar เล็ก + ชื่อ + บทบาท) + ปุ่มออกจากระบบ
            admin = จบแค่นั้น (ไม่มีปุ่ม login ซ้ำซ้อน) · portal user = + ปุ่มรอง "เข้าสู่ระบบ (ผู้ดูแล)"
            ไว้สลับไปล็อกอิน admin · ยังไม่ล็อกอิน = ปุ่มเข้าสู่ระบบเป็นปุ่มหลัก */}
        <div className="px-3.5 py-3 border-t border-white/10">
          {user ? (
            <>
              <div className="flex items-center gap-2.5 px-1 py-1.5">
                <span className="flex items-center justify-center w-8 h-8 shrink-0 rounded-full bg-[var(--blue)] text-white text-[12px] font-bold">
                  {(user.username || 'U')[0]?.toUpperCase()}
                </span>
                <div className="min-w-0 flex-1 leading-tight">
                  <div className="truncate text-[12.5px] font-semibold text-white">
                    {user.username || 'ผู้ใช้'}
                  </div>
                  <div className="text-[10.5px] text-white/45">
                    {isAdmin ? 'ผู้ดูแลระบบ (Admin)' : 'ผู้ใช้งานทั่วไป'}
                  </div>
                </div>
              </div>
              <button
                onClick={handleLogout}
                className="mt-1.5 w-full flex items-center justify-center gap-2 px-3 py-2 rounded-lg bg-white/5 border border-white/10 text-[12px] font-semibold text-white/75 transition-colors hover:bg-[var(--red)] hover:border-[var(--red)] hover:text-white"
              >
                <Icon name="logout" size="sm" />
                ออกจากระบบ (Logout)
              </button>
              {/* portal user (role:user) — ปุ่มรองสลับไปล็อกอินเป็นผู้ดูแล
                  (admin ไม่โชว์: /login จะ redirect admin กลับหน้าแรกอยู่แล้ว — ดู App.jsx) */}
              {!isAdmin && (
                <button
                  onClick={() => navigate('/login')}
                  title="สลับไปล็อกอินด้วยบัญชีผู้ดูแลระบบ (Admin)"
                  className="mt-1.5 w-full flex items-center justify-center gap-1.5 px-3 py-1.5 rounded-lg text-[11.5px] font-medium text-[var(--blue-b)] transition-colors hover:bg-[var(--blue)]/20 hover:text-white"
                >
                  <Icon name="admin_panel_settings" size="xs" />
                  เข้าสู่ระบบ (ผู้ดูแล)
                </button>
              )}
            </>
          ) : (
            <button
              onClick={() => navigate('/login')}
              className="w-full flex items-center justify-center gap-2 px-3 py-2 rounded-lg bg-white/5 border border-white/10 text-[12px] font-semibold text-[var(--blue-b)] transition-colors hover:bg-[var(--blue)]/25 hover:text-white"
            >
              <Icon name="login" size="sm" />
              เข้าสู่ระบบ (ผู้ดูแล)
            </button>
          )}
          <div className="mt-2.5 text-center text-[10px] text-white/30">INTRANIN EMS v1</div>
        </div>
      </aside>

      {/* Main — เต็มความกว้างทุกจอ (sidebar เป็น drawer overlay แล้ว ไม่ต้องกันพื้นที่ซ้าย) */}
      <div className="min-w-0 flex-1 flex flex-col min-h-screen">
        {/* Topbar ชุดเดียว (B2) — ซ้าย: โลโก้+INTRANIN · กลาง: Nav Pills เมนูหลัก · ขวา: Avatar+ปุ่มสามขีด */}
        <header className="sticky top-0 z-20 h-[var(--topbar-h)] grid grid-cols-[1fr_auto_1fr] items-center gap-2 px-3 sm:px-6 bg-[var(--surface)]/85 backdrop-blur border-b border-[var(--g200)]">
          {/* ซ้ายสุด — โลโก้ + ชื่อระบบ (จุดเดียวของทั้งระบบ) */}
          <div className="min-w-0 flex items-center gap-2 sm:gap-3 justify-self-start">
            <button
              onClick={() => navigate('/')}
              className="flex items-center gap-2 shrink-0"
              title="หน้าแรก"
            >
              <span className="flex items-center justify-center w-8 h-8 rounded-lg bg-[var(--blue)] text-white">
                <Icon name="inventory_2" size="sm" />
              </span>
              <span className="leading-tight text-left">
                <span className="block text-[13px] font-bold tracking-wide text-[var(--text)]">INTRANIN</span>
                <span className="hidden sm:block text-[10px] text-[var(--tmuted)]">Equipment Mgmt</span>
              </span>
            </button>
          </div>

          {/* ตรงกลาง — Nav Pills เมนูหลัก 5 รายการ (จอ ≥lg; จอเล็กใช้ drawer / bottom nav แทน)
              Active: พื้นเข้ม --ink + ตัวอักษร/ไอคอนขาวล้วน | Inactive: เทา --g500 บนพื้นสว่างโปร่ง */}
          <nav
            id="topbar-pills"
            aria-label="เมนูหลัก"
            className="justify-self-center hidden lg:flex items-center max-w-full overflow-x-auto no-scrollbar rounded-full bg-[var(--g100)] p-1 ring-1 ring-[var(--g200)]"
          >
            {TOPBAR_PILLS.map((item) => (
              <NavLink
                key={item.to}
                to={item.to}
                end={item.end}
                aria-label={item.label}
                title={item.label}
                className={({ isActive }) =>
                  `flex shrink-0 items-center gap-1.5 h-8 px-3 rounded-full text-[13px] font-medium transition-colors focus:outline-none focus-visible:ring-2 focus-visible:ring-[var(--blue-glow)] ${
                    isActive
                      ? 'bg-[var(--ink)] text-white shadow-[var(--sh-xs)]'
                      : 'text-[var(--g500)] hover:bg-[var(--surface)] hover:text-[var(--g600)]'
                  }`
                }
              >
                <Icon name={item.icon} size="xs" />
                <span className="hidden xl:inline">{item.label}</span>
                <span className="xl:hidden">{item.short || item.label}</span>
              </NavLink>
            ))}
          </nav>

          {/* ขวาสุด — Avatar ผู้ใช้ (จุดเดียว) + ปุ่มสามขีดเปิด drawer เมนูเสริม */}
          <div className="min-w-0 flex items-center gap-1.5 justify-self-end">
            <UserMenu name={user?.username} role={user?.role} onLogout={handleLogout} />
            <button
              onClick={() => setSidebarOpen((v) => !v)}
              className="flex items-center justify-center w-9 h-9 rounded-lg text-[var(--tsub)] hover:bg-[var(--g100)] shrink-0"
              title={sidebarOpen ? 'ปิดเมนู' : 'เปิดเมนู'}
              aria-label="เปิดเมนู"
              aria-expanded={sidebarOpen}
              aria-controls="app-sidebar"
            >
              <Icon name="menu" size="sm" />
            </button>
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
                <span className={`flex items-center justify-center w-8 h-8 rounded-lg transition-colors ${isActive ? 'bg-[var(--blue)]/25' : ''}`}>
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

// Avatar + โปรไฟล์ — จุดเดียวบน Topbar: กดแล้วกาง dropdown (ชื่อ/บทบาท + ออกจากระบบ)
function UserMenu({ name, role, onLogout }) {
  const isAdmin = role === 'admin';
  const [open, setOpen] = useState(false);
  const rootRef = useRef(null);

  // ปิด dropdown เมื่อกดนอกช่อง / กด Esc
  useEffect(() => {
    if (!open) return;
    const onDown = (e) => {
      if (rootRef.current && !rootRef.current.contains(e.target)) setOpen(false);
    };
    const onKey = (e) => { if (e.key === 'Escape') setOpen(false); };
    document.addEventListener('mousedown', onDown);
    window.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('mousedown', onDown);
      window.removeEventListener('keydown', onKey);
    };
  }, [open]);

  return (
    <div ref={rootRef} className="relative shrink-0">
      <button
        onClick={() => setOpen((v) => !v)}
        aria-label="เมนูผู้ใช้"
        aria-haspopup="menu"
        aria-expanded={open}
        title={`${name || 'ผู้ใช้'}${isAdmin ? ' (ผู้ดูแลระบบ)' : ''}`}
        className="flex items-center gap-1.5 h-9 rounded-full pl-1 pr-1 sm:pr-2 hover:bg-[var(--g100)] transition-colors"
      >
        <span className="flex items-center justify-center w-8 h-8 rounded-full bg-[var(--blue)] text-white text-[12px] font-bold">
          {(name || 'U')[0]?.toUpperCase()}
        </span>
        <span className="hidden sm:block max-w-[8rem] truncate text-[12px] font-medium text-[var(--tsub)]">
          {name || 'ผู้ใช้'}
        </span>
        {isAdmin && (
          <span className="hidden sm:block px-1.5 py-0.5 rounded bg-[var(--amber-l)] text-[var(--amber-d)] text-[10px] font-bold">
            ADMIN
          </span>
        )}
        <Icon
          name="expand_more"
          size="xs"
          className={`hidden sm:block text-[var(--tmuted)] transition-transform duration-150 ${open ? 'rotate-180' : ''}`}
        />
      </button>

      {open && (
        <div
          role="menu"
          aria-label="โปรไฟล์ผู้ใช้"
          className="absolute right-0 top-[calc(100%+6px)] z-50 w-56 rounded-xl bg-[var(--surface)] border border-[var(--g200)] shadow-[var(--sh-lg)] p-1.5"
        >
          <div className="px-2.5 py-2 min-w-0">
            <div className="text-[13px] font-bold text-[var(--text)] truncate">{name || 'ผู้ใช้'}</div>
            <div className="text-[11px] text-[var(--tmuted)]">{isAdmin ? 'ผู้ดูแลระบบ (Admin)' : 'ผู้ใช้งานทั่วไป'}</div>
          </div>
          <div className="h-px bg-[var(--g200)] my-1" aria-hidden="true" />
          <button
            role="menuitem"
            onClick={() => {
              setOpen(false);
              onLogout();
            }}
            className="w-full flex items-center gap-2 px-2.5 py-2 rounded-lg text-[13px] font-medium text-[var(--red)] hover:bg-[var(--red-l)] transition-colors"
          >
            <Icon name="logout" size="sm" />
            ออกจากระบบ
          </button>
        </div>
      )}
    </div>
  );
}
