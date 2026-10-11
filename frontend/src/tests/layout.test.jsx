// layout.test.jsx — B2 Concept: Topbar ชุดเดียว (โลโก้ | Nav Pills กลาง | Avatar + สามขีด)
// + Active pill พื้นเข้มตัวขาว (contrast สูง) + Drawer "เมนูเสริม" + Avatar จุดเดียว
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { render, screen, fireEvent, within, act } from '@testing-library/react';
import { MemoryRouter } from 'react-router-dom';
import Layout from '../components/layout/Layout.jsx';
import { useAuth } from '../context/AuthContext.jsx';

// Mock useAuth — Layout ต้องการ user/logout/portalMode (ไม่ยุ่ง network จริง)
// logout ประกาศใน factory (vi.mock hoisted) แล้วดึงกลับผ่าน useAuth() เพื่อ assert ได้
// authState ผ่าน vi.hoisted — ให้แต่ละเทสต์สลับ user ได้ (admin = default / portal user)
const authState = vi.hoisted(() => ({ user: { username: 'tester', role: 'admin' } }));
vi.mock('../context/AuthContext.jsx', () => {
  const logout = vi.fn();
  return {
    useAuth: () => ({ ...authState, logout, portalHomeUrl: '', portalMode: false }),
  };
});

// ทุกเทสต์เริ่มจาก user admin (default) — เทสต์ portal user ตั้งทับเฉพาะตัวเอง
beforeEach(() => {
  authState.user = { username: 'tester', role: 'admin' };
});

const renderLayout = (route = '/') =>
  render(
    <MemoryRouter initialEntries={[route]}>
      <Layout />
    </MemoryRouter>
  );

const asideOf = (container) => container.querySelector('aside#app-sidebar');
const pillsOf = (container) => container.querySelector('nav#topbar-pills');
const pillOf = (container, href) =>
  [...pillsOf(container).querySelectorAll('a')].find((a) => a.getAttribute('href') === href);

describe('Layout — B2 Topbar (ชุดเดียว ไม่ซ้ำซ้อน)', () => {
  it('เหลือ header เดียว + โลโก้ INTRANIN จุดเดียวใน Layout (ไม่มี Header ตัวเก่าซ้ำ)', () => {
    const { container } = renderLayout('/');
    expect(container.querySelectorAll('header')).toHaveLength(1);
    const logos = [...container.querySelectorAll('span')].filter((s) => s.textContent === 'INTRANIN');
    expect(logos).toHaveLength(1);
  });

  it('Nav Pills กลาง Topbar (≥lg) มีเมนูหลักครบ 5 รายการตามลำดับ', () => {
    const { container } = renderLayout('/');
    const nav = pillsOf(container);
    expect(nav).toBeTruthy();
    expect(nav.getAttribute('aria-label')).toBe('เมนูหลัก');
    expect(nav.className).toContain('hidden lg:flex'); // จอ <lg ใช้ drawer / bottom nav แทน
    const hrefs = [...nav.querySelectorAll('a')].map((a) => a.getAttribute('href'));
    expect(hrefs).toEqual(['/', '/farm', '/bundle', '/stock', '/scan']);
  });

  it('โครง Topbar 3 โซน: Avatar อยู่ก่อนปุ่มสามขีด (ขวาสุด)', () => {
    const { container } = renderLayout('/');
    const header = container.querySelector('header');
    expect(header.className).toContain('grid-cols-[1fr_auto_1fr]');
    const avatar = screen.getByRole('button', { name: 'เมนูผู้ใช้' });
    const hamburger = screen.getByRole('button', { name: 'เปิดเมนู' });
    expect(
      avatar.compareDocumentPosition(hamburger) & Node.DOCUMENT_POSITION_FOLLOWING
    ).toBeTruthy();
  });
});

describe('Layout — Active/Inactive Nav Pill (contrast สูง)', () => {
  it('Active pill: พื้นเข้ม --ink (#0F172A) + ตัวหนังสือ/ไอคอนขาวล้วน', () => {
    const { container } = renderLayout('/stock');
    const active = pillOf(container, '/stock');
    expect(active.className).toContain('bg-[var(--ink)]');
    expect(active.className).toContain('text-white');
    // ต้องไม่ใช่ tint โปร่งแสงแบบเมนู drawer อีกต่อไป
    expect(active.className).not.toContain('bg-[var(--blue)]/20');
    // ไอคอนในแท็บ active แสดงจริง (รับสีขาวจาก pill)
    expect(active.querySelector('.msi')).toBeTruthy();
  });

  it('Inactive pill: ตัวหนังสือเทา --g500 (#64748B) บนพื้นโปร่ง ไม่โดนพื้นเข้ม', () => {
    const { container } = renderLayout('/stock');
    const idle = pillOf(container, '/scan');
    expect(idle.className).toContain('text-[var(--g500)]');
    expect(idle.className).not.toContain('bg-[var(--ink)]');
    expect(idle.className).not.toContain('text-white');
  });
});

describe('Layout — Avatar จุดเดียว + Footer drawer (ตัวตนผู้ใช้)', () => {
  it('Avatar dropdown ยังเป็นจุดเดียวบน Topbar (ปุ่มเมนูผู้ใช้มีปุ่มเดียว)', () => {
    renderLayout('/');
    expect(screen.getAllByRole('button', { name: 'เมนูผู้ใช้' })).toHaveLength(1);
  });

  it('ล็อกอินเป็น admin: footer แสดงตัวตนผู้ใช้ (Avatar เล็ก + ชื่อ) + ปุ่มออกจากระบบ — ไม่มีปุ่มเข้าสู่ระบบซ้ำซ้อน (admin เปิด /login ไม่ได้อยู่แล้ว)', () => {
    const { container } = renderLayout('/');
    // เปิด drawer ก่อน — ตอนปิด aside เป็น aria-hidden + inert ทำให้ query ตาม role ไม่เห็น
    fireEvent.click(screen.getByRole('button', { name: 'เปิดเมนู' }));
    const aside = asideOf(container);
    // footer โชว์ชื่อผู้ใช้ปัจจุบันแทนปุ่ม "เข้าสู่ระบบ (ผู้ดูแล)" เดิม
    expect(within(aside).getByText('tester')).toBeTruthy();
    expect(within(aside).getByRole('button', { name: /ออกจากระบบ/ })).toBeTruthy();
    expect(within(aside).queryByText(/เข้าสู่ระบบ/)).toBeNull();
  });

  it('ล็อกอินเป็น portal user (role:user): ยังต้องมีปุ่ม "เข้าสู่ระบบ (ผู้ดูแล)" คู่กับตัวตน + ออกจากระบบ (ทางเข้า admin login จุดเดียวของระบบ)', () => {
    authState.user = { username: 'portal_user', role: 'user', portal: true };
    const { container } = renderLayout('/');
    fireEvent.click(screen.getByRole('button', { name: 'เปิดเมนู' }));
    const aside = asideOf(container);
    expect(within(aside).getByText('portal_user')).toBeTruthy();
    expect(within(aside).getByRole('button', { name: /ออกจากระบบ/ })).toBeTruthy();
    expect(within(aside).getByRole('button', { name: /เข้าสู่ระบบ \(ผู้ดูแล\)/ })).toBeTruthy();
  });

  it('กดปุ่มออกจากระบบใน drawer → เรียก logout', async () => {
    renderLayout('/');
    fireEvent.click(screen.getByRole('button', { name: 'เปิดเมนู' }));
    // handleLogout เป็น async (await logout → navigate) — ครอบด้วย act กัน act(...) warning
    await act(async () => {
      fireEvent.click(screen.getByRole('button', { name: /ออกจากระบบ/ }));
    });
    const { logout } = useAuth();
    expect(logout).toHaveBeenCalledTimes(1);
  });

  it('กด Avatar → dropdown โปรไฟล์ (ชื่อ + ออกจากระบบ); กดออกจากระบบเรียก logout', async () => {
    renderLayout('/');
    fireEvent.click(screen.getByRole('button', { name: 'เมนูผู้ใช้' }));
    const menu = screen.getByRole('menu');
    expect(within(menu).getByText('tester')).toBeTruthy();
    // handleLogout เป็น async (await logout → navigate) — ครอบด้วย act กัน act(...) warning
    await act(async () => {
      fireEvent.click(within(menu).getByRole('menuitem', { name: 'ออกจากระบบ' }));
    });
    const { logout } = useAuth();
    expect(logout).toHaveBeenCalledTimes(1);
  });

  it('กด Esc → ปิด dropdown โปรไฟล์', () => {
    renderLayout('/');
    fireEvent.click(screen.getByRole('button', { name: 'เมนูผู้ใช้' }));
    expect(screen.getByRole('menu')).toBeTruthy();
    fireEvent.keyDown(window, { key: 'Escape' });
    expect(screen.queryByRole('menu')).toBeNull();
  });

  it('ปุ่มสามขีดเปิด drawer เมนูเสริม + overlay คลุมจอ; กด Esc ปิดกลับ', () => {
    const { container } = renderLayout('/');
    const sb = asideOf(container);
    expect(sb.className).toContain('-translate-x-full');
    fireEvent.click(screen.getByRole('button', { name: 'เปิดเมนู' }));
    expect(sb.className).toContain('translate-x-0');
    const overlay = [...container.querySelectorAll('div')].find((d) =>
      d.className.includes('bg-black/40')
    );
    expect(overlay).toBeTruthy();
    fireEvent.keyDown(window, { key: 'Escape' });
    expect(sb.className).toContain('-translate-x-full');
  });

  it('Drawer = Clean List: เมนูเสริมทั้ง 4 กางพร้อมกันทันที — ไม่มี accordion "เพิ่มเติม" ให้กดซ้ำซ้อน', () => {
    const { container } = renderLayout('/');
    fireEvent.click(screen.getByRole('button', { name: 'เปิดเมนู' }));
    const aside = asideOf(container);
    expect(aside.textContent).toContain('เพิ่มเติม');
    expect(aside.querySelector('a[href="/users"]')).toBeTruthy();
    // accordion ถูกถอนถาวร — ไม่มีปุ่มหัวข้อ "เพิ่มเติม" ให้กดกางอีกต่อไป
    expect(screen.queryByRole('button', { name: /เพิ่มเติม/ })).toBeNull();
    // รายการเมนูเสริมโชว์ทันทีเมื่อเปิด drawer — ไม่ต้องกดกางก่อน
    ['/asset', '/qr', '/history', '/report'].forEach((p) =>
      expect(aside.querySelector(`a[href="${p}"]`)).toBeTruthy()
    );
    // เมนูหลัก 5 รายการอยู่บน pills เท่านั้น — ไม่ซ้ำใน drawer
    ['/', '/farm', '/bundle', '/stock', '/scan'].forEach((p) =>
      expect(aside.querySelector(`a[href="${p}"]`)).toBeNull()
    );
  });

  it('Drawer polish: ไอคอนนำหน้าทุกรายการ + hover ไฮไลท์น้ำเงินโปร่งแสง + เส้นคั่นบางๆ ใต้หัวข้อ', () => {
    const { container } = renderLayout('/');
    fireEvent.click(screen.getByRole('button', { name: 'เปิดเมนู' }));
    const aside = asideOf(container);
    const asset = aside.querySelector('a[href="/asset"]');
    // hover state นุ่มนวล — ไฮไลท์พื้นหลังน้ำเงินโปร่งแสง + ตัวหนังสือขาวสว่าง
    expect(asset.className).toContain('hover:bg-[var(--blue)]/15');
    expect(asset.className).toContain('hover:text-white');
    // ไอคอนนำหน้าแต่ละเมนูแสดงจริง
    expect(asset.querySelector('.msi')).toBeTruthy();
    // เส้นคั่นบางๆ (border-bottom จาก var(--border)) คั่นหัวข้อกับรายการเมนู
    const heading = [...aside.querySelectorAll('div')].find(
      (d) => d.textContent === 'เพิ่มเติม' && d.className.includes('border-b')
    );
    expect(heading).toBeTruthy();
    expect(heading.className).toContain('border-[var(--border)]/25');
  });

  it('เลือกเมนูใน drawer แล้ว drawer ปิดอัตโนมัติ', () => {
    const { container } = renderLayout('/');
    fireEvent.click(screen.getByRole('button', { name: 'เปิดเมนู' }));
    fireEvent.click(screen.getByRole('link', { name: 'Admin Tools' }));
    expect(asideOf(container).className).toContain('-translate-x-full');
  });
});

describe('Layout — โครงเดิมที่ยังต้องคงอยู่', () => {
  it('เนื้อหาเต็มความกว้าง — ไม่มี margin-left กันที่ให้ sidebar', () => {
    const { container } = renderLayout('/');
    const main = container.querySelector('main').parentElement;
    expect(main.className).not.toContain('lg:ml-[var(--sb-w)]');
  });

  it('bottom nav มือถือยังอยู่ครบ: tile active เป็นโปร่งแสง --blue/25', () => {
    const { container } = renderLayout('/');
    const links = [...container.querySelectorAll('a')].filter(
      (a) => a.getAttribute('href') === '/'
    );
    const bottomNav = links[links.length - 1]; // ลำดับ DOM: pills มาก่อน bottom nav
    expect(bottomNav.querySelector('span').className).toContain('bg-[var(--blue)]/25');
  });
});