// navigation.test.jsx — เทสต์โครงเมนูหลัง UX simplify + Step 3 (Farm Monitor ขึ้นเมนูหลัก)
import { describe, it, expect } from 'vitest';
import { NAV_GROUPS, MOBILE_NAV, ROUTE_META, getRouteMeta } from '../navigation.js';
import { UI_FLAGS } from '../uiConfig.js';

describe('navigation — โครงเมนูใหม่ (Step 3: Farm Monitor ขึ้นเมนูหลัก)', () => {
  it('งานประจำวัน 5 ปุ่ม: หน้าแรก / Farm Monitor / Bundle / เบิก–คืนของ / ย้าย–โอนอุปกรณ์ (1-Click ทั้งคู่)', () => {
    const main = NAV_GROUPS.find((g) => g.label === 'งานประจำวัน');
    expect(main.items.map((i) => i.to)).toEqual(['/', '/farm', '/bundle', '/stock', '/scan']);
  });

  it('Farm Monitor (/farm) ไม่ถูกซ่อนในซับเมนู "เพิ่มเติม" อีกต่อไป (Step 3 — unnesting)', () => {
    const more = NAV_GROUPS.find((g) => g.label === 'เพิ่มเติม');
    expect(more.items.map((i) => i.to)).not.toContain('/farm');
    expect(more.items.map((i) => i.to)).not.toContain('/bundle');
  });

  it('กลุ่ม "เพิ่มเติม" ถูกพับเก็บไว้ (collapsed) และมีหน้าข้อมูลครบทุกหน้า', () => {
    const more = NAV_GROUPS.find((g) => g.label === 'เพิ่มเติม');
    expect(more.collapsed).toBe(true);
    const tos = more.items.map((i) => i.to);
    ['/asset', '/qr', '/history', '/report', '/dashboard'].forEach((p) =>
      expect(tos).toContain(p)
    );
  });

  it('/dashboard (Dashboard แบบเต็ม) เป็นของ admin เท่านั้น', () => {
    const more = NAV_GROUPS.find((g) => g.label === 'เพิ่มเติม');
    expect(more.items.find((i) => i.to === '/dashboard').adminOnly).toBe(true);
  });

  it('กลุ่มผู้ดูแลระบบ adminOnly + settings ย้ายมารวมอยู่ที่นี่', () => {
    const admin = NAV_GROUPS.find((g) => g.label === 'ผู้ดูแลระบบ');
    expect(admin.adminOnly).toBe(true);
    expect(admin.items.map((i) => i.to)).toContain('/settings');
  });

  it('bottom nav มือถือ 4 ปุ่ม', () => {
    expect(MOBILE_NAV.map((i) => i.to)).toEqual(['/', '/stock', '/scan', '/qr']);
  });

  it('ROUTE_META ลงทะเบียนครบทุก route ในเมนู', () => {
    [...NAV_GROUPS.flatMap((g) => g.items), ...MOBILE_NAV].forEach((i) => {
      expect(ROUTE_META[i.to]).toBeTruthy();
    });
  });

  it('ROUTE_META ของ /farm เปลี่ยนชื่อเป็น Farm Monitor (Step 3)', () => {
    expect(ROUTE_META['/farm'].title).toBe('Farm Monitor');
  });

  it('getRouteMeta: path ไม่รู้จัก → กลับหน้าแรก', () => {
    expect(getRouteMeta('/path-ไม่มีจริง')).toBe(ROUTE_META['/']);
  });
});

describe('UI_FLAGS — หลักการ "มีแต่ไม่แสดง"', () => {
  it('กราฟซ่อน / เมนูแบบย่อเปิด', () => {
    expect(UI_FLAGS.charts).toBe(false);
    expect(UI_FLAGS.simpleMenu).toBe(true);
  });
});