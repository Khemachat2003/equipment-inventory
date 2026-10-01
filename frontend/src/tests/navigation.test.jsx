// navigation.test.jsx — เทสต์โครงเมนูหลัง UX simplify (เมนู 3 ปุ่ม + เพิ่มเติม พับเก็บ)
import { describe, it, expect } from 'vitest';
import { NAV_GROUPS, HELP_NAV, MOBILE_NAV, ROUTE_META, getRouteMeta } from '../navigation.js';
import { UI_FLAGS } from '../uiConfig.js';

describe('navigation — โครงเมนูใหม่ (UX simplify)', () => {
  it('งานประจำวันมี 3 ปุ่มพอดี: หน้าแรก / เบิก–คืนของ / ย้าย–โอนอุปกรณ์', () => {
    const main = NAV_GROUPS.find((g) => g.label === 'งานประจำวัน');
    expect(main.items.map((i) => i.to)).toEqual(['/', '/stock', '/scan']);
  });

  it('กลุ่ม "เพิ่มเติม" ถูกพับเก็บไว้ (collapsed) และมีหน้าเดิมครบทุกหน้า', () => {
    const more = NAV_GROUPS.find((g) => g.label === 'เพิ่มเติม');
    expect(more.collapsed).toBe(true);
    const tos = more.items.map((i) => i.to);
    ['/farm', '/asset', '/bundle', '/qr', '/history', '/report', '/dashboard'].forEach((p) =>
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

  it('มีทางเข้าคู่มือ (/help) และ bottom nav มือถือ 3 ปุ่ม', () => {
    expect(HELP_NAV.to).toBe('/help');
    expect(MOBILE_NAV.map((i) => i.to)).toEqual(['/', '/stock', '/scan']);
  });

  it('ROUTE_META ลงทะเบียนครบทุก route ในเมนู + /help', () => {
    [...NAV_GROUPS.flatMap((g) => g.items), HELP_NAV, ...MOBILE_NAV].forEach((i) => {
      expect(ROUTE_META[i.to]).toBeTruthy();
    });
  });

  it('getRouteMeta: /help → วิธีใช้งาน, path ไม่รู้จัก → กลับหน้าแรก', () => {
    expect(getRouteMeta('/help').title).toBe('วิธีใช้งาน');
    expect(getRouteMeta('/path-ไม่มีจริง')).toBe(ROUTE_META['/']);
  });
});

describe('UI_FLAGS — หลักการ "มีแต่ไม่แสดง"', () => {
  it('กราฟซ่อน / เมนูแบบย่อเปิด / ทัวร์เปิด', () => {
    expect(UI_FLAGS.charts).toBe(false);
    expect(UI_FLAGS.simpleMenu).toBe(true);
    expect(UI_FLAGS.onboardingTour).toBe(true);
  });
});