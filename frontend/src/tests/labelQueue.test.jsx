// labelQueue.test.jsx — ตัวช่วยหน้าพิมพ์ฉลาก QR (/qr): เรียงล่าสุดอยู่บน + ตัวกรองวันรับเข้า/เลข PO + จัดกลุ่มตามวันรับเข้า
import { describe, it, expect } from 'vitest';
import {
  serialDate,
  createdTsOf,
  receivedTsOf,
  dateKeyOf,
  thaiDayLabel,
  compareLatestFirst,
  matchesLabelFilters,
  lastNDaysRange,
  groupRowsByDate,
  serialsFromSearchParams,
} from '../utils/labelQueue.js';

describe('serialDate — แยกวันที่จาก Serial (SN-PART-DDMMYYYY-รัน)', () => {
  it('อ่านวันที่ DDMMYYYY ได้ (21/06/2025)', () => {
    const d = new Date(serialDate('SN-PCB001-21062025-0001'));
    expect([d.getDate(), d.getMonth() + 1, d.getFullYear()]).toEqual([21, 6, 2025]);
  });

  it('Part มีขีดกลางก็อ่านได้ (ยึด segment วันที่ก่อนเลขรันท้ายสุด)', () => {
    const d = new Date(serialDate('SN-SENSOR-1-09102026-0004'));
    expect([d.getDate(), d.getMonth() + 1, d.getFullYear()]).toEqual([9, 10, 2026]);
  });

  it('Serial ไม่มีวันที่/ค่าว่าง → null', () => {
    expect(serialDate('SN-ABC-0001')).toBeNull();
    expect(serialDate('')).toBeNull();
    expect(serialDate(undefined)).toBeNull();
  });

  it('วันที่ไม่มีจริง (31/02/2025) และเดือนเกิน 12 → null', () => {
    expect(serialDate('SN-X-31022025-0001')).toBeNull();
    expect(serialDate('SN-X-15132025-0001')).toBeNull();
  });
});

describe('createdTsOf — วันที่สร้างรายชิ้น (createdAt > Serial > receivedAt)', () => {
  it('ใช้ createdAt ก่อนถ้ามี (โหมด Local Inventory)', () => {
    const asset = { createdAt: '2026-10-10T00:00:00.000Z', serialNumber: 'SN-X-01012025-0001' };
    expect(createdTsOf(asset)).toBe(Date.parse('2026-10-10T00:00:00.000Z'));
  });

  it('ไม่มี createdAt → ใช้วันที่ฝังใน Serial', () => {
    const ts = createdTsOf({ serialNumber: 'SN-X-15062025-0002' });
    expect(ts).toBe(serialDate('SN-X-15062025-0002'));
  });

  it('ไม่มีทั้งสอง → ใช้ receivedAt / ไม่มีอะไรเลย → 0', () => {
    expect(createdTsOf({ receivedAt: '2026-03-01T00:00:00.000Z' })).toBe(Date.parse('2026-03-01T00:00:00.000Z'));
    expect(createdTsOf({})).toBe(0);
  });
});

describe('receivedTsOf — วันที่รับเข้า (receivedAt ก่อน, ไม่มี → วันที่สร้าง)', () => {
  it('ใช้ receivedAt จริงก่อน', () => {
    const asset = { receivedAt: '2026-01-02T00:00:00.000Z', serialNumber: 'SN-X-01012025-0001' };
    expect(receivedTsOf(asset)).toBe(Date.parse('2026-01-02T00:00:00.000Z'));
  });

  it('อุปกรณ์เพิ่มใหม่ที่ยังไม่ผูกล็อต → ถือว่ารับเข้าวันที่สร้าง (โผล่ใน "รับเข้า 2 วันล่าสุด" เสมอ)', () => {
    expect(receivedTsOf({ serialNumber: 'SN-X-01062025-0001' })).toBe(serialDate('SN-X-01062025-0001'));
  });
});

describe('dateKeyOf / thaiDayLabel — key และป้ายวันแบบเวลาไทย (Asia/Bangkok)', () => {
  it('02:00 ของไทย (ข้ามวัน UTC แล้ว) → key ยังเป็นวันเดียวกันของไทย', () => {
    expect(dateKeyOf(Date.parse('2025-06-15T02:00:00+07:00'))).toBe('2025-06-15');
    expect(dateKeyOf(0)).toBe('');
    expect(dateKeyOf(Number.NaN)).toBe('');
  });

  it('ป้ายวันแบบไทย (ปี พ.ศ.) เช่น 16 มิ.ย. 2568', () => {
    const ts = Date.parse('2025-06-16T10:00:00+07:00');
    expect(thaiDayLabel(ts)).toMatch(/16\s*มิ\.ย\.\s*2568/);
    expect(thaiDayLabel(0)).toBe('');
  });
});

describe('compareLatestFirst — เรียง "ล่าสุดอยู่บนสุดเสมอ" (Created Date DESC)', () => {
  it('วันสร้างมาก่อนอยู่บน; วันเดียวกัน → เลขรันท้าย Serial มากกว่าอยู่บน', () => {
    const old1 = { serialNumber: 'SN-A-01012025-0001' };
    const new1 = { serialNumber: 'SN-A-15062025-0001' };
    const new2 = { serialNumber: 'SN-A-15062025-0002' };
    const sorted = [old1, new2, new1].sort(compareLatestFirst).map((a) => a.serialNumber);
    expect(sorted).toEqual(['SN-A-15062025-0002', 'SN-A-15062025-0001', 'SN-A-01012025-0001']);
  });

  it('createdAt แม่นระดับเวลา (Local Inventory) → ชิ้นที่สร้างทีหลังอยู่บน', () => {
    const earlier = { createdAt: '2026-01-01T09:00:00.000Z', serialNumber: 'SN-X-01012026-0001' };
    const later = { createdAt: '2026-01-01T10:00:00.000Z', serialNumber: 'SN-X-01012026-0002' };
    expect([earlier, later].sort(compareLatestFirst)[0]).toBe(later);
  });
});

describe('matchesLabelFilters — ตัวกรองรวม (ค้นหา/สถานะ/เลข PO/ช่วงวันรับเข้า)', () => {
  const asset = {
    serialNumber: 'SN-X-15062025-0001',
    name: 'เซนเซอร์อุณหภูมิ',
    status: 'ใช้งานได้',
    poNumber: 'PO-99',
    receivedAt: '2025-06-15T04:00:00+07:00',
  };

  it('ค้นหาได้ทั้ง Serial/ชื่อ (ตัวพิมพ์เล็ก-ใหญ่เท่ากัน)', () => {
    expect(matchesLabelFilters(asset, { q: 'sn-x-1506' })).toBe(true);
    expect(matchesLabelFilters(asset, { q: 'เซนเซอร์' })).toBe(true);
    expect(matchesLabelFilters(asset, { q: 'ไม่มีแน่นอน' })).toBe(false);
  });

  it('กรองสถานะตรงตัว (all = ทั้งหมด)', () => {
    expect(matchesLabelFilters(asset, { status: 'ใช้งานได้' })).toBe(true);
    expect(matchesLabelFilters(asset, { status: 'สำรอง' })).toBe(false);
    expect(matchesLabelFilters(asset, { status: 'all' })).toBe(true);
  });

  it('กรองตามเลข PO (ตัวพิมพ์เล็กก็เจอ)', () => {
    expect(matchesLabelFilters(asset, { po: 'po-9' })).toBe(true);
    expect(matchesLabelFilters(asset, { po: 'PO-100' })).toBe(false);
  });

  it('กรองช่วงวันรับเข้า (from/to แบบ key YYYY-MM-DD)', () => {
    expect(matchesLabelFilters(asset, { from: '2025-06-15' })).toBe(true);
    expect(matchesLabelFilters(asset, { from: '2025-06-16' })).toBe(false);
    expect(matchesLabelFilters(asset, { to: '2025-06-15' })).toBe(true);
    expect(matchesLabelFilters(asset, { to: '2025-06-14' })).toBe(false);
    expect(matchesLabelFilters(asset, { from: '2025-06-15', to: '2025-06-15' })).toBe(true);
  });

  it('อุปกรณ์ไม่มีวันที่ให้อ้าง + เปิดตัวกรองวัน → ตัดออก (เหมือนหน้าทะเบียน)', () => {
    expect(matchesLabelFilters({ serialNumber: 'SN-PLAIN-001' }, { from: '2025-06-01' })).toBe(false);
    expect(matchesLabelFilters({ serialNumber: 'SN-PLAIN-001' }, {})).toBe(true);
  });
});

describe('lastNDaysRange — ช่วง "N วันล่าสุด" สำหรับปุ่มด่วน', () => {
  const now = new Date('2026-10-15T10:00:00+07:00');

  it('2 วันล่าสุด = เมื่อวานถึงวันนี้ (รวมสองวัน)', () => {
    expect(lastNDaysRange(2, now)).toEqual({ from: '2026-10-14', to: '2026-10-15' });
  });

  it('1 วันล่าสุด = วันนี้เท่านั้น', () => {
    expect(lastNDaysRange(1, now)).toEqual({ from: '2026-10-15', to: '2026-10-15' });
  });
});

describe('groupRowsByDate — จัดกลุ่มตามวันรับเข้า (กลุ่มไม่รู้วัน ท้ายสุด)', () => {
  const rows = [
    { serialNumber: 'SN-B-16062025-0002', receivedAt: '2025-06-16T09:00:00+07:00' },
    { serialNumber: 'SN-B-16062025-0001', receivedAt: '2025-06-16T04:00:00+07:00' },
    { serialNumber: 'SN-A-15062025-0001', receivedAt: '2025-06-15T04:00:00+07:00' },
    { serialNumber: 'SN-PLAIN-0001' }, // ไม่มีวันที่เลย
  ];

  it('จับกลุ่มตามวัน + คงลำดับเดิมภายในกลุ่ม + ไม่รู้วันไปท้าย', () => {
    const groups = groupRowsByDate(rows);
    expect(groups.map((g) => g.key)).toEqual(['2025-06-16', '2025-06-15', 'unknown']);
    expect(groups[0].items.map((a) => a.serialNumber)).toEqual(['SN-B-16062025-0002', 'SN-B-16062025-0001']);
    expect(groups[2].label).toBe('ไม่ระบุวันที่รับเข้า');
  });

  it('ป้ายกลุ่มเป็นวันไทย (พ.ศ.)', () => {
    expect(groupRowsByDate(rows)[0].label).toMatch(/16\s*มิ\.ย\.\s*2568/);
  });

  it('รายการว่าง → ไม่มีกลุ่ม', () => {
    expect(groupRowsByDate([])).toEqual([]);
  });
});

describe('serialsFromSearchParams — deep-link เข้าหน้าพิมพ์ฉลาก', () => {
  it('?serial=SN-1 (แบบเดิมจากหน้าทะเบียน)', () => {
    expect(serialsFromSearchParams(new URLSearchParams('serial=SN-1'))).toEqual(['SN-1']);
  });

  it('?serials=SN-1,SN-2 รองรับทั้งคอมมาและช่องว่าง + ตัดซ้ำ', () => {
    expect(serialsFromSearchParams(new URLSearchParams('serials=SN-1,SN-2'))).toEqual(['SN-1', 'SN-2']);
    expect(serialsFromSearchParams(new URLSearchParams('serials=SN-1%20SN-2'))).toEqual(['SN-1', 'SN-2']);
    expect(serialsFromSearchParams(new URLSearchParams('serials=SN-1,SN-1'))).toEqual(['SN-1']);
  });

  it('มีทั้ง serials และ serial → ใช้ serials (ชุดพิมพ์ทันทีมาก่อน)', () => {
    expect(serialsFromSearchParams(new URLSearchParams('serials=SN-1&serial=SN-3'))).toEqual(['SN-1']);
  });

  it('ไม่มีพารามิเตอร์ → ไม่เลือกอะไร', () => {
    expect(serialsFromSearchParams(new URLSearchParams(''))).toEqual([]);
  });
});
