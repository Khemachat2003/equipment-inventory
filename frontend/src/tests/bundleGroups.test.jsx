// tests/bundleGroups.test.jsx
// ก้อน ⑥ (Issue A) — ตรรกะจัดกลุ่มหน้า "ชุดติดตั้งฟาร์ม (Bundle)":
// ชุดไหนอยู่คลัง / ชุดไหนอยู่ฟาร์มไหน (ใช้เป็นทั้งเกณฑ์จัดกลุ่มโซนและนับ chip ฟาร์ม)
import { describe, it, expect } from 'vitest';
import { farmKeyOf, STOCK_KEY } from '../utils/bundleGroups.js';

describe('farmKeyOf — ชุดอยู่คลังหรือฟาร์มไหน', () => {
  it('In Stock = คลัง (คีย์ว่าง) เสมอ แม้ farmId หลงเหลืออยู่', () => {
    expect(farmKeyOf({ status: 'In Stock', farmId: 'SITE-1' })).toBe('');
  });

  it('Deployed ใช้ farmId เป็นคีย์หลัก', () => {
    expect(farmKeyOf({ status: 'Deployed', farmId: 'SITE-1', farmName: 'ฟาร์มหนองแค' })).toBe('SITE-1');
  });

  it('ไม่มี farmId → fallback ชื่อฟาร์ม → ตำแหน่ง (ชุดเก่าที่ไม่มี farmId ก็จับกลุ่มได้)', () => {
    expect(farmKeyOf({ status: 'Deployed', farmName: 'ฟาร์ม A' })).toBe('ฟาร์ม A');
    expect(farmKeyOf({ status: 'Deployed', location: 'ฟาร์ม B' })).toBe('ฟาร์ม B');
  });

  it('ชุดนอกคลังที่ไม่มีข้อมูลฟาร์มเลย = คลัง (ไม่หลุดจากจอ)', () => {
    expect(farmKeyOf({ status: 'Maintenance' })).toBe('');
  });

  it('ตัดช่องว่างหัวท้ายของคีย์ฟาร์ม', () => {
    expect(farmKeyOf({ status: 'Deployed', farmId: '  SITE-1  ' })).toBe('SITE-1');
  });

  it('STOCK_KEY คือค่าพิเศษของ chip "คลัง"', () => {
    expect(STOCK_KEY).toBe('__stock__');
  });
});