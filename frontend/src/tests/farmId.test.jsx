// farmId.test.jsx — ตัวช่วย auto รหัสฟาร์ม/โรงเรือน (ใช้ใน Farm/TransferModal/DeployModal)
import { describe, it, expect } from 'vitest';
import { cleanId, suggestSiteId, suggestHouseId, FARM_TYPES } from '../utils/farmId.js';

describe('farmId — auto รหัสฟาร์ม/โรงเรือน', () => {
  it('suggestSiteId เอาเฉพาะ A-Z0-9 + uppercase', () => {
    expect(suggestSiteId('Farm Bangpa 01')).toBe('FARMBANGPA01');
    expect(suggestSiteId('ฟาร์มไทยล้วน')).toBe(''); // ชื่อไทยล้วน → ให้ user กรอกรหัสเอง
  });

  it('cleanId ตัดสัญลักษณ์ + จำกัดความยาว', () => {
    expect(cleanId('abcdefghij-kl', 5)).toBe('ABCDE');
    expect(cleanId('', 5)).toBe('');
  });

  it('suggestHouseId ไล่เลข H01, H02 ตามที่มีอยู่แล้วในฟาร์ม', () => {
    expect(suggestHouseId('SITE001', [])).toBe('SITE001-H01');
    expect(suggestHouseId('SITE001', ['SITE001-H01', 'SITE001-H02'])).toBe('SITE001-H03');
    expect(suggestHouseId('SITE001', ['site001-h07'])).toBe('SITE001-H08');
    expect(suggestHouseId('SITE001', ['OTHER-H09'])).toBe('SITE001-H01'); // คนละฟาร์มไม่นับ
    expect(suggestHouseId('', [])).toBe('');
  });

  it('FARM_TYPES ครบ 4 ประเภท', () => {
    expect(FARM_TYPES).toEqual(['สัตว์ปีก', 'สัตว์บก', 'สุกร', 'อื่นๆ']);
  });
});