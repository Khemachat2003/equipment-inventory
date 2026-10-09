import { describe, it, expect, beforeEach } from 'vitest';
import { saveReprint, readReprint, REPRINT_KEY } from '../utils/labelReprint.js';

const CFG = { mode: 'a4', lW: 50, lH: 25, lMargin: 2, pm: 10, gapX: 3, gapY: 3, bgColor: '#ffffff', showBadge: false };

describe('labelReprint — สแนปช็อตพิมพ์ครั้งล่าสุด (ปุ่มพิมพ์ซ้ำ)', () => {
  beforeEach(() => { localStorage.clear(); });

  it('บันทึกแล้วอ่านกลับได้ตรง (cfg + serials + at)', () => {
    const snap = saveReprint(CFG, ['SN-A', 'SN-B']);
    expect(snap.serials).toEqual(['SN-A', 'SN-B']);
    const r = readReprint();
    expect(r.cfg).toEqual(CFG);
    expect(r.serials).toEqual(['SN-A', 'SN-B']);
    expect(r.at).toBeGreaterThan(0);
  });

  it('serials ว่าง → ไม่บันทึก และ readReprint คืน null', () => {
    expect(saveReprint(CFG, [])).toBeNull();
    expect(readReprint()).toBeNull();
  });

  it('ข้อมูลใน localStorage เสีย (JSON ไม่ถูก) → คืน null ไม่ throw', () => {
    localStorage.setItem(REPRINT_KEY, '{oops');
    expect(readReprint()).toBeNull();
  });

  it('บันทึกทับของเดิมได้ (กดพิมพ์รอบใหม่ = สแนปช็อตใหม่)', () => {
    saveReprint(CFG, ['SN-A']);
    saveReprint({ ...CFG, mode: 'label' }, ['SN-B', 'SN-C']);
    const r = readReprint();
    expect(r.cfg.mode).toBe('label');
    expect(r.serials).toEqual(['SN-B', 'SN-C']);
  });
});