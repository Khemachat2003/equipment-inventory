// locationPath.test.jsx — Step 2: Farm Location Primary Indicator + Location Consistency
// ตรวจว่า pill ตำแหน่ง (variant="primary") โชว์ฟาร์ม/โรงเรือนทันที และ buildLocation
// อ่านตำแหน่งของอุปกรณ์ในชุดจากคอลัมน์ Location เสมอ (ตรงกันทุกหน้า Asset/Scan/Bundle)
import { describe, it, expect } from 'vitest';
import { render } from '@testing-library/react';
import LocationPath from '../components/ui/LocationPath.jsx';
import { buildLocation } from '../utils/location.js';

describe('LocationPath variant="primary" — ตำแหน่งฟาร์มเป็น Primary Indicator', () => {
  it('โชว์ ฟาร์ม › โรงเรือน เป็น pill เด่นทันที (ไม่ต้องเปิด modal)', () => {
    const { container } = render(
      <LocationPath variant="primary" siteName="ฟาร์ม A" houseName="โรงเรือน 1" houseId="F1-H01" location="จุดติดตั้ง A5" bundleId="" />
    );
    expect(container.textContent).toContain('ฟาร์ม A');
    expect(container.textContent).toContain('F1-H01 โรงเรือน 1');
    expect(container.textContent).toContain('จุดติดตั้ง A5');
    // pill มี title เก็บเส้นทางเต็ม + ไอคอนพิน
    expect(container.querySelector('[title]')).toBeTruthy();
  });

  it('อยู่คลังกลาง → pill แสดง คลังกลาง (Stock)', () => {
    const { container } = render(<LocationPath variant="primary" siteName="Intranin" location="Stock" />);
    expect(container.textContent).toContain('คลังกลาง');
  });

  it('variant default ยังทำงานเหมือนเดิม (ไม่ทำหน้าอื่นพัง)', () => {
    const { container } = render(<LocationPath siteName="ฟาร์ม A" houseName="โรงเรือน 1" location="-" />);
    expect(container.textContent).toContain('ฟาร์ม A');
    expect(container.textContent).toContain('โรงเรือน 1');
  });
});

describe('Location consistency — อุปกรณ์ในชุดต้องโชว์ฟาร์มเดียวกันทุกหน้า', () => {
  it('อุปกรณ์ในชุด: siteName = ชื่อชุด, ฟาร์มอ่านจากคอลัมน์ Location — ไม่เอาชื่อชุดมาเป็นฟาร์ม', () => {
    const loc = buildLocation({
      siteName: 'ชุดตู้ควบคุมพัดลม',
      location: 'ฟาร์ม B',
      houseName: 'โรงเรือน 2',
      bundleId: 'BDL-001',
    });
    expect(loc.full).toContain('ฟาร์ม B');
    expect(loc.full).toContain('โรงเรือน 2');
    expect(loc.full).not.toContain('ชุดตู้ควบคุมพัดลม');
  });

  it('อุปกรณ์เดี่ยวที่ฟาร์ม: siteName = ฟาร์ม, location = จุดติดตั้ง', () => {
    const loc = buildLocation({ siteName: 'ฟาร์ม A', houseName: 'โรงเรือน 1', location: 'จุดติดตั้ง A5' });
    expect(loc.full).toBe('ฟาร์ม A  ›  โรงเรือน 1  ›  จุดติดตั้ง A5');
    expect(loc.kind).toBe('site');
  });
});