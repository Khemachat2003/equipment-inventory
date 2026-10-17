// assetLocationFilter.test.jsx — deep-link จากการ์ดสรุปสถานะหน้าแรก → /asset
//   • ?location=stock → เฉพาะอุปกรณ์ที่ยังอยู่คลังกลาง (ตามตัวเลขการ์ด "อยู่ในคลัง" เป๊ะ)
//   • ?location=farm  → เฉพาะอุปกรณ์ที่ผูกเข้ากับฟาร์ม/โรงเรือนแล้ว (ตามการ์ด "ติดตั้งแล้ว")
//   • ?status=repair  → กลุ่มสถานะต้องซ่อม: ส่งซ่อม/รอซ่อม/ชำรุด·สูญหาย (ตามการ์ด "ส่งซ่อม")
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import Asset from '../pages/Asset.jsx';

vi.mock('axios', () => ({ default: { get: vi.fn(), post: vi.fn() } }));
import axios from 'axios';

const ASSETS = [
  // คลังกลาง (สถานะ ใช้งานได้)
  { assetId: '0001', partNumber: 'SENSOR', code: 'SENSOR', name: 'เซนเซอร์อุณหภูมิ', serialNumber: 'SN-1', status: 'ใช้งานได้', location: 'Stock', siteName: 'Intranin' },
  // สมาชิกชุดที่ deploy ที่ฟาร์ม A (Location = ฟาร์ม)
  { assetId: '0002', partNumber: 'SENSOR', code: 'SENSOR', name: 'เซนเซอร์อุณหภูมิ', serialNumber: 'SN-2', status: 'In Bundle', location: 'ฟาร์ม A', siteName: 'ชุดตู้ควบคุม 1', bundleId: 'BDL-1', bundleName: 'ชุดตู้ควบคุม 1' },
  // ตัวเดี่ยวที่ฟาร์ม B — ส่งซ่อม
  { assetId: '0003', partNumber: 'SENSOR', code: 'SENSOR', name: 'เซนเซอร์ความชื้น', serialNumber: 'SN-3', status: 'ส่งซ่อม', location: 'โรงเรือน 2', siteName: 'ฟาร์ม B', houseName: 'โรงเรือน 2' },
  // คลังกลาง — ชำรุด/สูญหาย
  { assetId: '0004', partNumber: 'SENSOR', code: 'SENSOR', name: 'เซนเซอร์แสง', serialNumber: 'SN-4', status: 'ชำรุด/สูญหาย', location: 'Stock', siteName: 'Intranin' },
  // ตัวเดี่ยวที่ฟาร์ม B — รอซ่อม (สถานะต่างจาก "ส่งซ่อม" แต่ต้องนับในกลุ่มต้องซ่อม)
  { assetId: '0005', partNumber: 'SENSOR', code: 'SENSOR', name: 'เซนเซอร์ฝุ่น', serialNumber: 'SN-5', status: 'รอซ่อม', location: 'โรงเรือน 1', siteName: 'ฟาร์ม B' },
];

beforeEach(() => {
  axios.get.mockImplementation((url) => {
    if (url === '/api/assets') return Promise.resolve({ data: ASSETS });
    if (url === '/api/part-catalog') return Promise.resolve({ data: [{ partNumber: 'SENSOR', partName: 'เซนเซอร์', category: 'Sensors & Weather' }] });
    return Promise.resolve({ data: [] });
  });
  axios.post.mockResolvedValue({ data: { success: true } });
});

// ทดสอบ deep-link ทีละเคส — ตั้ง URL จริงก่อน mount (Asset อ่าน window.location.search ตรง ๆ)
function goto(query) {
  window.history.replaceState(null, '', `/asset${query}`);
}
afterEach(() => {
  window.history.replaceState(null, '', '/asset');
});

async function renderPage() {
  const utils = render(<Asset />);
  await screen.findByTestId('part-sidebar'); // รอโหลดเสร็จ
  return utils;
}

describe('Asset — deep-link ตัวกรองตำแหน่งจากการ์ดหน้าแรก', () => {
  it('?location=stock → เห็นเฉพาะชิ้นในคลังกลาง (2) พร้อมชิป "ตำแหน่ง: คลังกลาง"', async () => {
    goto('?location=stock');
    await renderPage();
    expect(screen.queryAllByText('SN-1').length).toBeGreaterThan(0); // คลัง (Intranin/Stock)
    expect(screen.queryAllByText('SN-4').length).toBeGreaterThan(0); // คลัง (ชำรุด)
    expect(screen.queryAllByText('SN-2')).toHaveLength(0); // ชุด deploy ที่ฟาร์ม A — ไม่โชว์
    expect(screen.queryAllByText('SN-3')).toHaveLength(0); // ฟาร์ม B — ไม่โชว์
    expect(screen.queryAllByText('SN-5')).toHaveLength(0); // ฟาร์ม B — ไม่โชว์
    expect(screen.getByTestId('kpi-total').textContent).toContain('2');
    expect(screen.getByText('ตำแหน่ง: คลังกลาง')).toBeTruthy();
  });

  it('?location=farm → เห็นเฉพาะชิ้นที่ผูกเข้ากับฟาร์มแล้ว (3) พร้อมชิปตำแหน่ง', async () => {
    goto('?location=farm');
    await renderPage();
    expect(screen.queryAllByText('SN-2').length).toBeGreaterThan(0); // ชุด deploy ที่ฟาร์ม A
    expect(screen.queryAllByText('SN-3').length).toBeGreaterThan(0); // ฟาร์ม B
    expect(screen.queryAllByText('SN-5').length).toBeGreaterThan(0); // ฟาร์ม B
    expect(screen.queryAllByText('SN-1')).toHaveLength(0); // คลัง — ไม่โชว์
    expect(screen.queryAllByText('SN-4')).toHaveLength(0); // คลัง — ไม่โชว์
    expect(screen.getByTestId('kpi-total').textContent).toContain('3');
    expect(screen.getByText('ตำแหน่ง: ติดตั้งที่ฟาร์ม/โรงเรือน')).toBeTruthy();
  });

  it('?status=repair → รวมทั้ง ส่งซ่อม / รอซ่อม / ชำรุด (3) พร้อมชิป "สถานะ: ส่งซ่อม / ชำรุด"', async () => {
    goto('?status=repair');
    await renderPage();
    expect(screen.queryAllByText('SN-3').length).toBeGreaterThan(0); // ส่งซ่อม
    expect(screen.queryAllByText('SN-4').length).toBeGreaterThan(0); // ชำรุด/สูญหาย
    expect(screen.queryAllByText('SN-5').length).toBeGreaterThan(0); // รอซ่อม
    expect(screen.queryAllByText('SN-1')).toHaveLength(0); // ใช้งานได้ — ไม่โชว์
    expect(screen.queryAllByText('SN-2')).toHaveLength(0); // In Bundle (ที่ฟาร์ม) — ไม่โชว์
    expect(screen.getByTestId('kpi-total').textContent).toContain('3');
    expect(screen.getByText('สถานะ: ส่งซ่อม / ชำรุด')).toBeTruthy();
  });

  it('กดปุ่ม x บนชิปตำแหน่ง → ล้างตัวกรอง เห็นทุกชิ้น (5)', async () => {
    goto('?location=stock');
    await renderPage();
    fireEvent.click(screen.getByRole('button', { name: 'ล้างตัวกรองตำแหน่ง' }));
    await waitFor(() => expect(screen.queryByRole('button', { name: 'ล้างตัวกรองตำแหน่ง' })).toBeNull());
    expect(screen.queryAllByText('SN-2').length).toBeGreaterThan(0);
    expect(screen.queryAllByText('SN-3').length).toBeGreaterThan(0);
    expect(screen.getByTestId('kpi-total').textContent).toContain('5');
  });
});
