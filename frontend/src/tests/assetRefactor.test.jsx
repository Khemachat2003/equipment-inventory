// assetRefactor.test.jsx — ทะเบียนรายชิ้น Full Refactor (B2 Strict Terminology + Enterprise UI)
// ตรวจ: ศัพท์ไทยล้วนบนหน้าจอ (ไม่มี Part / In Bundle / (Stock) หลุดขึ้นจอ),
//        KPI Glassmorphic ตรงตัวเลข, ตารางแถว 52px + ชื่อตัวหนา + Action Icons พร้อม tooltip
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, within } from '@testing-library/react';
import Asset from '../pages/Asset.jsx';

vi.mock('axios', () => ({ default: { get: vi.fn(), post: vi.fn() } }));
import axios from 'axios';

const CATALOG = [{ partNumber: 'SENSOR', partName: 'เซนเซอร์อุณหภูมิ', category: 'Sensors & Weather' }];
const ASSETS = [
  { assetId: '0001', partNumber: 'SENSOR', code: 'SENSOR', name: 'เซนเซอร์อุณหภูมิ', serialNumber: 'SN-1', status: 'ใช้งานได้', location: 'Stock', siteName: 'Intranin', batchId: 'INB-OLD', poNumber: 'PO-OLD', receivedAt: '2026-10-01T00:00:00.000Z' },
  { assetId: '0002', partNumber: 'SENSOR', code: 'SENSOR', name: 'เซนเซอร์อุณหภูมิ', serialNumber: 'SN-2', status: 'In Bundle', location: 'ฟาร์ม A', siteName: 'ชุดตู้ควบคุม 1', bundleId: 'BDL-1', bundleName: 'ชุดตู้ควบคุม 1', batchId: 'INB-OLD', poNumber: 'PO-OLD', receivedAt: '2026-10-02T00:00:00.000Z' },
  { assetId: '0003', partNumber: 'SENSOR', code: 'SENSOR', name: 'เซนเซอร์ความชื้น', serialNumber: 'SN-3', status: 'ส่งซ่อม', location: 'โรงเรือน 2', siteName: 'ฟาร์ม B', houseName: 'โรงเรือน 2', batchId: 'INB-NEW', poNumber: 'PO-NEW', receivedAt: '2026-10-09T00:00:00.000Z' },
];
const POS = [
  { key: 'po:PO-OLD', poNumber: 'PO-OLD', batches: [{ batchId: 'INB-OLD', count: 2 }], receivedAt: '2026-10-01T00:00:00.000Z', assetCount: 2, stockCount: 2, batchCount: 1, status: 'active', manuallyClosed: false },
];

beforeEach(() => {
  axios.get.mockImplementation((url) => {
    if (url === '/api/assets') return Promise.resolve({ data: ASSETS });
    if (url === '/api/part-catalog') return Promise.resolve({ data: CATALOG });
    if (url.startsWith('/api/inbound-pos')) return Promise.resolve({ data: url.includes('status=active') ? POS.filter((po) => po.status === 'active') : POS });
    return Promise.resolve({ data: [] });
  });
  axios.post.mockResolvedValue({ data: { success: true } });
});

async function renderPage() {
  const utils = render(<Asset />);
  await screen.findByText('PO PO-OLD'); // รอโหลดข้อมูลเสร็จ
  return utils;
}

describe('Asset (ทะเบียนรายชิ้น) — B2 Strict Terminology', () => {
  it('ไม่มีศัพท์ระบบ Part / In Bundle / (Stock) หลุดขึ้นหน้าจอ — แสดงภาษาไทยแทน', async () => {
    const { container } = await renderPage();
    const text = container.textContent;
    expect(text).not.toContain('Part');
    expect(text).not.toContain('In Bundle');
    expect(text).not.toContain('(Stock)');
    // สถานะข้อมูล "In Bundle" → "อยู่ในชุดติดตั้ง" · ตำแหน่ง Stock → "คลังกลาง"
    expect(text).toContain('อยู่ในชุดติดตั้ง');
    expect(text).toContain('คลังกลาง');
    expect(text).toContain('รายการอุปกรณ์'); // Sidebar + คอลัมน์หลัก (แทนคำว่า Part)
    // Mobile selector ก็ใช้ภาษาไทย
    expect(screen.getByText('ทุกรายการอุปกรณ์ (3)')).toBeTruthy();
  });
});

describe('Asset — KPI Glassmorphic Cards', () => {
  it('แสดงครบ อุปกรณ์ทั้งหมด / ใช้งานได้ / อยู่ในชุดติดตั้ง / ส่งซ่อม ตรงตัวเลขจริง', async () => {
    await renderPage();
    expect(within(screen.getByTestId('kpi-total')).getByText('3')).toBeTruthy();
    expect(within(screen.getByTestId('kpi-usable')).getByText('1')).toBeTruthy();
    expect(within(screen.getByTestId('kpi-bundled')).getByText('1')).toBeTruthy();
    expect(within(screen.getByTestId('kpi-repair')).getByText('1')).toBeTruthy();
  });
});

describe('Asset — Enterprise Data Table', () => {
  it('แถวสูงขั้นต่ำ 52px + ชื่ออุปกรณ์ตัวหนาสี ink อ่านชัด', async () => {
    const { container } = await renderPage();
    expect(container.querySelector('td[class*="h-[52px]"]')).toBeTruthy();
    const nameCell = screen.getAllByText('เซนเซอร์ความชื้น')[0];
    expect(nameCell.className).toContain('font-bold');
    expect(nameCell.className).toContain('text-[var(--ink)]');
  });

  it('ปุ่มจัดการเป็น Clean Action Icons พร้อม Hover Tooltip (มี aria-label + ข้อความ tooltip)', async () => {
    await renderPage();
    expect(screen.getAllByRole('button', { name: 'ดูประวัติการเคลื่อนไหว' })).toHaveLength(3);
    expect(screen.getAllByRole('button', { name: 'โอนย้ายอุปกรณ์' })).toHaveLength(3);
    expect(screen.getAllByRole('link', { name: 'พิมพ์ฉลาก QR' })).toHaveLength(3);
    // tooltip = span ลอย (pointer-events-none) ที่วางข้างปุ่ม — ข้อความตรงกับ action
    const tooltip = screen.getAllByRole('button', { name: 'โอนย้ายอุปกรณ์' })[0].parentElement.querySelector('span[class*="pointer-events-none"]');
    expect(tooltip.textContent).toBe('โอนย้ายอุปกรณ์');
  });
});

// ═══ Consolidate Action Buttons — ลบปุ่ม "รับเข้า" (สีเขียว) ที่ซ้ำซ้อนกับ "เพิ่มอุปกรณ์ใหม่" ═══
// เพิ่มอุปกรณ์ใหม่ = Primary Action เดียว — เปิด AddDeviceModal ที่มี Fast-Track Intake Flow
// (กรอกข้อมูล → จัดชุด → โอนย้าย → พิมพ์ฉลาก) ครอบคลุมทุกการรับเข้าแล้ว
describe('Asset — Consolidated Primary Action (Action Bar คลีน)', () => {
  it('Action Bar เหลือ Primary "เพิ่มอุปกรณ์ใหม่" ปุ่มเดียว + ปุ่มรีเฟรช — ไม่มีปุ่ม "รับเข้า" ซ้ำซ้อน', async () => {
    await renderPage();
    // Primary Blue — ปุ่มรับเข้าอุปกรณ์เพียงปุ่มเดียวของหน้า
    expect(screen.getAllByRole('button', { name: 'เพิ่มอุปกรณ์ใหม่' })).toHaveLength(1);
    // ปุ่มรีเฟรชข้อมูล (Icon Button)
    expect(screen.getByRole('button', { name: 'รีเฟรชข้อมูล' })).toBeTruthy();
    // ไม่มีปุ่ม "รับเข้า" ที่ทำหน้าที่ซ้ำซ้อนหลงเหลืออยู่
    expect(screen.queryByRole('button', { name: 'รับเข้า' })).toBeNull();
  });

  it('กด "เพิ่มอุปกรณ์ใหม่" → เปิด AddDeviceModal (Fast-Track Intake Flow ครบใน Modal เดียว)', async () => {
    await renderPage();
    fireEvent.click(screen.getByRole('button', { name: 'เพิ่มอุปกรณ์ใหม่' }));
    // Modal เปิด — หัว Modal "เพิ่มอุปกรณ์ใหม่" แสดงเพิ่มจากปุ่มใน Hero Bar + ฟอร์มรับเข้า (Step 1)
    expect(screen.getAllByText('เพิ่มอุปกรณ์ใหม่').length).toBeGreaterThanOrEqual(2);
    expect(screen.getByText('① อุปกรณ์อะไร? (พิมพ์ชื่อหรือ Part Number)')).toBeTruthy();
  });
});
