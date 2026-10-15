// assetFilters.test.jsx — Unified Search & Date Filter Bar + Sidebar หมวดหมู่ + เลือกอุปกรณ์ (ทะเบียนรายชิ้น)
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
  await screen.findByText('PO PO-OLD');
  return utils;
}

describe('Asset — Unified Filter Bar', () => {
  it('ช่องค้นหาเดียวค้นได้ทั้ง Serial/ชื่อ/รหัส และ ล็อต/PO/ผู้ขาย', async () => {
    await renderPage();
    const searchInput = screen.getByPlaceholderText(/ค้นหาชื่ออุปกรณ์/);
    // ค้นด้วยเลข PO → เจอเฉพาะอุปกรณ์ของ PO นั้น
    fireEvent.change(searchInput, { target: { value: 'PO-NEW' } });
    expect(screen.queryAllByText('SN-1')).toHaveLength(0);
    expect(screen.queryAllByText('SN-3').length).toBeGreaterThan(0);
    // ค้นด้วย Serial → เจอรายชิ้นตรง
    fireEvent.change(searchInput, { target: { value: 'SN-2' } });
    expect(screen.queryAllByText('SN-3')).toHaveLength(0);
    expect(screen.queryAllByText('SN-2').length).toBeGreaterThan(0);
  });

  it('มีช่องวันที่รูปแบบไทย 2 ช่อง (placeholder วว/ดด/ปปปป — ไม่ใช่ mm/dd/yyyy ของ browser)', async () => {
    await renderPage();
    expect(screen.getAllByPlaceholderText('วว/ดด/ปปปป')).toHaveLength(2);
    expect(screen.getByLabelText('รับเข้าตั้งแต่')).toBeTruthy();
    expect(screen.getByLabelText('รับเข้าถึง')).toBeTruthy();
  });

  it('กรองช่วงวันรับเข้าด้วยรูปแบบไทย (พิมพ์ทั้ง ค.ศ. และ พ.ศ. ได้)', async () => {
    await renderPage();
    const fromInput = screen.getByLabelText('รับเข้าตั้งแต่');
    const toInput = screen.getByLabelText('รับเข้าถึง');
    // ค.ศ. — 1–5 ต.ค. 2026: SN-1 (1 ต.ค.) / SN-2 (2 ต.ค.) เข้าเงื่อนไข, SN-3 (9 ต.ค.) หลุด
    fireEvent.change(fromInput, { target: { value: '01/10/2026' } });
    fireEvent.change(toInput, { target: { value: '05/10/2026' } });
    expect(screen.queryAllByText('SN-3')).toHaveLength(0);
    expect(screen.queryAllByText('SN-1').length).toBeGreaterThan(0);
    expect(screen.queryAllByText('SN-2').length).toBeGreaterThan(0);
    // พ.ศ. — กันยายน 2569 = กันยายน 2026: ทุกชิ้น (รับเข้าต.ค.) หลุดเงื่อนไขจบเดือน ก.ย.
    fireEvent.change(toInput, { target: { value: '30/09/2569' } });
    expect(screen.queryAllByText('SN-1')).toHaveLength(0);
  });
});

describe('Asset — Sidebar รายการอุปกรณ์', () => {
  it('หัว "รายการอุปกรณ์" + รายการ "ทุกรายการอุปกรณ์" + จัดกลุ่มหมวดหมู่ด้วยเส้นคั่นบาง', async () => {
    await renderPage();
    const sidebar = screen.getByTestId('part-sidebar');
    expect(within(sidebar).getByText('รายการอุปกรณ์')).toBeTruthy();
    expect(within(sidebar).getByText('ทุกรายการอุปกรณ์')).toBeTruthy();
    const groupHeader = within(sidebar).getByText('เซนเซอร์ & สภาพอากาศ').closest('div');
    expect(groupHeader.className).toContain('border-t');
    expect(groupHeader.className).toContain('border-[var(--border)]');
  });
});

describe('Asset — การเลือกอุปกรณ์ (Bulk select)', () => {
  it('checkbox หัวตารางเลือกทั้งหมดในหน้า → แถบสรุปจำนวนตรง และกดซ้ำเพื่อยกเลิกได้', async () => {
    await renderPage();
    const selectAll = screen.getByRole('checkbox', { name: 'เลือกทั้งหมดในหน้านี้' });
    fireEvent.click(selectAll);
    expect(screen.getByText(/เลือกแล้ว 3 ชิ้น/)).toBeTruthy();
    screen.getAllByRole('checkbox', { name: 'เลือก SN-1' }).forEach((cb) => expect(cb.checked).toBe(true));
    fireEvent.click(screen.getByText('ล้างที่เลือก'));
    expect(screen.queryByText(/เลือกแล้ว/)).toBeNull();
  });

  it('กดการ์ด PO → เลือกอุปกรณ์ทั้ง PO + ตัวกรองรวมช่องเดียวเป็นเลข PO นั้น', async () => {
    await renderPage();
    fireEvent.click(screen.getByText('PO PO-OLD'));
    expect(screen.getByText(/เลือกแล้ว 2 ชิ้น/)).toBeTruthy();
    expect(screen.getByPlaceholderText(/ค้นหาชื่ออุปกรณ์/).value).toBe('PO-OLD');
  });
});
