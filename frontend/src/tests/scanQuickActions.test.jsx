import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, waitFor, within } from '@testing-library/react';
import { MemoryRouter, useLocation } from 'react-router-dom';
import ScanPage from '../pages/ScanPage.jsx';

// Episode 3 — Focused QR Scanner: สแกน/พิมพ์ Serial → การ์ดรายละเอียด + Quick Actions 4 ปุ่มใต้ผลสแกน
// (ดูรายละเอียด / โอนย้ายอุปกรณ์ / จัดการชุดติดตั้ง / พิมพ์ฉลากซ้ำ) และส่วน PO Tracking ต้องไม่อยู่บนหน้านี้อีกต่อไป

vi.mock('axios', () => ({ default: { get: vi.fn(), post: vi.fn(), delete: vi.fn() } }));
vi.mock('../components/TransferModal.jsx', () => ({
  default: ({ serial, onSuccess, onClose }) => (
    <div role="dialog">
      <span>{serial}</span>
      <button onClick={onSuccess}>simulate transfer success</button>
      <button onClick={onClose}>close transfer</button>
    </div>
  ),
}));
import axios from 'axios';

function RouteLocation() {
  const location = useLocation();
  return <output data-testid="route">{location.pathname}{location.search}</output>;
}

const looseAsset = {
  assetId: 'A1', serialNumber: 'SN-1', name: 'Sensor', code: 'SEN',
  status: 'สำรอง', location: 'Stock', siteName: 'Intranin', user: '-',
};
const bundledAsset = {
  assetId: 'A2', serialNumber: 'SN-2', name: 'Relay', code: 'RELY',
  status: 'ใช้งานได้', location: 'Stock', siteName: 'Intranin', user: '-',
  bundleId: 'BDL-1', bundleName: 'Control cabinet',
};

function mockApi() {
  axios.get.mockImplementation((url) => {
    if (url === '/api/assets') return Promise.resolve({ data: [looseAsset, bundledAsset] });
    if (url === '/api/bundles') return Promise.resolve({ data: [{ bundleId: 'BDL-9', bundleName: 'Sensor cabinet', status: 'In Stock' }] });
    if (url.startsWith('/api/asset-history')) return Promise.resolve({ data: [] });
    return Promise.resolve({ data: [] });
  });
  axios.post.mockResolvedValue({ data: { success: true, message: 'เพิ่ม 1 อุปกรณ์เข้าชุดสำเร็จ' } });
  axios.delete.mockResolvedValue({ data: { success: true } });
}

async function findSerial(serial) {
  const manual = await screen.findByPlaceholderText(/เช่น SN-SEN/);
  fireEvent.change(manual, { target: { value: serial } });
  fireEvent.click(screen.getByRole('button', { name: 'ค้นหา' }));
  await screen.findByText('พบอุปกรณ์');
}

describe('ScanPage — Focused QR Scanner (Episode 3)', () => {
  beforeEach(() => {
    mockApi();
    vi.spyOn(window, 'open').mockImplementation(() => null);
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('ไม่แสดงส่วนติดตามอุปกรณ์ตาม PO อีกต่อไป (ย้ายไปหน้าแรกแล้ว)', async () => {
    render(<MemoryRouter><ScanPage /></MemoryRouter>);
    await screen.findByText('ค้นหาอุปกรณ์');
    expect(screen.queryByTestId('po-dispatch-heading')).toBeNull();
    expect(screen.queryByText(/ติดตามและจัดการอุปกรณ์ตาม PO/)).toBeNull();
  });

  it('พิมพ์ค้นหา Serial → แสดงการ์ดอุปกรณ์พร้อม Quick Actions ครบ 4 ปุ่ม', async () => {
    render(<MemoryRouter><ScanPage /></MemoryRouter>);
    await findSerial('SN-1');
    expect(screen.getByRole('button', { name: /ดูรายละเอียด/ })).toBeTruthy();
    expect(screen.getByRole('button', { name: /โอนย้ายอุปกรณ์/ })).toBeTruthy();
    expect(screen.getByRole('button', { name: /จัดการชุดติดตั้ง/ })).toBeTruthy();
    expect(screen.getByRole('button', { name: /พิมพ์ฉลากซ้ำ/ })).toBeTruthy();
  });

  it('กด "ดูรายละเอียด" → ไปหน้าทะเบียนพร้อมคำค้น ?q= และ "พิมพ์ฉลากซ้ำ" → เปิดหน้าฉลาก QR ของ Serial นั้น', async () => {
    render(<MemoryRouter><ScanPage /><RouteLocation /></MemoryRouter>);
    await findSerial('SN-1');
    fireEvent.click(screen.getByRole('button', { name: /ดูรายละเอียด/ }));
    await waitFor(() => expect(screen.getByTestId('route').textContent).toBe('/asset?q=SN-1'));
    fireEvent.click(screen.getByRole('button', { name: /พิมพ์ฉลากซ้ำ/ }));
    expect(window.open).toHaveBeenCalledWith('/qr?serial=SN-1', '_blank');
  });

  it('กด "โอนย้ายอุปกรณ์" → เปิดฟอร์มโอนย้ายของ Serial นั้นทันที', async () => {
    render(<MemoryRouter><ScanPage /></MemoryRouter>);
    await findSerial('SN-1');
    fireEvent.click(screen.getByRole('button', { name: /โอนย้ายอุปกรณ์/ }));
    expect(await screen.findByRole('dialog')).toBeTruthy();
    expect(within(screen.getByRole('dialog')).getByText('SN-1')).toBeTruthy();
  });

  it('กด "จัดการชุดติดตั้ง" บนอุปกรณ์ที่ยังไม่มีชุด → เลือกชุดแล้วย้ายเข้าได้ (bulk add)', async () => {
    render(<MemoryRouter><ScanPage /></MemoryRouter>);
    await findSerial('SN-1');
    fireEvent.click(screen.getByRole('button', { name: /จัดการชุดติดตั้ง/ }));
    const select = await screen.findByLabelText('เลือกชุดติดตั้ง');
    fireEvent.change(select, { target: { value: 'BDL-9' } });
    fireEvent.click(screen.getByRole('button', { name: /ย้ายเข้าชุดนี้/ }));
    await waitFor(() => expect(axios.post).toHaveBeenCalledWith('/api/bundles/BDL-9/assets/bulk', { assetIds: ['A1'] }));
    await waitFor(() => expect(screen.queryByLabelText('เลือกชุดติดตั้ง')).toBeNull());
  });

  it('อุปกรณ์ที่อยู่ในชุดแล้ว → แสดงชุดปัจจุบันและนำออกจากชุดได้', async () => {
    render(<MemoryRouter><ScanPage /></MemoryRouter>);
    await findSerial('SN-2');
    fireEvent.click(screen.getByRole('button', { name: /จัดการชุดติดตั้ง/ }));
    expect(await screen.findByText(/อยู่ในชุด:/)).toBeTruthy();
    fireEvent.click(screen.getByRole('button', { name: /นำออกจากชุด/ }));
    await waitFor(() => expect(axios.delete).toHaveBeenCalledWith('/api/bundles/BDL-1/assets/A2'));
    await waitFor(() => expect(screen.queryByText(/อยู่ในชุด:/)).toBeNull());
  });
});
