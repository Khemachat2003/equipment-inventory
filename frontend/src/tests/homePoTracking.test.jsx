import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, waitFor, within } from '@testing-library/react';
import { MemoryRouter } from 'react-router-dom';
import Home from '../pages/Home.jsx';

// Clean Home Layout + Transfer Center Modal:
//   • การ์ด "ติดตามและจัดการอุปกรณ์ตาม PO" ถูกถอดออกจากหน้าแรกถาวร — หน้าแรกสะอาด ไม่ดันเนื้อหา/ไม่ต้อง scroll เกินจำเป็น
//   • ปุ่มวงกลม "ย้ายอุปกรณ์" → เปิด Transfer Center Modal ป๊อปอัปกลางหน้าจอทันที (แท็บ 1: ตาม PO · แท็บ 2: อุปกรณ์ทั่วไป)
//   • ข้อมูลอุปกรณ์/ชุดใช้ของหน้าแรก (preload) — ไม่ดึง /api/assets ซ้ำ
//   • โอนย้ายสำเร็จ → ปิด Modal + toast แจ้งเตือน + หน้าแรกรีเฟรช (bump dataVersion)

vi.mock('axios', () => ({ default: { get: vi.fn(), post: vi.fn(), delete: vi.fn() } }));
vi.mock('../components/TransferModal.jsx', () => ({
  default: ({ serial, onSuccess, onClose }) => (
    <div role="dialog" data-testid="transfer-form">
      <span>{serial}</span>
      <button onClick={onSuccess}>simulate transfer success</button>
      <button onClick={onClose}>close transfer</button>
    </div>
  ),
}));
import axios from 'axios';

const inPOAsset = { assetId: 'A1', serialNumber: 'SN-1', name: 'Sensor', location: 'Stock', siteName: 'Intranin', poNumber: 'PO-1' };
const bundledAsset = { assetId: 'A2', serialNumber: 'SN-2', name: 'Relay', location: 'Stock', siteName: 'Intranin', poNumber: 'PO-1', bundleId: 'BDL-9', bundleName: 'Sensor cabinet' };
// อุปกรณ์เดี่ยวไม่อยู่ใน PO ใด — ต้องโอนย้ายได้จากแท็บ "โอนย้ายอุปกรณ์ทั่วไป (นอก PO)"
const looseAsset = { assetId: 'A5', serialNumber: 'SN-5', name: 'Loose sensor', location: 'Stock', siteName: 'Intranin', poNumber: '' };

function mockApi() {
  axios.get.mockImplementation((url) => {
    if (url === '/api/inbound-pos') return Promise.resolve({ data: [{ key: 'po:PO-1', poNumber: 'PO-1', status: 'active', stockCount: 1, assetCount: 2 }] });
    if (url === '/api/assets') return Promise.resolve({ data: [inPOAsset, bundledAsset, looseAsset] });
    if (url === '/api/bundles') return Promise.resolve({ data: [{ bundleId: 'BDL-9', bundleName: 'Sensor cabinet', status: 'In Stock' }] });
    if (url === '/api/dashboard-full') return Promise.resolve({ data: { farmSites: [], statusCount: {}, assets: [], totalFarmAssets: 0, totalStockAssets: 0 } });
    return Promise.resolve({ data: [] });
  });
}

const assetCallCount = () => axios.get.mock.calls.filter(([url]) => url === '/api/assets').length;

async function openTransferCenter() {
  fireEvent.click(screen.getByRole('button', { name: 'ย้ายอุปกรณ์' }));
  return screen.findByTestId('transfer-center');
}

describe('Home — Clean Home Layout + Transfer Center Modal', () => {
  beforeEach(() => {
    sessionStorage.clear();
    mockApi();
  });

  it('หน้าแรกสะอาด: ไม่มีการ์ด "ติดตามและจัดการอุปกรณ์ตาม PO" ติดถาวร — กดปุ่ม "ย้ายอุปกรณ์" เปิด Transfer Center พร้อม 2 แท็บ (ใช้ preload ไม่ดึง /api/assets ซ้ำ)', async () => {
    render(<MemoryRouter><Home /></MemoryRouter>);
    await screen.findByRole('button', { name: 'ย้ายอุปกรณ์' });
    // ก่อนกด: ไม่มีการ์ด PO บนหน้าแรก (Clean Layout) และยังไม่เปิด Modal
    expect(screen.queryByText(/ติดตามและจัดการอุปกรณ์ตาม PO/)).toBeNull();
    expect(screen.queryByTestId('transfer-center')).toBeNull();
    const callsBefore = assetCallCount();

    const modal = await openTransferCenter();
    // แท็บครบ 2 แท็บตาม spec
    expect(within(modal).getByRole('tab', { name: 'โอนย้ายตาม PO' })).toBeTruthy();
    expect(within(modal).getByRole('tab', { name: 'โอนย้ายอุปกรณ์ทั่วไป (นอก PO)' })).toBeTruthy();
    await waitFor(() => expect(axios.get.mock.calls.some(([url]) => url === '/api/inbound-pos')).toBe(true));
    await within(modal).findByText(/PO-1/); // dropdown PO โหลดแล้ว
    expect(assetCallCount()).toBe(callsBefore); // preload ทำงาน — ไม่ดึงอุปกรณ์ซ้ำ
  });

  it('แท็บ 2: โอนย้ายอุปกรณ์ทั่วไปที่ไม่อยู่ใน PO — ค้นหา → โอนย้าย → ปิด Modal + toast แจ้งเตือน + หน้าแรกรีเฟรช', async () => {
    render(<MemoryRouter><Home /></MemoryRouter>);
    await screen.findByRole('button', { name: 'ย้ายอุปกรณ์' });
    await openTransferCenter();
    fireEvent.click(screen.getByRole('tab', { name: 'โอนย้ายอุปกรณ์ทั่วไป (นอก PO)' }));
    await screen.findByTestId('general-transfer-section');

    const transferCallsBefore = assetCallCount();
    fireEvent.click(await screen.findByRole('button', { name: 'Transfer SN-5' }));
    const form = await screen.findByTestId('transfer-form');
    expect(within(form).getByText('SN-5')).toBeTruthy();
    fireEvent.click(within(form).getByRole('button', { name: 'simulate transfer success' }));

    await waitFor(() => expect(screen.queryByTestId('transfer-center')).toBeNull()); // ปิด Modal เมื่อโอนสำเร็จ
    expect(await screen.findByText('ย้ายอุปกรณ์สำเร็จแล้ว')).toBeTruthy(); // toast แจ้งเตือน (ToastHost ระดับหน้า)
    await waitFor(() => expect(assetCallCount()).toBeGreaterThan(transferCallsBefore)); // onTransferSuccess → หน้าแรกรีเฟรช
  });

  it('รายการที่อยู่ใน PO ที่เลือกอยู่จะไม่โผล่ซ้ำในแท็บอุปกรณ์ทั่วไป', async () => {
    render(<MemoryRouter><Home /></MemoryRouter>);
    await screen.findByRole('button', { name: 'ย้ายอุปกรณ์' });
    await openTransferCenter();
    await screen.findByTestId('panel-po'); // แท็บ 1 (ตาม PO) เป็นแท็บเริ่มต้น
    fireEvent.click(screen.getByRole('tab', { name: 'โอนย้ายอุปกรณ์ทั่วไป (นอก PO)' }));
    const generalBox = await screen.findByTestId('general-transfer-section');

    expect(within(generalBox).queryByRole('button', { name: 'Transfer SN-1' })).toBeNull(); // SN-1 อยู่ใน PO-1 — จัดการในแท็บ PO
    expect(within(generalBox).getByRole('button', { name: 'Transfer SN-5' })).toBeTruthy(); // SN-5 ไม่อยู่ใน PO — โอนย้ายได้ที่นี่
  });
});