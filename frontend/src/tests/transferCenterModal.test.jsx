import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, waitFor, within } from '@testing-library/react';
import { MemoryRouter, useLocation } from 'react-router-dom';
import TransferCenterModal from '../components/TransferCenterModal.jsx';
import { DISPATCH_QUEUE_KEY } from '../utils/dispatchQueue.js';

vi.mock('axios', () => ({ default: { get: vi.fn(), post: vi.fn() } }));
vi.mock('../components/TransferModal.jsx', () => ({ default: ({ serial, onSuccess, onClose }) => <div role="dialog" data-testid="transfer-form"><span>{serial}</span><button onClick={onSuccess}>simulate transfer success</button><button onClick={onClose}>close transfer</button></div> }));
import axios from 'axios';

function RouteLocation() {
  const location = useLocation();
  return <output data-testid="route">{location.pathname}{location.search}</output>;
}

async function openCenter() {
  const onClose = vi.fn();
  render(<MemoryRouter><TransferCenterModal open onClose={onClose} /><RouteLocation /></MemoryRouter>);
  await screen.findByText(/SN-1/); // รายการอุปกรณ์ของ PO-1 โหลดแล้ว
  return { onClose };
}

describe('Transfer Center Modal — แท็บโอนย้ายตาม PO', () => {
  beforeEach(() => {
    sessionStorage.clear();
    axios.get.mockImplementation((url) => {
      if (url.startsWith('/api/inbound-pos')) return Promise.resolve({ data: [{ key: 'po:PO-1', poNumber: 'PO-1', status: 'active', stockCount: 2, assetCount: 3 }, { key: 'po:PO-OLD', poNumber: 'PO-OLD', status: 'closed', stockCount: 0, assetCount: 1 }] });
      if (url === '/api/assets') return Promise.resolve({ data: [
        { assetId: 'A1', serialNumber: 'SN-1', name: 'Sensor', location: 'Stock', siteName: 'Intranin', poNumber: 'PO-1', batchId: 'INB-1' },
        { assetId: 'A2', serialNumber: 'SN-2', name: 'Relay', location: 'Stock', siteName: 'Intranin', poNumber: 'PO-1', batchId: 'INB-1', bundleId: 'BDL-9', bundleName: 'Sensor cabinet' },
        { assetId: 'A3', serialNumber: 'SN-3', name: 'Sensor', location: 'Farm A', siteName: 'Farm A', poNumber: 'PO-1', batchId: 'INB-1' },
        { assetId: 'A4', serialNumber: 'SN-4', name: 'Controller', location: 'Farm B', siteName: 'Farm B', poNumber: 'PO-OLD', batchId: 'INB-OLD' },
      ] });
      return Promise.resolve({ data: [] });
    });
    axios.post.mockResolvedValue({ data: { success: true } });
  });

  it('เปิดมาเจอ 2 แท็บ + แท็บ PO เริ่มต้น: สถานะ Serial ครบ (ในคลัง / ใน Bundle / ย้ายแล้ว)', async () => {
    await openCenter();
    expect(screen.getByRole('tab', { name: 'โอนย้ายตาม PO' })).toBeTruthy();
    expect(screen.getByRole('tab', { name: 'โอนย้ายอุปกรณ์ทั่วไป (นอก PO)' })).toBeTruthy();
    expect(screen.getAllByText(/Sensor cabinet/).length).toBeGreaterThan(0);
    expect(screen.getByText(/Farm A/)).toBeTruthy();
    await waitFor(() => expect(screen.getByTestId('dispatch-create-bundle').textContent).toContain('(1)')); // auto-select ของในคลัง
  });

  it('สร้าง Bundle ใหม่: stage เฉพาะ Serial ในคลังที่ยังไม่อยู่ Bundle → คิว + เด้งไป /bundle?dispatch=1', async () => {
    await openCenter();
    // รอ auto-select อุปกรณ์ในคลังก่อนกด — ปุ่ม enable เมื่อมีรายการเลือก (1) (กัน race กดก่อนเอฟเฟกต์เลือกอัตโนมัติแล้วคลิกไม่ผล)
    await waitFor(() => expect(screen.getByTestId('dispatch-create-bundle').textContent).toContain('(1)'));
    fireEvent.click(screen.getByTestId('dispatch-create-bundle'));
    await waitFor(() => expect(screen.getByTestId('route').textContent).toBe('/bundle?dispatch=1'));
    expect(JSON.parse(sessionStorage.getItem(DISPATCH_QUEUE_KEY))).toMatchObject({ poNumber: 'PO-1', items: [{ assetId: 'A1', serialNumber: 'SN-1' }] });
  });

  it('PO ปิดแล้ว: แสดงสถานะล่าสุดเพื่อติดตาม แต่ไม่มีปุ่มย้ายซ้ำ', async () => {
    await openCenter();
    fireEvent.change(screen.getAllByRole('combobox')[0], { target: { value: 'PO-OLD' } });
    expect(await screen.findByText(/SN-4/)).toBeTruthy();
    expect(screen.getByText(/Farm B/)).toBeTruthy();
    expect(screen.queryByRole('button', { name: 'Transfer SN-4' })).toBeNull();
  });

  it('ย้ายรายตัว: เปิดฟอร์มโอนย้าย → สำเร็จ → ปิด Transfer Center + แจ้ง onTransferSuccess', async () => {
    const { onClose } = await openCenter();
    fireEvent.click(screen.getByRole('button', { name: 'Transfer SN-1' }));
    const form = await screen.findByTestId('transfer-form');
    expect(within(form).getByText('SN-1')).toBeTruthy();
    fireEvent.click(within(form).getByRole('button', { name: 'simulate transfer success' }));
    await waitFor(() => expect(onClose).toHaveBeenCalledTimes(1)); // ปิด Modal ทันทีเมื่อโอนสำเร็จ
  });

  it('เพิ่ม Serial ที่เลือกเข้า Bundle เดิม (bulk add) — Modal เปิดค้างไว้ทำงานต่อได้', async () => {
    const { onClose } = await openCenter();
    fireEvent.change(screen.getByRole('combobox', { name: 'Existing Bundle' }), { target: { value: 'BDL-9' } });
    fireEvent.click(screen.getByTestId('dispatch-add-existing'));
    await waitFor(() => expect(axios.post).toHaveBeenCalledWith('/api/bundles/BDL-9/assets/bulk', { assetIds: ['A1'] }));
    await waitFor(() => expect(screen.getAllByRole('status')[0].textContent).toContain('Sensor cabinet'));
    expect(onClose).not.toHaveBeenCalled(); // จัดชุดไม่ใช่การโอนย้าย — ไม่ปิด Modal
  });
});