// tests/fastTrackIntake.test.jsx — Episode 2: Fast-Track Multi-Step Guided Flow ใน AddDeviceModal
// Step 1 รับเข้า → Step 2 จัดชุด (ข้ามได้ / เลือกชุดเดิม / สร้างชุดใหม่ทันที + รันเลขรหัสอัตโนมัติ)
//        → Step 3 โอนย้าย (ข้ามได้ + ปุ่มด่วนเพิ่มฟาร์ม/โรงเรือนใหม่ใน Modal) → Step 4 พิมพ์ฉลาก QR
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import AddDeviceModal from '../components/AddDeviceModal.jsx';

vi.mock('axios', () => ({ default: { get: vi.fn(), post: vi.fn() } }));
import axios from 'axios';

const SITES = [
  { siteId: 'FS001', siteName: 'ฟาร์มต้นแบบ', farmType: 'สัตว์ปีก' },
  { siteId: 'FS002', siteName: 'ฟาร์มสุกร', farmType: 'สุกร' },
];
const HOUSES = [{ houseId: 'FS001-H01', siteId: 'FS001', houseName: 'โรงเรือน 1' }];
const BUNDLES = [{ bundleId: 'BDL-018', bundleName: 'ชุดวัดอากาศ', description: '', status: 'In Stock', location: 'Stock', assetIds: ['0001'] }];
const NEW_SN = 'SN-SENSOR-1-09102026-0004';

beforeEach(() => {
  axios.get.mockImplementation((url) => {
    if (url === '/api/farm-sites') return Promise.resolve({ data: SITES });
    if (url === '/api/part-catalog') return Promise.resolve({ data: [{ partNumber: 'SENSOR-1', partName: 'Sensor', totalQty: 3 }] });
    if (url === '/api/bundles') return Promise.resolve({ data: BUNDLES });
    if (url.startsWith('/api/farm-houses/')) return Promise.resolve({ data: HOUSES });
    return Promise.resolve({ data: [] });
  });
  axios.post.mockImplementation((url) => {
    if (url === '/api/bulk-add-asset') return Promise.resolve({ data: { success: true, added: 1, serials: [NEW_SN], assetIds: ['A099'], batchId: 'INB-1', receivedAt: '2026-10-15T03:00:00.000Z', poNumber: 'PO-1', supplier: '' } });
    if (url === '/api/add-farm-site') return Promise.resolve({ data: { success: true, site: { siteId: 'FSNEW', siteName: 'ฟาร์มใหม่ริมถนน', farmType: 'สัตว์ปีก' } } });
    return Promise.resolve({ data: { success: true, message: 'ok' } });
  });
});

async function receiveOne(onDone = vi.fn()) {
  render(<AddDeviceModal open onClose={vi.fn()} onDone={onDone} />);
  fireEvent.change(screen.getByPlaceholderText(/SEN\.TEMP-RS485/), { target: { value: 'SENSOR-1' } });
  fireEvent.click(await screen.findByText('Sensor', { selector: 'span' }));
  fireEvent.click(screen.getByRole('button', { name: /เพิ่ม 1 ชิ้น/ }));
  await screen.findByText(/เพิ่ม 1 ชิ้นสำเร็จ/);
  // ปุ่มพิมพ์ฉลาก (Episode 1) ยังอยู่บนหน้าสรุปผลทันที — พร้อม stepper ของ Fast-Track
  expect(screen.getByRole('button', { name: /พิมพ์ฉลาก QR ทันที/ })).toBeTruthy();
  expect(screen.getByText(/② จัดชุดติดตั้ง/)).toBeTruthy();
}

describe('Fast-Track Intake Flow — Step 2 (จัดชุดติดตั้ง, ข้ามได้)', () => {
  it('ข้าม Step 2 → Step 3 → ข้าม Step 3 → Step 4 พิมพ์ฉลาก + เสร็จ (onDone)', async () => {
    const openSpy = vi.spyOn(window, 'open').mockImplementation(() => null);
    const onDone = vi.fn();
    await receiveOne(onDone);
    fireEvent.click(screen.getByRole('button', { name: /ข้าม — ไม่จัดชุดตอนนี้/ }));
    expect(await screen.findByText(/③ โอนย้ายไปยังฟาร์ม/)).toBeTruthy();
    fireEvent.click(screen.getByRole('button', { name: /ข้าม — เก็บที่เดิมไว้ก่อน/ }));
    expect(await screen.findByText(/④ พิมพ์ฉลาก QR/)).toBeTruthy();
    fireEvent.click(screen.getByRole('button', { name: /เปิดหน้าพิมพ์ฉลาก QR/ }));
    expect(openSpy).toHaveBeenCalledWith(`/qr?serials=${encodeURIComponent(NEW_SN)}`, '_blank', 'noopener');
    fireEvent.click(screen.getByRole('button', { name: /^เสร็จ/ }));
    expect(onDone).toHaveBeenCalledTimes(1);
    openSpy.mockRestore();
  });

  it('เลือกชุดที่มีอยู่ → เพิ่มเข้าชุดผ่าน bulk endpoint (ส่ง Asset IDs ที่เพิ่งรับเข้า) แล้วไป Step 3', async () => {
    await receiveOne();
    await screen.findByRole('option', { name: /BDL-018/ });
    fireEvent.change(screen.getByLabelText('เลือกชุดติดตั้ง'), { target: { value: 'BDL-018' } });
    fireEvent.click(screen.getByRole('button', { name: /เพิ่ม 1 ชิ้นเข้าชุดนี้/ }));
    await waitFor(() => expect(axios.post).toHaveBeenCalledWith('/api/bundles/BDL-018/assets/bulk', { assetIds: ['A099'] }));
    expect(await screen.findByText(/③ โอนย้ายไปยังฟาร์ม/)).toBeTruthy();
    expect(await screen.findByText(/ชุดวัดอากาศ \(BDL-018\)/)).toBeTruthy();
  });

  it('สร้างชุดใหม่ทันที — รหัสชุด Autofill รันต่อจากรหัสล่าสุด (BDL-018 → BDL-019) และเพิ่มอุปกรณ์เข้าชุดใหม่เลย', async () => {
    await receiveOne();
    fireEvent.click(screen.getByRole('button', { name: /สร้างชุดใหม่ทันที/ }));
    const codeInput = await screen.findByPlaceholderText('เช่น BDL-001');
    expect(codeInput.value).toBe('BDL-019');
    fireEvent.change(screen.getByPlaceholderText('เช่น ชุดตู้ควบคุมฟาร์ม 1'), { target: { value: 'ชุดใหม่จากหน้ารับเข้า' } });
    fireEvent.click(screen.getByRole('button', { name: /บันทึก/ }));
    await waitFor(() => expect(axios.post).toHaveBeenCalledWith('/api/bundles', expect.objectContaining({ bundleId: 'BDL-019', bundleName: 'ชุดใหม่จากหน้ารับเข้า' })));
    await waitFor(() => expect(axios.post).toHaveBeenCalledWith('/api/bundles/BDL-019/assets/bulk', { assetIds: ['A099'] }));
    expect(await screen.findByText(/③ โอนย้ายไปยังฟาร์ม/)).toBeTruthy();
    expect(await screen.findByText(/สร้างชุด BDL-019/)).toBeTruthy();
  });
});

describe('Fast-Track Intake Flow — Step 3 (โอนย้าย, ข้ามได้ + เพิ่มฟาร์ม/โรงเรือนใหม่ได้ทันที)', () => {
  async function goStep3() {
    await receiveOne();
    fireEvent.click(screen.getByRole('button', { name: /ข้าม — ไม่จัดชุดตอนนี้/ }));
    await screen.findByText(/③ โอนย้ายไปยังฟาร์ม/);
  }

  it('เลือกฟาร์ม/โรงเรือนปลายทาง → ยืนยันโอนย้ายรายชิ้นผ่าน /api/transfer-asset แล้วไป Step 4', async () => {
    await goStep3();
    fireEvent.click(screen.getByRole('button', { name: /ฟาร์ม\/โรงเรือน — ส่งปลายทางเลย/ }));
    fireEvent.change(screen.getByLabelText('ฟาร์มปลายทาง'), { target: { value: 'FS001' } });
    await screen.findByRole('option', { name: /FS001-H01/ });
    fireEvent.change(screen.getByLabelText('โรงเรือนปลายทาง'), { target: { value: 'FS001-H01' } });
    fireEvent.click(screen.getByRole('button', { name: /ยืนยันโอนย้าย \(1 ชิ้น\)/ }));
    await waitFor(() => expect(axios.post).toHaveBeenCalledWith('/api/transfer-asset', expect.objectContaining({
      serialNumber: NEW_SN,
      action: 'ย้ายตำแหน่งอุปกรณ์',
      siteName: 'ฟาร์มต้นแบบ',
      houseId: 'FS001-H01',
      houseName: 'โรงเรือน 1',
    })));
    expect(await screen.findByText(/④ พิมพ์ฉลาก QR/)).toBeTruthy();
    expect(await screen.findByText(/ย้าย 1 ชิ้นไป ฟาร์มต้นแบบ › โรงเรือน 1 แล้ว/)).toBeTruthy();
  });

  it('ปุ่มด่วน "เพิ่มฟาร์ม/โรงเรือนใหม่" ซ้อนใน Modal — สร้างฟาร์มได้ทันทีแล้วเลือกให้เลย (ไม่ต้องออกจากหน้าเพิ่มอุปกรณ์)', async () => {
    await goStep3();
    fireEvent.click(screen.getByRole('button', { name: /ฟาร์ม\/โรงเรือน — ส่งปลายทางเลย/ }));
    fireEvent.click(screen.getByRole('button', { name: /เพิ่มฟาร์ม\/โรงเรือนใหม่/ }));
    expect(await screen.findByText('เพิ่มฟาร์มใหม่')).toBeTruthy();
    fireEvent.change(screen.getByPlaceholderText(/ชื่อฟาร์ม \*/), { target: { value: 'ฟาร์มใหม่ริมถนน' } });
    fireEvent.change(screen.getByPlaceholderText(/รหัสฟาร์ม \*/), { target: { value: 'FSNEW' } });
    fireEvent.click(screen.getByRole('button', { name: /บันทึกฟาร์ม/ }));
    await waitFor(() => expect(axios.post).toHaveBeenCalledWith('/api/add-farm-site', expect.objectContaining({ siteId: 'FSNEW', siteName: 'ฟาร์มใหม่ริมถนน' })));
    expect(screen.getByLabelText('ฟาร์มปลายทาง').value).toBe('FSNEW');
  });
});
