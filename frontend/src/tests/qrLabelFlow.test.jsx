// qrLabelFlow.test.jsx — Episode 1: หน้าพิมพ์ฉลาก QR (/qr) เรียงล่าสุดอยู่บน + ตัวกรอง PO/วันรับเข้า + ปุ่มด่วน "รับเข้า 2 วันล่าสุด"
//   + deep-link ?serials= (จากปุ่ม "พิมพ์ฉลาก QR ทันที" หลังเพิ่มอุปกรณ์ใน AddDeviceModal)
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen } from '@testing-library/react';
import { MemoryRouter } from 'react-router-dom';
import QrPage from '../pages/QrPage.jsx';
import AddDeviceModal from '../components/AddDeviceModal.jsx';
import { lastNDaysRange } from '../utils/labelQueue.js';

vi.mock('axios', () => ({ default: { get: vi.fn(), post: vi.fn() } }));
import axios from 'axios';

// วันอ้างอิงแบบเวลาไทย (Asia/Bangkok) — คำนวณจาก util เดียวกับหน้าจริง เพื่อไม่ติด timezone ของเครื่องที่รันเทส
const keyDaysAgo = (days) => lastNDaysRange(1, new Date(Date.now() - days * 24 * 60 * 60 * 1000)).to;
const at = (key) => `${key}T04:00:00+07:00`; // 04:00 ของวันนั้นตามเวลาไทย — อยู่ในวันเดียวกันแน่นอน
const ddmmyyyy = (key) => {
  const [y, m, d] = key.split('-');
  return `${d}${m}${y}`;
};

const TODAY = keyDaysAgo(0);
const YDA = keyDaysAgo(1);
const OLD = keyDaysAgo(5);

const SN_NEW2 = `SN-NEW2-${ddmmyyyy(TODAY)}-0002`;
const SN_NEW1 = `SN-NEW1-${ddmmyyyy(TODAY)}-0001`;
const SN_MID = `SN-MID-${ddmmyyyy(YDA)}-0001`;
const SN_OLD = `SN-OLD-${ddmmyyyy(OLD)}-0001`;

// API ส่งมาแบบ "เก่าอยู่บน" (ลำดับ sheet เดิม) — หน้าต้องเรียงใหม่เป็นล่าสุดอยู่บนเอง
const ASSETS = [
  { serialNumber: SN_OLD, name: 'อุปกรณ์เก่า', status: 'ใช้งานได้', poNumber: 'PO-OLD', batchId: 'INB-OLD', receivedAt: at(OLD) },
  { serialNumber: SN_MID, name: 'อุปกรณ์เมื่อวาน', status: 'ใช้งานได้', poNumber: '', batchId: 'INB-YDA', receivedAt: at(YDA) },
  { serialNumber: SN_NEW1, name: 'อุปกรณ์ใหม่ 1', status: 'ใช้งานได้', poNumber: 'PO-NEW', batchId: 'INB-NEW', receivedAt: at(TODAY) },
  { serialNumber: SN_NEW2, name: 'อุปกรณ์ใหม่ 2', status: 'สำรอง', poNumber: 'PO-NEW', batchId: 'INB-NEW', receivedAt: at(TODAY) },
];

beforeEach(() => {
  axios.get.mockImplementation((url) => {
    if (url === '/api/assets') return Promise.resolve({ data: ASSETS });
    if (url === '/api/part-catalog') return Promise.resolve({ data: [{ partNumber: 'SENSOR-1', partName: 'Sensor', totalQty: 3 }] });
    return Promise.resolve({ data: [] });
  });
  // POST สำหรับ AddDeviceModal (bulk-add-asset) — คืน Serial จริงที่ backend สร้าง
  axios.post.mockResolvedValue({
    data: { success: true, added: 1, serials: ['SN-SENSOR-1-09102026-0004'], batchId: 'INB-20261009-001', receivedAt: at(TODAY), poNumber: 'PO-NEW', supplier: 'Acme' },
  });
});

function renderPage(entry = '/qr') {
  return render(<MemoryRouter initialEntries={[entry]}><QrPage /></MemoryRouter>);
}

async function renderPageLoaded(entry = '/qr') {
  const utils = renderPage(entry);
  await screen.findAllByText(SN_OLD); // รอโหลดเสร็จ (Serial ปรากฏในตาราง + การ์ดมือถือ)
  return utils;
}

describe('QrPage — Smart Sorting & Grouping', () => {
  it('เรียง "ล่าสุดอยู่บนสุดเสมอ" แม้ API ส่งเก่าอยู่บน (Created Date DESC + เลขรัน)', async () => {
    const { container } = await renderPageLoaded();
    const text = container.textContent;
    expect(text.indexOf(SN_NEW2)).toBeLessThan(text.indexOf(SN_NEW1));
    expect(text.indexOf(SN_NEW1)).toBeLessThan(text.indexOf(SN_MID));
    expect(text.indexOf(SN_MID)).toBeLessThan(text.indexOf(SN_OLD));
  });

  it('แสดงคอลัมน์ รับเข้า/เลข PO + หัวกลุ่มตามวันรับเข้า', async () => {
    await renderPageLoaded();
    expect(screen.getAllByText('เลข PO').length).toBeGreaterThan(0);
    // หัวกลุ่ม "รับเข้า ..." มีป้ายวัน + จำนวนดวงต่อกลุ่ม (เดสก์ท็อป + มือถือ)
    expect(screen.getAllByText(/รับเข้า \d/).length).toBeGreaterThanOrEqual(2);
    expect(screen.getAllByText('PO-NEW').length).toBeGreaterThan(0);
  });

  it('กรองตามเลข PO ได้ทันที', async () => {
    await renderPageLoaded();
    fireEvent.change(screen.getByLabelText('กรองตามเลข PO'), { target: { value: 'PO-OLD' } });
    expect(screen.queryAllByText(SN_NEW1)).toHaveLength(0);
    expect(screen.queryAllByText(SN_NEW2)).toHaveLength(0);
    expect(screen.queryAllByText(SN_MID)).toHaveLength(0);
    expect(screen.getAllByText(SN_OLD).length).toBeGreaterThan(0);
  });

  it('ปุ่มด่วน "รับเข้า 2 วันล่าสุด" กรองเฉพาะวันนี้+เมื่อวาน และกดซ้ำเพื่อปิด', async () => {
    await renderPageLoaded();
    fireEvent.click(screen.getByRole('button', { name: /รับเข้า 2 วันล่าสุด/ }));
    expect(screen.queryAllByText(SN_OLD)).toHaveLength(0);
    expect(screen.getAllByText(SN_NEW1).length).toBeGreaterThan(0);
    expect(screen.getAllByText(SN_MID).length).toBeGreaterThan(0);
    // กดซ้ำ → กลับมาเห็นทั้งหมด
    fireEvent.click(screen.getByRole('button', { name: /รับเข้า 2 วันล่าสุด/ }));
    expect(screen.getAllByText(SN_OLD).length).toBeGreaterThan(0);
  });

  it('กรองช่วงวันรับเข้าด้วยช่องวันที่ไทย (พิมพ์ พ.ศ. ได้)', async () => {
    await renderPageLoaded();
    const [y, m, d] = TODAY.split('-');
    fireEvent.change(screen.getByLabelText('รับเข้าตั้งแต่'), { target: { value: `${d}/${m}/${Number(y) + 543}` } });
    expect(screen.queryAllByText(SN_MID)).toHaveLength(0);
    expect(screen.queryAllByText(SN_OLD)).toHaveLength(0);
    expect(screen.getAllByText(SN_NEW1).length).toBeGreaterThan(0);
    expect(screen.getAllByText(SN_NEW2).length).toBeGreaterThan(0);
  });
});

describe('QrPage — deep-link ?serials= (จากปุ่มพิมพ์ทันทีหลังเพิ่มอุปกรณ์)', () => {
  it('เลือก Serial ตามพารามิเตอร์ให้พร้อมพิมพ์เลย (แสดง "2 รายการ")', async () => {
    await renderPageLoaded(`/qr?serials=${encodeURIComponent(`${SN_NEW1},${SN_NEW2}`)}`);
    expect(screen.getByText('2 รายการ')).toBeTruthy();
  });
});

describe('AddDeviceModal — Direct-to-Print (พิมพ์ฉลากทันทีหลังบันทึกสำเร็จ)', () => {
  it('หน้าสรุปผลมีปุ่ม "พิมพ์ฉลาก QR ทันที" → เปิด /qr?serials=... ของ Serial ที่เพิ่งสร้าง', async () => {
    const openSpy = vi.spyOn(window, 'open').mockImplementation(() => null);
    const onDone = vi.fn();
    render(<AddDeviceModal open onClose={vi.fn()} onDone={onDone} />);

    fireEvent.change(screen.getByPlaceholderText(/SEN\.TEMP-RS485/), { target: { value: 'SENSOR-1' } });
    fireEvent.click(await screen.findByText('Sensor', { selector: 'span' }));
    fireEvent.click(screen.getByRole('button', { name: /เพิ่ม 1 ชิ้น/ }));

    const printBtn = await screen.findByRole('button', { name: /พิมพ์ฉลาก QR ทันที/ });
    fireEvent.click(printBtn);
    expect(openSpy).toHaveBeenCalledTimes(1);
    const [url, target] = openSpy.mock.calls[0];
    expect(url).toBe(`/qr?serials=${encodeURIComponent('SN-SENSOR-1-09102026-0004')}`);
    expect(target).toBe('_blank');

    // ปุ่ม "เสร็จ" = ปิดหน้าต่างทำงานต่อตามปกติ (onClose/onDone flow เดิม)
    fireEvent.click(screen.getByRole('button', { name: /^เสร็จ/ }));
    expect(onDone).toHaveBeenCalledTimes(1);
    openSpy.mockRestore();
  });
});
