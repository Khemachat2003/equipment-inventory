// transferModal.test.jsx — TransferModal โครงใหม่ "เลือกปลายทางก่อน"
// ปุ่มปลายทาง 3 ปุ่ม ต้อง map เป็น action/status เดิมของ API แบบเป๊ะ (payload /api/transfer-asset ไม่เปลี่ยน)
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import TransferModal from '../components/TransferModal.jsx';

vi.mock('axios', () => ({
  default: {
    get: vi.fn(() => Promise.resolve({ data: [] })),
    post: vi.fn(() => Promise.resolve({ data: { success: true } })),
  },
}));
import axios from 'axios';

const SITES = [
  { siteId: 'SITE001', siteName: 'ฟาร์มเดิม', farmType: 'สัตว์ปีก' },
  { siteId: 'SITE002', siteName: 'ฟาร์มสุกร', farmType: 'สุกร' },
];

describe('TransferModal — เลือกปลายทางก่อน (รีดีไซน์ 2026-10)', () => {
  beforeEach(() => {
    axios.get.mockImplementation((url) =>
      Promise.resolve({ data: String(url).includes('farm-sites') ? SITES : [] }));
    axios.post.mockClear();
  });

  it('แสดงปลายทาง 3 ปุ่มใหญ่ + กล่องรายละเอียดเพิ่มถูกพับไว้ default', () => {
    render(<TransferModal open onClose={() => {}} serial="SN-1" current={{ siteName: 'คลัง', location: '', user: 'สมชาย' }} />);
    expect(screen.getByText('ไปฟาร์ม')).toBeTruthy();
    expect(screen.getByText('คืนคลังกลาง')).toBeTruthy();
    expect(screen.getByText('ส่งซ่อม')).toBeTruthy();
    // โค้ดเดิม "มีแต่ไม่แสดง": สถานะ/ประเภทการโอนย้าย ยังอยู่แต่พับไว้
    expect(screen.queryByText('สถานะใหม่')).toBeNull();
    expect(screen.queryByText('ประเภทการโอนย้าย (แก้เฉพาะกรณีพิเศษ)')).toBeNull();
  });

  it('กด "รายละเอียดเพิ่ม" แล้วเห็นสถานะ/ประเภทการโอนย้าย/หมายเหตุ (โค้ดเดิมครบ)', () => {
    render(<TransferModal open onClose={() => {}} serial="SN-1" current={{}} />);
    fireEvent.click(screen.getByRole('button', { name: /รายละเอียดเพิ่ม/ }));
    expect(screen.getByText('สถานะใหม่')).toBeTruthy();
    expect(screen.getByText('ประเภทการโอนย้าย (แก้เฉพาะกรณีพิเศษ)')).toBeTruthy();
    expect(screen.getByText('หมายเหตุ')).toBeTruthy();
  });

  it('ไปฟาร์ม (default) → payload action/status ตามระบบเดิม', async () => {
    const { container } = render(<TransferModal open onClose={() => {}} serial="SN-1" current={{ siteName: 'คลัง', location: 'Stock', user: '' }} />);
    // (จุดติดตั้ง เป็น <input list> ซึ่ง role ก็ combobox เหมือนกัน → คว้า select ตัวเดียวของฟอร์มตรง ๆ)
    await screen.findByText(/ฟาร์มเดิม \(สัตว์ปีก\)/); // รอ axios mock โหลดรายการฟาร์มเข้า dropdown ก่อน
    fireEvent.change(container.querySelector('select'), { target: { value: 'SITE001' } });
    fireEvent.click(screen.getByRole('button', { name: /ยืนยันโอนย้าย/ }));
    await waitFor(() => expect(axios.post).toHaveBeenCalled());
    const [url, body] = axios.post.mock.calls[0];
    expect(url).toBe('/api/transfer-asset');
    expect(body.action).toBe('ย้ายตำแหน่งอุปกรณ์');
    expect(body.status).toBe('ใช้งานได้');
    expect(body.siteName).toBe('ฟาร์มเดิม');
    expect(body.houseId).toBe('-');
  });

  it('คืนคลังกลาง → action คืนคลังสินค้า + สถานะ สำรอง + siteName Intranin + location Stock', async () => {
    render(<TransferModal open onClose={() => {}} serial="SN-2" current={{ siteName: 'ฟาร์มเดิม', location: 'ร้าน A', user: '' }} />);
    fireEvent.click(screen.getByText('คืนคลังกลาง'));
    fireEvent.click(screen.getByRole('button', { name: /ยืนยันโอนย้าย/ }));
    await waitFor(() => expect(axios.post).toHaveBeenCalled());
    const [, body] = axios.post.mock.calls[0];
    expect(body.action).toBe('คืนคลังสินค้า');
    expect(body.status).toBe('สำรอง');
    expect(body.siteName).toBe('Intranin');
    expect(body.location).toBe('Stock');
    expect(body.houseId).toBe('-');
  });

  it('ส่งซ่อม → action ส่งซ่อมภายนอก + สถานะ ส่งซ่อม', async () => {
    render(<TransferModal open onClose={() => {}} serial="SN-3" current={{ siteName: 'ฟาร์มเดิม', location: '', user: '' }} />);
    fireEvent.click(screen.getByText('ส่งซ่อม'));
    fireEvent.click(screen.getByRole('button', { name: /ยืนยันโอนย้าย/ }));
    await waitFor(() => expect(axios.post).toHaveBeenCalled());
    const [, body] = axios.post.mock.calls[0];
    expect(body.action).toBe('ส่งซ่อมภายนอก');
    expect(body.status).toBe('ส่งซ่อม');
  });
});