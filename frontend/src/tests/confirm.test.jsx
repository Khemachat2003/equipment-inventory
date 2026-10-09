import { describe, it, expect } from 'vitest';
import { useState } from 'react';
import { render, screen, fireEvent } from '@testing-library/react';
import { showConfirm, ConfirmHost } from '../components/ui/Confirm.jsx';

// หน้าจำลอง: ปุ่มเรียก showConfirm + โชว์ผลลัพธ์ + ConfirmHost ติดตั้งครบ (เหมือนหน้าจริง)
function Probe({ opts }) {
  const [result, setResult] = useState('');
  async function ask() {
    const ok = await showConfirm(opts);
    setResult(ok ? 'ผลลัพธ์: ยืนยัน' : 'ผลลัพธ์: ยกเลิก');
  }
  return (
    <div>
      <button onClick={ask}>เรียกยืนยัน</button>
      <span data-testid="result">{result}</span>
      <ConfirmHost />
    </div>
  );
}

describe('ConfirmHost — กล่องยืนยันแทน window.confirm', () => {
  it('ยังไม่มีการเรียก showConfirm → ไม่ render อะไรเลย (null)', () => {
    const { container } = render(<ConfirmHost />);
    expect(container.firstChild).toBeNull();
  });

  it('แสดง title / message / confirmLabel ที่ส่งเข้า', async () => {
    render(<Probe opts={{ title: 'นำ SN-X ออกจากชุด?', message: 'ข้อความอธิบาย', confirmLabel: 'นำออกจากชุด' }} />);
    fireEvent.click(screen.getByText('เรียกยืนยัน'));
    expect(await screen.findByText('นำ SN-X ออกจากชุด?')).toBeTruthy();
    expect(screen.getByText('ข้อความอธิบาย')).toBeTruthy();
    expect(screen.getByText('นำออกจากชุด')).toBeTruthy();
  });

  it('ไม่ส่ง label → ใช้ default (ยืนยัน / ยกเลิก)', async () => {
    render(<Probe opts={{ title: 'ยืนยันอะไรสักอย่าง' }} />);
    fireEvent.click(screen.getByText('เรียกยืนยัน'));
    await screen.findByText('ยืนยันอะไรสักอย่าง');
    expect(screen.getByText('ยืนยัน')).toBeTruthy();
    expect(screen.getByText('ยกเลิก')).toBeTruthy();
  });

  it('กดปุ่มยืนยัน → Promise resolve true และกล่องปิด', async () => {
    render(<Probe opts={{ title: 'คืนชุดนี้กลับเข้าคลัง?', confirmLabel: 'คืนเข้าคลัง' }} />);
    fireEvent.click(screen.getByText('เรียกยืนยัน'));
    fireEvent.click(await screen.findByText('คืนเข้าคลัง'));
    expect(await screen.findByText('ผลลัพธ์: ยืนยัน')).toBeTruthy();
    expect(screen.queryByText('คืนชุดนี้กลับเข้าคลัง?')).toBeNull(); // กล่องปิดแล้ว
  });

  it('กดยกเลิก → resolve false', async () => {
    render(<Probe opts={{ title: 'ยืนยันลบผู้ใช้ test01?', confirmLabel: 'ลบผู้ใช้' }} />);
    fireEvent.click(screen.getByText('เรียกยืนยัน'));
    fireEvent.click(await screen.findByText('ยกเลิก'));
    expect(await screen.findByText('ผลลัพธ์: ยกเลิก')).toBeTruthy();
  });

  it('กด Esc → resolve false', async () => {
    render(<Probe opts={{ title: 'ยืนยันลบผู้ใช้ test02?' }} />);
    fireEvent.click(screen.getByText('เรียกยืนยัน'));
    await screen.findByText('ยืนยันลบผู้ใช้ test02?');
    fireEvent.keyDown(window, { key: 'Escape' });
    expect(await screen.findByText('ผลลัพธ์: ยกเลิก')).toBeTruthy();
  });

  it('คลิกพื้นหลัง (นอกการ์ด) → resolve false', async () => {
    render(<Probe opts={{ title: 'ยืนยันลบผู้ใช้ test03?' }} />);
    fireEvent.click(screen.getByText('เรียกยืนยัน'));
    await screen.findByText('ยืนยันลบผู้ใช้ test03?');
    fireEvent.click(screen.getByRole('dialog')); // target = พื้นหลังพอดี
    expect(await screen.findByText('ผลลัพธ์: ยกเลิก')).toBeTruthy();
  });
});