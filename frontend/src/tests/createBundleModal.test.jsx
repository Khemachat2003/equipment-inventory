// tests/createBundleModal.test.jsx — Episode 2: Auto-Increment Bundle Code
// เปิดฟอร์มสร้างชุดติดตั้งใหม่ → ระบบตรวจรหัสชุด "ล่าสุด" ในระบบ (เช่น BDL-018)
// แล้ว Autofill รหัสถัดไปให้ทันที (BDL-019) — ผู้ใช้แก้ไขได้ + กันกรอกรหัสซ้ำ
import { describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen } from '@testing-library/react';
import CreateBundleModal from '../components/CreateBundleModal.jsx';

const EXISTING = Array.from({ length: 18 }, (_, i) => `BDL-${String(i + 1).padStart(3, '0')}`);

describe('CreateBundleModal — Auto-Increment Bundle Code', () => {
  it('Autofill รหัสถัดไปจากรหัสชุดล่าสุดในระบบ (BDL-018 → BDL-019)', () => {
    render(<CreateBundleModal onClose={vi.fn()} onSubmit={vi.fn()} existingIds={EXISTING} />);
    expect(screen.getByPlaceholderText('เช่น BDL-001').value).toBe('BDL-019');
  });

  it('ยังไม่มีชุดในระบบเลย → เริ่มที่ BDL-001', () => {
    render(<CreateBundleModal onClose={vi.fn()} onSubmit={vi.fn()} existingIds={[]} />);
    expect(screen.getByPlaceholderText('เช่น BDL-001').value).toBe('BDL-001');
  });

  it('ผู้ใช้แก้ไขรหัสได้ — onSubmit ได้รับ payload ตามที่กรอก', () => {
    const onSubmit = vi.fn();
    render(<CreateBundleModal onClose={vi.fn()} onSubmit={onSubmit} existingIds={EXISTING} />);
    fireEvent.change(screen.getByPlaceholderText('เช่น BDL-001'), { target: { value: 'BDL-099' } });
    fireEvent.change(screen.getByPlaceholderText('เช่น ชุดตู้ควบคุมฟาร์ม 1'), { target: { value: 'ชุดทดสอบ' } });
    fireEvent.click(screen.getByRole('button', { name: /บันทึก/ }));
    expect(onSubmit).toHaveBeenCalledWith({ bundleId: 'BDL-099', bundleName: 'ชุดทดสอบ', description: '', status: 'In Stock' });
  });

  it('กันรหัสซ้ำ: พิมพ์รหัสที่มีอยู่แล้ว → เตือนและปุ่มบันทึกถูกปิด', () => {
    render(<CreateBundleModal onClose={vi.fn()} onSubmit={vi.fn()} existingIds={EXISTING} />);
    fireEvent.change(screen.getByPlaceholderText('เช่น BDL-001'), { target: { value: 'BDL-018' } });
    expect(screen.getByText(/รหัสนี้ถูกใช้แล้ว/)).toBeTruthy();
    expect(screen.getByRole('button', { name: /บันทึก/ }).disabled).toBe(true);
  });
});
