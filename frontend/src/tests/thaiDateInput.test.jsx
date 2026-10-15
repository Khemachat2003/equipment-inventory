// thaiDateInput.test.jsx — ช่องวันที่รูปแบบไทย วว/ดด/ปปปป (B2: ไม่โชว์ mm/dd/yyyy ของ browser)
import { describe, expect, it, vi } from 'vitest';
import { useState } from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import ThaiDateInput, { thaiDisplay, thaiToIso } from '../components/ui/ThaiDateInput.jsx';

describe('thaiToIso — ข้อความไทย วว/ดด/ปปปป → ISO', () => {
  it('ปี พ.ศ. (≥2400) → ลบ 543 เป็น ค.ศ.', () => {
    expect(thaiToIso('01/10/2568')).toBe('2025-10-01');
  });

  it('ปี ค.ศ. (1xxx–2xxx) → ใช้ตามที่พิมพ์', () => {
    expect(thaiToIso('01/10/2025')).toBe('2025-10-01');
  });

  it('พิมพ์ไม่ครบ 8 หลัก → "" (ยังไม่เปลี่ยนตัวกรอง)', () => {
    expect(thaiToIso('')).toBe('');
    expect(thaiToIso('01/10')).toBe('');
    expect(thaiToIso('0110202')).toBe('');
  });

  it('วันที่/เดือนไม่ถูกต้อง → null', () => {
    expect(thaiToIso('32/13/2025')).toBeNull();
    expect(thaiToIso('00/01/2025')).toBeNull();
    expect(thaiToIso('01/00/2025')).toBeNull();
  });
});

describe('thaiDisplay — ISO → รูปแบบไทย (ปี พ.ศ.)', () => {
  it('แปลงและบวก 543 ให้ปี', () => {
    expect(thaiDisplay('2025-10-01')).toBe('01/10/2568');
  });

  it('ค่าว่าง / รูปแบบผิด → ""', () => {
    expect(thaiDisplay('')).toBe('');
    expect(thaiDisplay('nope')).toBe('');
  });
});

describe('ThaiDateInput — ส่วนประกอบ', () => {
  it('แสดงค่าเดิมเป็นรูปแบบไทย + placeholder วว/ดด/ปปปป + ปุ่มปฏิทิน/ล้างวันที่', () => {
    render(<ThaiDateInput label="รับเข้าตั้งแต่" value="2025-10-01" onChange={() => {}} />);
    expect(screen.getByLabelText('รับเข้าตั้งแต่').value).toBe('01/10/2568');
    expect(screen.getByPlaceholderText('วว/ดด/ปปปป')).toBeTruthy();
    expect(screen.getByRole('button', { name: 'เปิดปฏิทินเลือกวันที่ รับเข้าตั้งแต่' })).toBeTruthy();
    expect(screen.getByRole('button', { name: 'ล้างวันที่ รับเข้าตั้งแต่' })).toBeTruthy();
  });

  it('พิมพ์รูปแบบไทย → onChange ได้ ISO (พ.ศ. → ค.ศ.) และปุ่มล้างคืนค่าว่าง', () => {
    // Harness มี state จริง — ปุ่ม "ล้างวันที่" จะโผล่ก็ต่อเมื่อ value ไม่ว่าง (แสดงผลตาม prop)
    const calls = [];
    function Harness() {
      const [v, setV] = useState('');
      return <ThaiDateInput label="ถึง" ariaLabel="รับเข้าถึง" value={v} onChange={(iso) => { calls.push(iso); setV(iso); }} />;
    }
    render(<Harness />);
    fireEvent.change(screen.getByLabelText('รับเข้าถึง'), { target: { value: '01/10/2568' } });
    expect(calls[calls.length - 1]).toBe('2025-10-01');
    expect(screen.getByLabelText('รับเข้าถึง').value).toBe('01/10/2568');
    fireEvent.click(screen.getByRole('button', { name: 'ล้างวันที่ ถึง' }));
    expect(calls[calls.length - 1]).toBe('');
    expect(screen.getByLabelText('รับเข้าถึง').value).toBe('');
  });

  it('พิมพ์ไม่ครบ → ยังไม่เรียก onChange (ไม่กรองครึ่ง ๆ) และ blur คืนข้อความเดิม', () => {
    const onChange = vi.fn();
    render(<ThaiDateInput label="ถึง" ariaLabel="รับเข้าถึง" value="" onChange={onChange} />);
    const input = screen.getByLabelText('รับเข้าถึง');
    fireEvent.change(input, { target: { value: '01/10' } });
    expect(onChange).not.toHaveBeenCalled();
    fireEvent.blur(input);
    expect(input.value).toBe('');
  });
});
