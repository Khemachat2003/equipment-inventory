import { describe, it, expect, vi } from 'vitest';
import { render, screen, fireEvent } from '@testing-library/react';
import Pagination from '../components/ui/Pagination.jsx';
import StatusBadge, { statusTone, historyTypeTone } from '../components/ui/StatusBadge.jsx';

describe('Pagination', () => {
  it('แสดงช่วงรายการ (1–20 จาก 105) และเลขหน้าถูกต้อง', () => {
    const { container } = render(
      <Pagination total={105} page={1} pageSize={20} onPage={() => {}} onPageSize={() => {}} />
    );
    expect(container.textContent).toContain('แสดง');
    expect(container.textContent).toContain('1–20');
    expect(container.textContent).toContain('จาก 105');
    // หน้า 1 → window 1..5
    [1, 2, 3, 4, 5].forEach((p) => {
      expect(screen.getByRole('button', { name: String(p) })).toBeTruthy();
    });
  });

  it('กดปุ่ม › → onPage(2)', () => {
    const onPage = vi.fn();
    render(<Pagination total={105} page={1} pageSize={20} onPage={onPage} onPageSize={() => {}} />);
    fireEvent.click(screen.getByText('›'));
    expect(onPage).toHaveBeenCalledWith(2);
  });

  it('total = 0 → ไม่ render อะไรเลย (null)', () => {
    const { container } = render(
      <Pagination total={0} page={1} pageSize={20} onPage={() => {}} onPageSize={() => {}} />
    );
    expect(container.firstChild).toBeNull();
  });

  it('กดปุ่ม ‹ ตอนหน้า 1 → disabled (onPage ไม่ถูกเรียก)', () => {
    const onPage = vi.fn();
    render(<Pagination total={105} page={1} pageSize={20} onPage={onPage} onPageSize={() => {}} />);
    const prev = screen.getByText('‹');
    expect(prev.disabled).toBe(true);
    fireEvent.click(prev);
    expect(onPage).not.toHaveBeenCalled();
  });
});

describe('statusTone', () => {
  it('จับคู่สถานะ → tone ถูกต้อง', () => {
    expect(statusTone('ใช้งานได้')).toBe('green');
    expect(statusTone('สำรอง')).toBe('amber');
    expect(statusTone('ส่งซ่อม')).toBe('red');
    expect(statusTone('ชำรุด')).toBe('gray');
    expect(statusTone('สูญหาย')).toBe('gray');
    expect(statusTone('')).toBe('green'); // fallback
  });
});

describe('StatusBadge', () => {
  it('แสดงข้อความสถานะ (หรือ - ถ้าว่าง)', () => {
    const { container, rerender } = render(<StatusBadge status="ใช้งานได้" />);
    expect(container.textContent).toBe('ใช้งานได้');
    rerender(<StatusBadge status="" />);
    expect(container.textContent).toBe('-');
  });
});

describe('historyTypeTone', () => {
  it('เบิก → blue, คืน → green, อื่นๆ → gray', () => {
    expect(historyTypeTone('เบิก')).toBe('blue');
    expect(historyTypeTone('คืน')).toBe('green');
    expect(historyTypeTone('โอน')).toBe('gray');
  });
});