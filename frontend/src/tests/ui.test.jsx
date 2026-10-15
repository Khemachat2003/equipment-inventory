import { describe, it, expect, vi } from 'vitest';
import { render, screen, fireEvent } from '@testing-library/react';
import Pagination from '../components/ui/Pagination.jsx';
import StatusBadge, { statusTone, statusLabel, historyTypeTone } from '../components/ui/StatusBadge.jsx';

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

describe('statusTone — B2 Modern Pill (เขียว=ใช้งานได้ · ฟ้า=อยู่ในชุดติดตั้ง/คลัง · ส้ม=ส่งซ่อม/สำรอง · เทา=ชำรุด/สูญ)', () => {
  it('จับคู่สถานะรายชิ้น → tone ถูกต้อง', () => {
    expect(statusTone('ใช้งานได้')).toBe('green');
    expect(statusTone('สำรอง')).toBe('amber');
    expect(statusTone('ส่งซ่อม')).toBe('amber');
    expect(statusTone('ชำรุด')).toBe('gray');
    expect(statusTone('สูญหาย')).toBe('gray');
    expect(statusTone('')).toBe('green'); // fallback
  });

  it('ค่าสถานะระบบของชุด/รายชิ้น (อังกฤษจาก backend) → tone ภาษาไทยถูกโทน', () => {
    expect(statusTone('In Bundle')).toBe('blue'); // อยู่ในชุดติดตั้ง
    expect(statusTone('In Stock')).toBe('blue'); // ชุดอยู่ในคลัง
    expect(statusTone('Deployed')).toBe('green'); // ติดตั้งแล้ว
    expect(statusTone('Maintenance')).toBe('amber'); // ซ่อมบำรุง
  });
});

describe('statusLabel — B2 Strict Terminology (แปลงศัพท์ระบบเป็นภาษาไทยจุดเดียว)', () => {
  it('ค่าข้อมูล "In Bundle" → แสดง "อยู่ในชุดติดตั้ง"', () => {
    expect(statusLabel('In Bundle')).toBe('อยู่ในชุดติดตั้ง');
  });

  it('สถานะชุดติดตั้ง: In Stock → อยู่ในคลัง / Deployed → ติดตั้งแล้ว / Maintenance → ซ่อมบำรุง', () => {
    expect(statusLabel('In Stock')).toBe('อยู่ในคลัง');
    expect(statusLabel('Deployed')).toBe('ติดตั้งแล้ว');
    expect(statusLabel('Maintenance')).toBe('ซ่อมบำรุง');
  });

  it('สถานะไทยอื่นคงค่าเดิม', () => {
    expect(statusLabel('ใช้งานได้')).toBe('ใช้งานได้');
    expect(statusLabel('ส่งซ่อม')).toBe('ส่งซ่อม');
  });
});

describe('StatusBadge — Modern Pill Badge', () => {
  it('แสดงข้อความสถานะ (หรือ - ถ้าว่าง) เป็น pill โค้งมีขอบสี + จุดสถานะ', () => {
    const { container, rerender } = render(<StatusBadge status="ใช้งานได้" />);
    expect(container.textContent).toBe('ใช้งานได้');
    const pill = container.firstChild;
    expect(pill.className).toContain('rounded-full');
    expect(pill.className).toContain('border-[var(--emerald-b)]');
    expect(pill.querySelector('span[aria-hidden="true"]')).toBeTruthy();
    rerender(<StatusBadge status="" />);
    expect(container.textContent).toBe('-');
  });

  it('ค่าข้อมูล "In Bundle" (จาก backend) → แสดงภาษาไทย "อยู่ในชุดติดตั้ง" tone ฟ้า — ไม่มีศัพท์ระบบหลุดขึ้นจอ', () => {
    const { container } = render(<StatusBadge status="In Bundle" />);
    expect(container.textContent).toBe('อยู่ในชุดติดตั้ง');
    expect(container.textContent).not.toContain('In Bundle');
    expect(container.firstChild.className).toContain('border-[var(--blue-b)]');
  });
});

describe('historyTypeTone', () => {
  it('เบิก → blue, คืน → green, อื่นๆ → gray', () => {
    expect(historyTypeTone('เบิก')).toBe('blue');
    expect(historyTypeTone('คืน')).toBe('green');
    expect(historyTypeTone('โอน')).toBe('gray');
  });
});
