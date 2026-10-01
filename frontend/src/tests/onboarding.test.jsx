// onboarding.test.jsx — ทัวร์แนะนำแสดงครั้งแรกครั้งเดียว (จำใน localStorage)
import { describe, it, expect, beforeEach } from 'vitest';
import { render, screen, fireEvent } from '@testing-library/react';
import OnboardingTour from '../components/OnboardingTour.jsx';

describe('OnboardingTour', () => {
  beforeEach(() => {
    try { localStorage.clear(); } catch { /* ignore */ }
  });

  it('ครั้งแรกแสดงสไลด์ 1/4 พร้อมปุ่ม ถัดไป/ข้าม', () => {
    render(<OnboardingTour />);
    expect(screen.getByText(/รู้จักระบบ 1\/4/)).toBeTruthy();
    expect(screen.getByText('ถัดไป')).toBeTruthy();
    expect(screen.getByText('ข้าม')).toBeTruthy();
  });

  it('กดข้าม → ปิด และจำใน localStorage (เปิดครั้งหน้าไม่ขึ้นอีก)', () => {
    const { unmount } = render(<OnboardingTour />);
    fireEvent.click(screen.getByText('ข้าม'));
    expect(screen.queryByText(/รู้จักระบบ/)).toBeNull();
    expect(localStorage.getItem('intranin-onboarded-v1')).toBeTruthy();
    unmount();
    render(<OnboardingTour />);
    expect(screen.queryByText(/รู้จักระบบ/)).toBeNull();
  });

  it('เคยดูแล้ว → ไม่แสดงอีก', () => {
    localStorage.setItem('intranin-onboarded-v1', '1');
    render(<OnboardingTour />);
    expect(screen.queryByText(/รู้จักระบบ/)).toBeNull();
  });

  it('กดถัดไปจนครบ 4 สไลด์ → ปิดเอง และจำว่าดูแล้ว', () => {
    render(<OnboardingTour />);
    for (let i = 0; i < 4; i++) {
      fireEvent.click(screen.getByText(i === 3 ? 'เริ่มใช้งาน' : 'ถัดไป'));
    }
    expect(screen.queryByText(/รู้จักระบบ/)).toBeNull();
    expect(localStorage.getItem('intranin-onboarded-v1')).toBeTruthy();
  });
});