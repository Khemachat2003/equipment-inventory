// farmBundleCount.test.jsx — แยกการนับ "ชุด (Bundles)" ออกจากจำนวนอุปกรณ์บนการ์ดภาพรวมฟาร์ม (/farm)
// บั๊กเดิม: bundleCount ถูกเพิ่มทีละ "ชิ้นอุปกรณ์สมาชิก" ในลูป assets → ชุดจริง 11 ชุดโชว์เป็น 61 (ตามจำนวนของ)
// Fix: นับเฉพาะรหัสชุด (Bundle ID) ที่ไม่ซ้ำกันที่ deploy อยู่ที่ฟาร์มนั้น — อุปกรณ์แยกนับที่ "อุปกรณ์ N"
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { render, screen, within } from '@testing-library/react';
import { MemoryRouter } from 'react-router-dom';
import Farm from '../pages/Farm.jsx';

vi.mock('axios', () => ({ default: { get: vi.fn(), post: vi.fn(), delete: vi.fn() } }));
import axios from 'axios';

// สมาชิกชุด deploy: SiteName = ชื่อชุด (ตามข้อมูลจริง — คอลัมน์ SiteName ของสมาชิกไม่ใช่ชื่อฟาร์ม)
// สถานะสมาชิก = "In Bundle" (ถูกบังคับตอนจัดชุด ตาม routes/bundle.js)
const member = (assetId, serial, bundleId, bundleName) => ({
  assetId,
  serialNumber: serial,
  name: 'อุปกรณ์ ' + assetId,
  status: 'In Bundle',
  bundleId,
  bundleName,
  siteName: bundleName,
  location: 'ฟาร์มหนองแค',
  houseName: '',
  houseId: '',
});

const ASSETS = [
  member('1', 'SN-1', 'BDL-001', 'ชุดวัดกลิ่น-001'),
  member('2', 'SN-2', 'BDL-001', 'ชุดวัดกลิ่น-001'),
  member('3', 'SN-3', 'BDL-001', 'ชุดวัดกลิ่น-001'),
  member('4', 'SN-4', 'BDL-002', 'ชุดวัดกลิ่น-002'),
  member('5', 'SN-5', 'BDL-002', 'ชุดวัดกลิ่น-002'),
  member('6', 'SN-6', 'BDL-002', 'ชุดวัดกลิ่น-002'),
];
const BUNDLES = [
  { bundleId: 'BDL-001', bundleName: 'ชุดวัดกลิ่น-001 (BDL-001)', status: 'Deployed', location: 'ฟาร์มหนองแค', assetIds: ['1', '2', '3'] },
  { bundleId: 'BDL-002', bundleName: 'ชุดวัดกลิ่น-002 (BDL-002)', status: 'Deployed', location: 'ฟาร์มหนองแค', assetIds: ['4', '5', '6'] },
  { bundleId: 'BDL-003', bundleName: 'ชุดเปล่า (BDL-003)', status: 'Deployed', location: 'ฟาร์มเขาซก', assetIds: [] },
  { bundleId: 'BDL-009', bundleName: 'ชุดในคลัง (BDL-009)', status: 'In Stock', location: '', assetIds: [] },
];

beforeEach(() => {
  axios.get.mockImplementation((url) => {
    if (url === '/api/assets') return Promise.resolve({ data: ASSETS });
    if (url === '/api/bundles') return Promise.resolve({ data: BUNDLES });
    return Promise.resolve({ data: [] });
  });
});

function renderFarm() {
  return render(
    <MemoryRouter>
      <Farm />
    </MemoryRouter>,
  );
}

function farmCard(name) {
  return screen.getByText(name).closest('button');
}

describe('การ์ดภาพรวมฟาร์ม (/farm) — นับ "ชุด" เฉพาะรหัสชุดไม่ซ้ำ แยกจาก "อุปกรณ์"', () => {
  it('2 ชุด × 3 ชิ้น ที่ฟาร์มเดียวกัน → "อุปกรณ์ 6" · "ชุด 2" (ไม่ดีดเป็น 6 ชุด)', async () => {
    renderFarm();
    const card = (await screen.findByText('ฟาร์มหนองแค')).closest('button');
    expect(card.textContent).toMatch(/6\s*อุปกรณ์/);
    expect(card.textContent).toMatch(/2\s*ชุด(?!\d)/); // ห้ามเป็น 6 ชุด (จำนวนชิ้นทับจำนวนชุด)
    expect(card.textContent).not.toMatch(/6\s*ชุด/);
  });

  it('ชุด deploy ที่ยังไม่มีสมาชิก → การ์ดฟาร์มยังโชว์ "ชุด 1" · "อุปกรณ์ 0"', async () => {
    renderFarm();
    const card = (await screen.findByText('ฟาร์มเขาซก')).closest('button');
    expect(card.textContent).toMatch(/1\s*ชุด(?!\d)/);
    expect(card.textContent).toMatch(/0\s*อุปกรณ์/);
  });

  it('ชุดที่ยังอยู่คลัง (In Stock) → ไม่สร้างการ์ดฟาร์ม', async () => {
    renderFarm();
    await screen.findByText('ฟาร์มหนองแค');
    expect(screen.queryByText('ชุดในคลัง')).toBeNull();
    const cardNames = (await screen.findAllByText(/ฟาร์ม/)).map((el) => el.textContent);
    expect(cardNames.join(' ')).not.toContain('ชุดในคลัง');
  });

  it('สมาชิกชุด deploy สถานะ "In Bundle" → นับ "ใช้งานได้" ครบทุกชิ้น (6/6)', async () => {
    renderFarm();
    const card = (await screen.findByText('ฟาร์มหนองแค')).closest('button');
    expect(within(card).getByText(/ใช้งานได้\s*6/)).toBeTruthy();
  });
});
