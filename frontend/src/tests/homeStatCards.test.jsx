// homeStatCards.test.jsx — การ์ดสรุปสถานะ 3 ใบด้านบนหน้าแรก
//   • ตัวเลขต้อง "นับจากตำแหน่ง/สถานะจริงของทุกชิ้น" ไม่กรองเฉพาะคำ "สำรอง" (เดิมขึ้น 0 ทั้งที่มีของในคลัง)
//   • กดการ์ด → deep-link ไป /asset ด้วยตัวกรองที่ตรงกับตัวเลขการ์ด
//     (อยู่ในคลัง → ?location=stock · ติดตั้งแล้ว → ?location=farm · ส่งซ่อม → ?status=repair)
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, waitFor, within } from '@testing-library/react';
import { MemoryRouter, Route, Routes, useLocation } from 'react-router-dom';
import Home from '../pages/Home.jsx';
import { isInStock } from '../utils/location.js';
import { isRepairStatus } from '../utils/status.js';

vi.mock('axios', () => ({ default: { get: vi.fn(), post: vi.fn(), delete: vi.fn() } }));
import axios from 'axios';

const A = (over = {}) => ({
  assetId: '1', code: 'PWR-1', name: 'Power Supply', serialNumber: 'SN-1',
  status: 'ใช้งานได้', siteName: 'Intranin', location: 'Stock', bundleId: '', houseName: '', houseId: '',
  ...over,
});

// ข้อมูลจำลองครบทุกกรณีตาม spec:
const ASSETS = [
  // คลังกลาง — สถานะ "ใช้งานได้" (เดิมกรองเฉพาะ "สำรอง" จึงหลุดจากการ์ด)
  A({ assetId: '1', serialNumber: 'SN-ST1', status: 'ใช้งานได้', siteName: 'Intranin', location: 'Stock' }),
  // คลังกลาง — สถานะ "สำรอง"
  A({ assetId: '2', serialNumber: 'SN-ST2', status: 'สำรอง', siteName: 'Intranin', location: 'Stock' }),
  // ไม่ระบุฟาร์มจัดเก็บ → ถือว่ายังอยู่คลัง
  A({ assetId: '3', serialNumber: 'SN-ST3', status: 'ใช้งานได้', siteName: '', location: '' }),
  // สมาชิกชุดที่ยังอยู่คลัง (ชุดยังไม่ deploy)
  A({ assetId: '4', serialNumber: 'SN-ST4', status: 'In Bundle', siteName: 'ชุดในคลัง', location: 'Stock', bundleId: 'BDL-2' }),
  // ตัวเดี่ยวที่ติดตั้งที่ฟาร์ม
  A({ assetId: '5', serialNumber: 'SN-F1', status: 'ใช้งานได้', siteName: 'ฟาร์มเขาซก', location: 'โรงเรือน A' }),
  // สมาชิกชุดที่ deploy ที่ฟาร์ม (Location = ชื่อฟาร์ม, SiteName = ชื่อชุด)
  A({ assetId: '6', serialNumber: 'SN-F2', status: 'In Bundle', siteName: 'ชุด A', location: 'ฟาร์มหนองแค', bundleId: 'BDL-1' }),
  // อยู่ที่ฟาร์มแต่ส่งซ่อม → นับทั้ง "ติดตั้งแล้ว" และ "ส่งซ่อม"
  A({ assetId: '7', serialNumber: 'SN-F3', status: 'ส่งซ่อม', siteName: 'ฟาร์มเขาซก', location: 'โรงเรือน B' }),
  // ชำรุดที่ยังอยู่คลัง → นับทั้ง "อยู่ในคลัง" และ "ส่งซ่อม"
  A({ assetId: '8', serialNumber: 'SN-R1', status: 'ชำรุด/สูญหาย', siteName: 'Intranin', location: 'Stock' }),
];
const BUNDLES = [
  { bundleId: 'BDL-1', bundleName: 'ชุด A', status: 'Deployed', location: 'ฟาร์มหนองแค' },
  { bundleId: 'BDL-2', bundleName: 'ชุดในคลัง', status: 'In Stock', location: 'Stock' },
];
const FARMS = [{ siteId: 'FS001', siteName: 'ฟาร์มเขาซก', farmType: 'สุกร', province: 'สระบุรี' }];

describe('utils — กติกาตำแหน่ง/สถานะที่ใช้ร่วมกัน (หน้าแรก + /asset)', () => {
  it('isInStock: Intranin (Stock)/ไม่ระบุฟาร์ม/ชุดยังไม่ deploy → คลัง · ฟาร์มจริง → ไม่ใช่คลัง', () => {
    expect(isInStock({ siteName: 'Intranin', location: 'Stock' })).toBe(true);
    expect(isInStock({ siteName: '', location: '' })).toBe(true);
    expect(isInStock({ siteName: 'ชุดในคลัง', location: 'Stock', bundleId: 'BDL-2' })).toBe(true);
    expect(isInStock({ siteName: 'ฟาร์มเขาซก', location: 'โรงเรือน A' })).toBe(false);
    expect(isInStock({ siteName: 'ชุด A', location: 'ฟาร์มหนองแค', bundleId: 'BDL-1' })).toBe(false);
  });

  it('isRepairStatus: ส่งซ่อม/รอซ่อม/ชำรุด·สูญ/Damaged → true · ใช้งานได้/สำรอง → false', () => {
    expect(isRepairStatus('ส่งซ่อม')).toBe(true);
    expect(isRepairStatus('รอซ่อม')).toBe(true);
    expect(isRepairStatus('ชำรุด/สูญหาย')).toBe(true);
    expect(isRepairStatus('Damaged')).toBe(true);
    expect(isRepairStatus('ใช้งานได้')).toBe(false);
    expect(isRepairStatus('สำรอง')).toBe(false);
    expect(isRepairStatus('')).toBe(false);
  });
});

// statusCount จงใจให้ "สำรอง: 0" — พิสูจน์ว่าตัวเลขการ์ดไม่ได้มาจาก statusCount อีกต่อไป
function mockApi() {
  axios.get.mockImplementation((url) => {
    if (url === '/api/assets') return Promise.resolve({ data: ASSETS });
    if (url === '/api/bundles') return Promise.resolve({ data: BUNDLES });
    if (url === '/api/farms') return Promise.resolve({ data: FARMS });
    if (url === '/api/dashboard-full') return Promise.resolve({ data: { assets: ASSETS, farmSites: [], statusCount: { 'สำรอง': 0 }, totalFarmAssets: 0, totalStockAssets: 0 } });
    return Promise.resolve({ data: [] });
  });
}

// probe route — จับ destination ของ navigate() (MemoryRouter ไม่แตะ window.location)
function RouteProbe() {
  const loc = useLocation();
  return <div data-testid="route-probe">{loc.pathname}{loc.search}</div>;
}

function renderHome() {
  return render(
    <MemoryRouter>
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/asset" element={<RouteProbe />} />
      </Routes>
    </MemoryRouter>,
  );
}

describe('Home — การ์ดสรุปสถานะ 3 ใบ (ตัวเลขจากตำแหน่งจริง + Deep-Link)', () => {
  beforeEach(() => {
    sessionStorage.clear();
    mockApi();
  });

  it('การ์ด "อยู่ในคลัง" นับทุกชิ้นที่ยังไม่ติดตั้งที่ฟาร์ม = 5 (ไม่ใช่ 0 เพราะกรองเฉพาะ "สำรอง")', async () => {
    renderHome();
    const card = await screen.findByRole('button', { name: /อยู่ในคลัง/ });
    // 5 ชิ้น = ST1 (ใช้งานได้ที่คลัง) + ST2 (สำรอง) + ST3 (ไม่ระบุฟาร์ม) + ST4 (ชุดยังไม่ deploy) + R1 (ชำรุดที่คลัง)
    expect(within(card).getByText('5')).toBeTruthy();
    // sub สะท้อนความจริง — ไม่ใช่ "พร้อมให้เบิก" เดิม
    expect(within(card).getByText('พร้อมใช้งานในคลังกลาง')).toBeTruthy();
    expect(screen.queryByText('พร้อมให้เบิก')).toBeNull();
  });

  it('กดการ์ด "อยู่ในคลัง" → /asset?location=stock (ห้ามใช้ ?status=สำรอง)', async () => {
    renderHome();
    fireEvent.click(await screen.findByRole('button', { name: /อยู่ในคลัง/ }));
    await waitFor(() => expect(screen.getByTestId('route-probe').textContent).toBe('/asset?location=stock'));
  });

  it('การ์ด "ติดตั้งแล้ว" นับเฉพาะชิ้นที่ผูกเข้ากับฟาร์ม = 3 และกด → /asset?location=farm', async () => {
    renderHome();
    const card = await screen.findByRole('button', { name: /ติดตั้งแล้ว/ });
    // 3 ชิ้น = F1 (ตัวเดี่ยว) + F2 (สมาชิกชุด deploy) + F3 (ตัวเดี่ยวส่งซ่อม — แต่ผูกฟาร์มแล้ว)
    expect(within(card).getByText('3')).toBeTruthy();
    // sub แสดงจำนวนฟาร์มจริง: เขาซก + หนองแค
    expect(within(card).getByText('ใน 2 ฟาร์ม')).toBeTruthy();
    fireEvent.click(card);
    await waitFor(() => expect(screen.getByTestId('route-probe').textContent).toBe('/asset?location=farm'));
  });

  it('การ์ด "ส่งซ่อม" นับ ส่งซ่อม/ชำรุด = 2 และกด → /asset?status=repair', async () => {
    renderHome();
    const card = await screen.findByRole('button', { name: /ส่งซ่อม/ });
    // 2 ชิ้น = F3 (ส่งซ่อม) + R1 (ชำรุด/สูญหาย)
    expect(within(card).getByText('2')).toBeTruthy();
    fireEvent.click(card);
    await waitFor(() => expect(screen.getByTestId('route-probe').textContent).toBe('/asset?status=repair'));
  });
});

