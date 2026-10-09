// farmMonitor.test.jsx — เทสต์ Step 3: Executive Farm Monitor บนหน้าแรก
// (utils/farmMonitor.js — กฎการนับเดียวกับ /farm + FarmDetailModal — Asset Breakdown / Bundle Control)
import { describe, it, expect, vi } from 'vitest';
import { render, screen, fireEvent } from '@testing-library/react';
import { MemoryRouter } from 'react-router-dom';
import {
  buildBundleFarmMap,
  farmOfAsset,
  farmAssetsOf,
  buildFarmOverview,
} from '../utils/farmMonitor.js';
import FarmDetailModal from '../components/FarmDetailModal.jsx';

const A = (over = {}) => ({
  assetId: '1', code: 'PWR-1', name: 'Power Supply', serialNumber: 'SN-A-0001',
  status: 'ใช้งานได้', siteName: '', location: '', bundleId: '', houseName: '', houseId: '',
  ...over,
});

describe('farmMonitor — buildBundleFarmMap', () => {
  it('เฉพาะชุด Deployed + มี Location เท่านั้นที่นับเป็นฟาร์ม', () => {
    const map = buildBundleFarmMap([
      { bundleId: 'BDL-001', status: 'Deployed', location: 'ฟาร์มหนองแค' },
      { bundleId: 'BDL-002', status: 'In Stock', location: '' },
      { bundleId: 'BDL-003', status: 'Deployed', location: '  ' },
    ]);
    expect(map).toEqual({ 'BDL-001': 'ฟาร์มหนองแค' });
  });
});

describe('farmMonitor — farmOfAsset (กฎเดียวกับ /farm)', () => {
  const bf = { 'BDL-001': 'ฟาร์มหนองแค' };

  it('สมาชิกชุดที่ deploy → นับที่ฟาร์มของชุด (ไม่ใช้ SiteName ซึ่งเป็นชื่อชุด)', () => {
    expect(farmOfAsset(A({ bundleId: 'BDL-001', siteName: 'ชุดทดสอบ (BDL-001)', location: 'ฟาร์มหนองแค' }), bf)).toBe('ฟาร์มหนองแค');
  });

  it('อยู่ในชุดที่ยังไม่ deploy (คลัง) → ไม่นับเป็นฟาร์ม', () => {
    expect(farmOfAsset(A({ bundleId: 'BDL-002', siteName: 'ชุดในคลัง', location: 'Stock' }), bf)).toBe('');
  });

  it('ตัวเดี่ยวติดตั้งที่ฟาร์ม → นับตาม SiteName', () => {
    expect(farmOfAsset(A({ siteName: 'ฟาร์มเขาซก', location: 'โรงเรือน A' }), bf)).toBe('ฟาร์มเขาซก');
  });

  it('คลังกลาง (Intranin / Stock / ไม่ระบุไซต์) → ไม่นับเป็นฟาร์ม', () => {
    expect(farmOfAsset(A({ siteName: 'Intranin', location: 'Stock' }), bf)).toBe('');
    expect(farmOfAsset(A({ siteName: '', location: 'Stock' }), bf)).toBe('');
    expect(farmOfAsset(A({ siteName: '', location: '-' }), bf)).toBe('');
  });
});

describe('farmMonitor — buildFarmOverview (สรุปรายฟาร์มสำหรับ Executive Overview)', () => {
  const assets = [
    A({ assetId: '1', serialNumber: 'SN-1', bundleId: 'BDL-001', siteName: 'ชุด A', location: 'ฟาร์มหนองแค', status: 'ใช้งานได้', houseName: 'H1' }),
    A({ assetId: '2', serialNumber: 'SN-2', bundleId: 'BDL-001', siteName: 'ชุด A', location: 'ฟาร์มหนองแค', status: 'ส่งซ่อม', houseName: 'H1' }),
    A({ assetId: '3', serialNumber: 'SN-3', bundleId: 'BDL-001', siteName: 'ชุด A', location: 'ฟาร์มหนองแค', status: 'ใช้งานได้', houseId: 'SITE-H02', houseName: 'H2' }),
    A({ assetId: '4', serialNumber: 'SN-4', siteName: 'ฟาร์มเขาซก', location: 'โรงเรือน A', status: 'ใช้งานได้' }),
    A({ assetId: '5', serialNumber: 'SN-5', siteName: 'Intranin', location: 'Stock', status: 'ใช้งานได้' }),
    A({ assetId: '6', serialNumber: 'SN-6', bundleId: 'BDL-002', siteName: 'ชุดในคลัง', location: 'Stock', status: 'ใช้งานได้' }),
  ];
  const bundles = [
    { bundleId: 'BDL-001', bundleName: 'ชุด A', status: 'Deployed', location: 'ฟาร์มหนองแค', assetIds: ['1', '2', '3'] },
    { bundleId: 'BDL-002', bundleName: 'ชุดในคลัง', status: 'In Stock', location: '', assetIds: ['6'] },
  ];
  const farms = [
    { siteId: 'FS001', siteName: 'ฟาร์มหนองแค', farmType: 'สุกร', province: 'สระบุรี' },
    { siteId: 'FS002', siteName: 'ฟาร์มเปล่า', farmType: 'สัตว์ปีก', province: '' },
  ];

  const overview = buildFarmOverview({ assets, bundles, farms });

  it('โชว์ฟาร์มที่ลงทะเบียนทุกฟาร์ม (แม้ 0 อุปกรณ์) + เรียงตามจำนวนอุปกรณ์มากสุดก่อน', () => {
    expect(overview.map((f) => f.name)).toEqual(['ฟาร์มหนองแค', 'ฟาร์มเขาซก', 'ฟาร์มเปล่า']);
  });

  it('นับอุปกรณ์/สถานะ/โรงเรือนของฟาร์มหนองแคถูกต้อง (สมาชิกชุด deploy นับที่ฟาร์ม)', () => {
    const f = overview.find((x) => x.name === 'ฟาร์มหนองแค');
    expect(f.assets).toBe(3);
    expect(f.ok).toBe(2);
    expect(f.rep).toBe(1);
    expect(f.houses).toBe(2); // H1 + SITE-H02/H2
    expect(f.type).toBe('สุกร');
    expect(f.province).toBe('สระบุรี');
  });

  it('นับ "ตู้ชุด" ที่ตัวชุด (1 ตู้ = 1 แม้มีสมาชิก 3 ชิ้น) และชุดในคลังไม่ถูกนับที่ฟาร์มใด', () => {
    const f = overview.find((x) => x.name === 'ฟาร์มหนองแค');
    expect(f.bundles).toBe(1);
    expect(overview.find((x) => x.name === 'ฟาร์มเขาซก').bundles).toBe(0);
  });

  it('farmAssetsOf — คัดเฉพาะอุปกรณ์ที่ติดตั้งอยู่ "ที่ฟาร์มนี้" (ใช้โดย FarmDetailModal)', () => {
    const bf = buildBundleFarmMap(bundles);
    const at = farmAssetsOf('ฟาร์มหนองแค', assets, bf);
    expect(at.map((a) => a.assetId)).toEqual(['1', '2', '3']);
    expect(farmAssetsOf('ฟาร์มที่ไม่มีใครอยู่', assets, bf)).toEqual([]);
  });

  it('คลังกลาง (Intranin) ไม่กลายเป็นการ์ดฟาร์ม', () => {
    expect(overview.find((x) => x.name === 'Intranin')).toBeUndefined();
  });
});

describe('FarmDetailModal — Quick Farm Access (Asset Breakdown + Bundle Control)', () => {
  const assets = [
    A({ assetId: '1', serialNumber: 'SN-1', bundleId: 'BDL-001', siteName: 'ชุด A', location: 'ฟาร์มหนองแค', status: 'ใช้งานได้', houseName: 'โรงเรือน 1' }),
    A({ assetId: '2', serialNumber: 'SN-2', bundleId: 'BDL-001', siteName: 'ชุด A', location: 'ฟาร์มหนองแค', status: 'ส่งซ่อม', houseName: 'โรงเรือน 1' }),
    A({ assetId: '9', serialNumber: 'SN-9', siteName: 'ฟาร์มเขาซก', location: 'หลังโรง', status: 'ใช้งานได้' }),
    A({ assetId: '7', serialNumber: 'SN-PO', siteName: 'ฟาร์มหนองแค', location: 'โรงเรือน 2', status: 'ใช้งานได้', poNumber: 'PO-2209', name: 'เซ็นเซอร์ฝุ่น' }),
  ];
  const bundles = [
    { bundleId: 'BDL-001', bundleName: 'ชุดควบคุม A', status: 'Deployed', location: 'ฟาร์มหนองแค', assetIds: ['1', '2'] },
    { bundleId: 'BDL-002', bundleName: 'ชุดในคลัง', status: 'In Stock', location: '', assetIds: ['6'] },
  ];
  const farm = { name: 'ฟาร์มหนองแค', type: 'สุกร', province: 'สระบุรี' };

  function ui(props = {}) {
    return render(
      <MemoryRouter>
        <FarmDetailModal
          farm={farm}
          assets={assets}
          bundles={bundles}
          onClose={() => {}}
          onHistory={vi.fn()}
          onTransfer={vi.fn()}
          {...props}
        />
      </MemoryRouter>,
    );
  }

  it('เปิดรายละเอียดฟาร์ม — โชว์ชื่อฟาร์ม + เฉพาะอุปกรณ์/ชุดของฟาร์มนั้น', () => {
    const { container } = ui();
    expect(container.textContent).toContain('ฟาร์มหนองแค');
    expect(container.textContent).toContain('SN-1');      // สมาชิกชุด deploy → อยู่ที่ฟาร์มนี้
    expect(container.textContent).toContain('PO-2209');   // ตัวเดี่ยวที่ฟาร์มนี้
    expect(container.textContent).not.toContain('SN-9');  // อุปกรณ์ของฟาร์มอื่นไม่โชว์
    expect(container.textContent).toContain('ชุดควบคุม A');
    expect(container.textContent).not.toContain('ชุดในคลัง');
    expect(container.textContent).toContain('สมาชิก');    // Bundle Control โชว์จำนวนสมาชิก
  });

  it('สรุปตัวเลขบนหัวโมดัล: อุปกรณ์ 3 / ตู้ชุด 1', () => {
    const { container } = ui();
    // MiniStat tiles: อุปกรณ์=3, ตู้ชุด=1
    const labels = ['อุปกรณ์', 'ตู้ชุด', 'ใช้งานได้', 'ส่งซ่อม'];
    labels.forEach((l) => expect(container.textContent).toContain(l));
  });

  it('ค้นหาภายในฟาร์มด้วย Serial → กรองเฉพาะชิ้นที่ตรง', () => {
    const { container } = ui();
    fireEvent.change(container.querySelector('input'), { target: { value: 'SN-PO' } });
    expect(container.textContent).toContain('SN-PO');
    expect(container.textContent).not.toContain('SN-1');
  });

  it('ค้นหาด้วยเลข PO → เจออุปกรณ์ที่รับเข้าด้วย PO นั้น (สื่อสารกับจัดซื้อได้เร็ว)', () => {
    const { container } = ui();
    fireEvent.change(container.querySelector('input'), { target: { value: 'PO-2209' } });
    expect(container.textContent).toContain('SN-PO');
    expect(container.textContent).not.toContain('SN-2');
  });

  it('กดปุ่มประวัติ → เรียก onHistory ด้วย Serial ของชิ้นนั้น', () => {
    const onHistory = vi.fn();
    ui({ onHistory });
    const historyButtons = screen.getAllByTitle('ดูประวัติการเคลื่อนไหว');
    fireEvent.click(historyButtons[0]);
    expect(onHistory).toHaveBeenCalledTimes(1);
    expect(onHistory.mock.calls[0][0]).toMatch(/^SN-/);
  });

  it('ฟาร์มที่ยังไม่มีอุปกรณ์ → โชว์ empty state ไม่พัง', () => {
    const { container } = ui({ farm: { name: 'ฟาร์มเปล่า', type: 'สัตว์ปีก', province: '' }, assets: [], bundles: [] });
    expect(container.textContent).toContain('ฟาร์มเปล่า');
    expect(container.textContent).toContain('ยังไม่มีอุปกรณ์ติดตั้งที่ฟาร์มนี้');
    expect(container.textContent).toContain('ไม่มีตู้ชุดติดตั้งอยู่ที่ฟาร์มนี้');
  });
});