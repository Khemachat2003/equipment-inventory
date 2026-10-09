import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, within } from '@testing-library/react';
import Asset from '../pages/Asset.jsx';

vi.mock('axios', () => ({ default: { get: vi.fn(), post: vi.fn() } }));
import axios from 'axios';

const ASSETS = [
  { assetId: 'A1', partNumber: 'SENSOR', code: 'SENSOR', name: 'Sensor', serialNumber: 'SN-1', status: 'Ready', location: 'Stock', siteName: 'Intranin', batchId: 'INB-OLD', poNumber: 'PO-OLD' },
  { assetId: 'A2', partNumber: 'SENSOR', code: 'SENSOR', name: 'Sensor', serialNumber: 'SN-2', status: 'Ready', location: 'Farm A', siteName: 'Farm A', batchId: 'INB-OLD', poNumber: 'PO-OLD' },
  { assetId: 'A3', partNumber: 'SENSOR', code: 'SENSOR', name: 'Sensor', serialNumber: 'SN-3', status: 'Ready', location: 'Farm B', siteName: 'Farm B', batchId: 'INB-NEW', poNumber: 'PO-NEW' },
];
const POS = [
  { key: 'po:PO-OLD', poNumber: 'PO-OLD', batchId: 'INB-OLD', batches: [{ batchId: 'INB-OLD', count: 2 }], receivedAt: '2026-10-01T00:00:00.000Z', assetCount: 2, stockCount: 1, batchCount: 1, status: 'active', manuallyClosed: false },
  { key: 'po:PO-NEW', poNumber: 'PO-NEW', batchId: 'INB-NEW', batches: [{ batchId: 'INB-NEW', count: 1 }], receivedAt: '2026-10-09T00:00:00.000Z', assetCount: 1, stockCount: 0, batchCount: 1, status: 'closed', manuallyClosed: false },
];

describe('Asset PO lifecycle', () => {
  beforeEach(() => {
    axios.get.mockImplementation((url) => {
      if (url === '/api/assets') return Promise.resolve({ data: ASSETS });
      if (url === '/api/part-catalog') return Promise.resolve({ data: [{ partNumber: 'SENSOR', partName: 'Sensor' }] });
      if (url.startsWith('/api/inbound-pos')) return Promise.resolve({ data: url.includes('status=active') ? POS.filter((po) => po.status === 'active') : POS });
      return Promise.resolve({ data: [] });
    });
    axios.post.mockResolvedValue({ data: { success: true } });
  });

  it('shows active PO quick access and filters history by lifecycle tab/search', async () => {
    render(<Asset />);
    expect(await screen.findByText('PO PO-OLD')).toBeTruthy();
    expect(screen.queryByText('PO PO-NEW')).toBeNull();
    fireEvent.click(screen.getByRole('button', { name: /ดูประวัติ PO ทั้งหมด/ }));
    const dialog = await screen.findByRole('dialog');
    fireEvent.click(within(dialog).getByRole('button', { name: /ปิดแล้ว/ }));
    expect(within(dialog).getByText(/PO PO-NEW/)).toBeTruthy();
    expect(within(dialog).queryByText(/PO PO-OLD/)).toBeNull();
  });

  it('allows a user to mark an active PO closed', async () => {
    render(<Asset />);
    fireEvent.click(await screen.findByRole('button', { name: /ดูประวัติ PO ทั้งหมด/ }));
    const dialog = await screen.findByRole('dialog');
    fireEvent.click(within(dialog).getByRole('button', { name: 'ปิด PO' }));
    expect(axios.post).toHaveBeenCalledWith('/api/inbound-pos/PO-OLD/close');
  });
});