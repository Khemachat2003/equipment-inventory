import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { MemoryRouter, useLocation } from 'react-router-dom';
import PODispatchPicker from '../components/PODispatchPicker.jsx';
import { DISPATCH_QUEUE_KEY } from '../utils/dispatchQueue.js';

vi.mock('axios', () => ({ default: { get: vi.fn(), post: vi.fn() } }));
vi.mock('../components/TransferModal.jsx', () => ({ default: ({ serial, onSuccess, onClose }) => <div role="dialog"><span>{serial}</span><button onClick={onSuccess}>simulate transfer success</button><button onClick={onClose}>close transfer</button></div> }));
import axios from 'axios';

function RouteLocation() {
  const location = useLocation();
  return <output data-testid="route">{location.pathname}{location.search}</output>;
}

async function openPO() {
  render(<MemoryRouter><PODispatchPicker /><RouteLocation /></MemoryRouter>);
  fireEvent.click(await screen.findByTestId('po-dispatch-heading'));
  await screen.findByText(/SN-1/);
}

describe('PO dispatch staging', () => {
  beforeEach(() => {
    sessionStorage.clear();
    axios.get.mockImplementation((url) => {
      if (url.startsWith('/api/inbound-pos')) return Promise.resolve({ data: [{ key: 'po:PO-1', poNumber: 'PO-1', status: 'active', stockCount: 2, assetCount: 3 }, { key: 'po:PO-OLD', poNumber: 'PO-OLD', status: 'closed', stockCount: 0, assetCount: 1 }] });
      if (url === '/api/assets') return Promise.resolve({ data: [
        { assetId: 'A1', serialNumber: 'SN-1', name: 'Sensor', location: 'Stock', siteName: 'Intranin', poNumber: 'PO-1', batchId: 'INB-1' },
        { assetId: 'A2', serialNumber: 'SN-2', name: 'Relay', location: 'Stock', siteName: 'Intranin', poNumber: 'PO-1', batchId: 'INB-1', bundleId: 'BDL-9', bundleName: 'Sensor cabinet' },
        { assetId: 'A3', serialNumber: 'SN-3', name: 'Sensor', location: 'Farm A', siteName: 'Farm A', poNumber: 'PO-1', batchId: 'INB-1' },
        { assetId: 'A4', serialNumber: 'SN-4', name: 'Controller', location: 'Farm B', siteName: 'Farm B', poNumber: 'PO-OLD', batchId: 'INB-OLD' },
      ] });
      return Promise.resolve({ data: [] });
    });
    axios.post.mockResolvedValue({ data: { success: true } });
  });

  it('stages only unbundled stock serials while showing previously assigned and moved items', async () => {
    await openPO();
    expect(screen.getAllByText(/Sensor cabinet/).length).toBeGreaterThan(0);
    expect(screen.getByText(/Farm A/)).toBeTruthy();
    await waitFor(() => expect(screen.getByTestId('dispatch-create-bundle').textContent).toContain('(1)'));
    fireEvent.click(screen.getByTestId('dispatch-create-bundle'));
    await waitFor(() => expect(screen.getByTestId('route').textContent).toBe('/bundle?dispatch=1'));
    expect(JSON.parse(sessionStorage.getItem(DISPATCH_QUEUE_KEY))).toMatchObject({ poNumber: 'PO-1', items: [{ assetId: 'A1', serialNumber: 'SN-1' }] });
  });

  it('keeps closed PO assets visible for status lookup without repeat-transfer actions', async () => {
    await openPO();
    fireEvent.change(screen.getAllByRole('combobox')[0], { target: { value: 'PO-OLD' } });
    expect(await screen.findByText(/SN-4/)).toBeTruthy();
    expect(screen.getByText(/Farm B/)).toBeTruthy();
    expect(screen.queryByRole('button', { name: 'Transfer SN-4' })).toBeNull();
  });

  it('offers a per-serial transfer for stock assets', async () => {
    await openPO();
    fireEvent.click(screen.getByRole('button', { name: 'Transfer SN-1' }));
    expect(await screen.findByRole('dialog')).toBeTruthy();
    expect(screen.getByText('SN-1')).toBeTruthy();
    fireEvent.click(screen.getByRole('button', { name: 'simulate transfer success' }));
    await waitFor(() => expect(screen.getAllByRole('status')[0].textContent.length).toBeGreaterThan(0));
  });

  it('adds selected serials to an existing Bundle', async () => {
    await openPO();
    fireEvent.change(screen.getByRole('combobox', { name: 'Existing Bundle' }), { target: { value: 'BDL-9' } });
    fireEvent.click(screen.getByTestId('dispatch-add-existing'));
    await waitFor(() => expect(axios.post).toHaveBeenCalledWith('/api/bundles/BDL-9/assets/bulk', { assetIds: ['A1'] }));
    await waitFor(() => expect(screen.getAllByRole('status')[0].textContent).toContain('Sensor cabinet'));
  });
});
