import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { MemoryRouter, Route, Routes } from 'react-router-dom';
import Bundle from '../pages/Bundle.jsx';
import { writeDispatchQueue } from '../utils/dispatchQueue.js';

vi.mock('axios', () => ({ default: { get: vi.fn(), post: vi.fn() } }));
import axios from 'axios';

describe('PO dispatch Bundle assembly', () => {
  beforeEach(() => {
    sessionStorage.clear();
    writeDispatchQueue({ poNumber: 'PO-1', items: [
      { assetId: 'A1', serialNumber: 'SN-1', name: 'Sensor', batchId: 'INB-1' },
      { assetId: 'A2', serialNumber: 'SN-2', name: 'Relay', batchId: 'INB-1' },
    ] });
    axios.get.mockImplementation((url) => {
      if (url === '/api/bundles') return Promise.resolve({ data: [] });
      if (url === '/api/farms') return Promise.resolve({ data: [] });
      if (url === '/api/assets') return Promise.resolve({ data: [
        { assetId: 'A1', serialNumber: 'SN-1', poNumber: 'PO-1', location: 'Stock', siteName: 'Intranin', bundleId: '' },
        { assetId: 'A2', serialNumber: 'SN-2', poNumber: 'PO-1', location: 'Stock', siteName: 'Intranin', bundleId: '' },
      ] });
      return Promise.resolve({ data: [] });
    });
    axios.post.mockResolvedValue({ data: { success: true, message: 'ok' } });
  });

  it('creates a Bundle from only the selected subset and leaves the rest staged', async () => {
    render(<MemoryRouter initialEntries={['/bundle?dispatch=1']}><Routes><Route path="/bundle" element={<Bundle />} /></Routes></MemoryRouter>);
    expect(await screen.findByText(/Bundle.*PO/)).toBeTruthy();
    fireEvent.click(screen.getByLabelText(/SN-2/));
    fireEvent.change(screen.getAllByText(/Bundle \*/)[1].parentElement.querySelector('input'), { target: { value: 'Fan cabinet 1' } });
    await waitFor(() => expect(screen.getByRole('button', { name: /Serial/ }).textContent).toContain('1'));
    fireEvent.click(screen.getByRole('button', { name: /Serial/ }));
    await waitFor(() => expect(axios.post).toHaveBeenCalledWith('/api/bundles/BDL-001/assets/bulk', { assetIds: ['A1'] }));
    expect(axios.post).toHaveBeenCalledWith('/api/bundles', expect.objectContaining({ bundleId: 'BDL-001', bundleName: 'Fan cabinet 1' }));
    expect(JSON.parse(sessionStorage.getItem('ems_dispatch_queue_v1'))).toMatchObject({ items: [{ assetId: 'A2', serialNumber: 'SN-2' }] });
  });
});
