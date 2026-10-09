import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import AddDeviceModal from '../components/AddDeviceModal.jsx';

vi.mock('axios', () => ({
  default: {
    get: vi.fn(),
    post: vi.fn(),
  },
}));

import axios from 'axios';

describe('inbound batch receiving', () => {
  beforeEach(() => {
    axios.get.mockImplementation((url) => {
      if (url === '/api/part-catalog') return Promise.resolve({ data: [{ partNumber: 'SENSOR-1', partName: 'Sensor', totalQty: 3 }] });
      if (url === '/api/inbound-batches') return Promise.resolve({ data: [{ batchId: 'INB-20261009-001', receivedAt: '2026-10-08T18:30:00.000Z', poNumber: 'PO-MORNING', supplier: 'Acme', count: 3 }] });
      return Promise.resolve({ data: [] });
    });
    axios.post.mockResolvedValue({ data: { success: true, added: 1, serials: ['SN-SENSOR-1-09102026-0004'], batchId: 'INB-20261009-001', receivedAt: '2026-10-08T18:30:00.000Z', poNumber: 'PO-MORNING', supplier: 'Acme' } });
  });

  it('lets the user search for and continue an existing batch', async () => {
    render(<AddDeviceModal open onClose={vi.fn()} onDone={vi.fn()} />);

    fireEvent.change(screen.getByPlaceholderText(/SEN\.TEMP-RS485/), { target: { value: 'SENSOR-1' } });
    fireEvent.click(await screen.findByText('Sensor', { selector: 'span' }));
    fireEvent.click(screen.getByText(/รายละเอียดเพิ่ม/));
    fireEvent.click(screen.getByRole('button', { name: 'ต่อล็อตเดิม' }));

    const batchOption = await screen.findByRole('option', { name: /INB-20261009-001/ });
    fireEvent.change(batchOption.closest('select'), { target: { value: 'INB-20261009-001' } });
    fireEvent.click(screen.getByRole('button', { name: /เพิ่ม 1 ชิ้น/ }));

    await waitFor(() => expect(axios.post).toHaveBeenCalledWith('/api/bulk-add-asset', expect.objectContaining({
      batchId: 'INB-20261009-001',
      poNumber: 'PO-MORNING',
      supplier: 'Acme',
    })));
    expect(await screen.findByText(/INB-20261009-001/)).toBeTruthy();
  });
});
