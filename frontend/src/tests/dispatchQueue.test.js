import { describe, expect, it } from 'vitest';
import { readDispatchQueue, writeDispatchQueue, removeDispatchedAssets, DISPATCH_QUEUE_KEY } from '../utils/dispatchQueue.js';

function memoryStorage() {
  const values = new Map();
  return {
    getItem: (key) => values.get(key) ?? null,
    setItem: (key, value) => values.set(key, value),
    removeItem: (key) => values.delete(key),
  };
}

describe('dispatch queue', () => {
  it('persists a selected PO subset for the Bundle route', () => {
    const storage = memoryStorage();
    const queue = { poNumber: 'PO-7', items: [{ assetId: 'A1', serialNumber: 'SN-1' }] };
    writeDispatchQueue(queue, storage);
    expect(storage.getItem(DISPATCH_QUEUE_KEY)).toContain('PO-7');
    expect(readDispatchQueue(storage)).toEqual(queue);
  });

  it('removes assigned serials and clears the queue when all are assembled', () => {
    const queue = { poNumber: 'PO-7', items: [{ assetId: 'A1' }, { assetId: 'A2' }] };
    expect(removeDispatchedAssets(queue, ['A1'])).toEqual({ poNumber: 'PO-7', items: [{ assetId: 'A2' }] });
    expect(removeDispatchedAssets(queue, ['A1', 'A2'])).toBeNull();
  });

  it('drops malformed or empty queue data', () => {
    const storage = memoryStorage();
    storage.setItem(DISPATCH_QUEUE_KEY, '{bad json');
    expect(readDispatchQueue(storage)).toBeNull();
    writeDispatchQueue(null, storage);
    expect(storage.getItem(DISPATCH_QUEUE_KEY)).toBeNull();
  });
});
