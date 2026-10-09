import { describe, expect, it } from 'vitest';
import { matchesInboundFilters, receivedDateKey } from '../utils/inbound.js';

describe('inbound tracking filters', () => {
  const asset = { batchId: 'INB-20261009-001', receivedAt: '2026-10-08T18:00:00.000Z', poNumber: 'PO-778', supplier: 'Acme' };

  it('uses Bangkok calendar date for received range filters', () => {
    expect(receivedDateKey(asset.receivedAt)).toBe('2026-10-09');
    expect(matchesInboundFilters(asset, { from: '2026-10-09', to: '2026-10-09' })).toBe(true);
    expect(matchesInboundFilters(asset, { from: '2026-10-10' })).toBe(false);
  });

  it('searches batch ID, PO, and supplier case insensitively', () => {
    expect(matchesInboundFilters(asset, { query: 'inb-20261009' })).toBe(true);
    expect(matchesInboundFilters(asset, { query: 'po-778' })).toBe(true);
    expect(matchesInboundFilters(asset, { query: 'acme' })).toBe(true);
    expect(matchesInboundFilters(asset, { query: 'unknown' })).toBe(false);
  });

  it('does not assign old assets to a date range or batch', () => {
    expect(matchesInboundFilters({ serialNumber: 'OLD-1' }, {})).toBe(true);
    expect(matchesInboundFilters({ serialNumber: 'OLD-1' }, { query: 'old-1' })).toBe(false);
    expect(matchesInboundFilters({ serialNumber: 'OLD-1' }, { from: '2026-10-01' })).toBe(false);
  });
});
