export function receivedDateKey(value) {
  if (!value) return '';
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Bangkok', year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(new Date(value));
  const get = (type) => parts.find((part) => part.type === type)?.value || '';
  return `${get('year')}-${get('month')}-${get('day')}`;
}

export function matchesInboundFilters(asset, { query = '', from = '', to = '' } = {}) {
  const needle = query.trim().toLowerCase();
  if (needle && ![asset.batchId, asset.poNumber, asset.supplier].some((value) => String(value || '').toLowerCase().includes(needle))) return false;
  const date = receivedDateKey(asset.receivedAt);
  if (from && (!date || date < from)) return false;
  if (to && (!date || date > to)) return false;
  return true;
}
