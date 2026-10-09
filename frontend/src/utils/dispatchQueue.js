export const DISPATCH_QUEUE_KEY = 'ems_dispatch_queue_v1';

export function readDispatchQueue(storage = globalThis.sessionStorage) {
  try {
    const value = JSON.parse(storage?.getItem(DISPATCH_QUEUE_KEY) || 'null');
    if (!value || !Array.isArray(value.items)) return null;
    const items = value.items.filter((item) => item?.assetId && item?.serialNumber);
    return items.length ? { poNumber: String(value.poNumber || ''), items } : null;
  } catch { return null; }
}

export function writeDispatchQueue(queue, storage = globalThis.sessionStorage) {
  if (!storage) return;
  if (!queue?.items?.length) storage.removeItem(DISPATCH_QUEUE_KEY);
  else storage.setItem(DISPATCH_QUEUE_KEY, JSON.stringify({ ...queue, items: queue.items }));
}

export function removeDispatchedAssets(queue, assetIds) {
  if (!queue) return null;
  const removed = new Set(assetIds);
  const items = queue.items.filter((item) => !removed.has(item.assetId));
  return items.length ? { ...queue, items } : null;
}
