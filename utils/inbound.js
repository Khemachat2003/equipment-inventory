function normalizeReceivedAt(value, now = new Date()) {
  if (!value) return now.toISOString();
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) throw new Error("วันที่รับเข้าไม่ถูกต้อง");
  return parsed.toISOString();
}

function createBatchId(receivedAt, existingIds = []) {
  const date = new Date(receivedAt);
  if (Number.isNaN(date.getTime())) throw new Error("วันที่รับเข้าไม่ถูกต้อง");
  const ymd = new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Bangkok", year: "numeric", month: "2-digit", day: "2-digit" }).format(date).replaceAll("-", "");
  const prefix = `INB-${ymd}-`;
  const max = existingIds.reduce((n, id) => {
    const match = String(id || "").match(new RegExp(`^INB-${ymd}-(\\d+)$`));
    return match ? Math.max(n, Number(match[1]) || 0) : n;
  }, 0);
  return `${prefix}${String(max + 1).padStart(3, "0")}`;
}

function createInboundRows({ batchId, receivedAt, poNumber = "", supplier = "", serials, assetIds, receivedBy, createdAt = new Date().toISOString() }) {
  return serials.map((serial, index) => [batchId, receivedAt, poNumber.trim(), supplier.trim(), serial, assetIds[index], receivedBy, createdAt]);
}

function inboundMap(rows = []) {
  return new Map(rows.filter((row) => row[4]).map((row) => [String(row[4]).trim(), {
    batchId: row[0] || "", receivedAt: row[1] || "", poNumber: row[2] || "", supplier: row[3] || "",
  }]));
}

module.exports = { normalizeReceivedAt, createBatchId, createInboundRows, inboundMap };
