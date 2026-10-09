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

function resolveInboundBatch({ batchId = "", receivedAt, poNumber = "", supplier = "", existingRows = [] }) {
  if (batchId) {
    const row = existingRows.find((item) => String(item[0] || "").trim() === batchId);
    if (!row) {
      const error = new Error("ไม่พบล็อตที่เลือก กรุณารีเฟรชรายการล็อตแล้วลองอีกครั้ง");
      error.status = 400;
      throw error;
    }
    return { batchId: row[0], receivedAt: row[1] || receivedAt, poNumber: row[2] || "", supplier: row[3] || "", continued: true };
  }
  return {
    batchId: createBatchId(receivedAt, existingRows.map((row) => row[0])),
    receivedAt,
    poNumber: String(poNumber || "").trim(),
    supplier: String(supplier || "").trim(),
    continued: false,
  };
}

function createInboundRows({ batchId, receivedAt, poNumber = "", supplier = "", serials, assetIds, receivedBy, createdAt = new Date().toISOString() }) {
  return serials.map((serial, index) => [batchId, receivedAt, poNumber.trim(), supplier.trim(), serial, assetIds[index], receivedBy, createdAt]);
}

function inboundMap(rows = []) {
  return new Map(rows.filter((row) => row[4]).map((row) => [String(row[4]).trim(), {
    batchId: row[0] || "", receivedAt: row[1] || "", poNumber: row[2] || "", supplier: row[3] || "",
  }]));
}

function summarizeInboundBatches(rows = []) {
  const grouped = new Map();
  for (const row of rows) {
    if (!row[0]) continue;
    const batch = grouped.get(row[0]) || { batchId: row[0], receivedAt: row[1] || "", poNumber: row[2] || "", supplier: row[3] || "", count: 0 };
    batch.count += 1;
    grouped.set(row[0], batch);
  }
  return [...grouped.values()].sort((a, b) => String(b.receivedAt).localeCompare(String(a.receivedAt)));
}

function summarizeInboundPOs(inboundRows = [], assets = [], statusRows = []) {
  const assetBySerial = new Map(assets.map((asset) => [String(asset.serialNumber || asset.serial || "").trim(), asset]));
  const savedStatus = new Map(statusRows.map((row) => [String(row.poNumber || "").trim(), row]));
  const grouped = new Map();
  for (const row of inboundRows) {
    const batchId = String(row[0] || "").trim();
    const poNumber = String(row[2] || "").trim();
    if (!batchId) continue;
    const key = poNumber ? `po:${poNumber}` : `batch:${batchId}`;
    let group = grouped.get(key);
    if (!group) {
      group = { key, poNumber, batches: new Map(), serials: new Set(), receivedAt: "", supplier: "" };
      grouped.set(key, group);
    }
    group.serials.add(String(row[4] || "").trim());
    group.supplier ||= String(row[3] || "").trim();
    if (!group.receivedAt || String(row[1] || "") > group.receivedAt) group.receivedAt = String(row[1] || "");
    const batch = group.batches.get(batchId) || { batchId, receivedAt: row[1] || "", count: 0 };
    batch.count += 1;
    if (String(row[1] || "") > String(batch.receivedAt || "")) batch.receivedAt = row[1];
    group.batches.set(batchId, batch);
  }
  return [...grouped.values()].map((group) => {
    const assetsInGroup = [...group.serials].map((serial) => assetBySerial.get(serial)).filter(Boolean);
    const stockCount = assetsInGroup.filter((asset) => {
      const location = String(asset.location || "").trim().toLowerCase();
      const site = String(asset.siteName || "").trim().toLowerCase();
      return location === "stock" || location === "intranin" || site === "intranin" || site === "stock";
    }).length;
    const manual = group.poNumber ? savedStatus.get(group.poNumber) : null;
    const status = manual?.status === "closed" ? "closed" : stockCount > 0 ? "active" : "closed";
    return {
      key: group.key, poNumber: group.poNumber, batchId: [...group.batches.keys()].at(-1) || "",
      batches: [...group.batches.values()].sort((a, b) => String(b.receivedAt).localeCompare(String(a.receivedAt))),
      batchCount: group.batches.size, assetCount: group.serials.size, stockCount,
      deployedCount: Math.max(0, group.serials.size - stockCount), receivedAt: group.receivedAt,
      supplier: group.supplier, status, manuallyClosed: manual?.status === "closed",
      closedAt: manual?.closedAt || "", closedBy: manual?.closedBy || "",
    };
  }).sort((a, b) => String(b.receivedAt).localeCompare(String(a.receivedAt)));
}

module.exports = { normalizeReceivedAt, createBatchId, resolveInboundBatch, createInboundRows, inboundMap, summarizeInboundBatches, summarizeInboundPOs };
