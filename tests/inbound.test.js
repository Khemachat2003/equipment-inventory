const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { normalizeReceivedAt, createBatchId, resolveInboundBatch, createInboundRows, inboundMap, summarizeInboundBatches, summarizeInboundPOs } = require("../utils/inbound");
const { appendInboundRows, readInboundRows } = require("../services/inboundStore");

test("normalizes inbound date and rejects invalid input", () => {
  assert.equal(normalizeReceivedAt("2026-10-09T03:00:00+07:00"), "2026-10-08T20:00:00.000Z");
  assert.throws(() => normalizeReceivedAt("not a date"), /วันที่รับเข้าไม่ถูกต้อง/);
});

test("generates sequential batch IDs using Bangkok arrival date", () => {
  const at = "2026-10-08T18:30:00.000Z";
  assert.equal(createBatchId(at), "INB-20261009-001");
  assert.equal(createBatchId(at, ["INB-20261009-001", "INB-20261009-003", "INB-20261008-008"]), "INB-20261009-004");
});

test("continues an existing batch with its stored receipt metadata", () => {
  const existingRows = [
    ["INB-20261009-001", "2026-10-08T18:30:00.000Z", "PO-MORNING", "Acme", "S1", "A1", "buyer", "2026-10-08T19:00:00.000Z"],
  ];
  assert.deepEqual(resolveInboundBatch({ batchId: "INB-20261009-001", receivedAt: "2026-10-09T08:00:00.000Z", poNumber: "PO-OTHER", supplier: "Other", existingRows }), {
    batchId: "INB-20261009-001", receivedAt: "2026-10-08T18:30:00.000Z", poNumber: "PO-MORNING", supplier: "Acme", continued: true,
  });
  assert.throws(() => resolveInboundBatch({ batchId: "INB-MISSING", receivedAt: "2026-10-09T08:00:00.000Z", existingRows }), { status: 400 });
});

test("summarizes every batch and counts each serial for all-history results", () => {
  const batches = summarizeInboundBatches([
    ["INB-OLD", "2026-10-01T00:00:00.000Z", "PO-1", "Acme", "S1"],
    ["INB-NEW", "2026-10-09T00:00:00.000Z", "PO-2", "Beta", "S2"],
    ["INB-OLD", "2026-10-01T00:00:00.000Z", "PO-1", "Acme", "S3"],
  ]);
  assert.deepEqual(batches.map(({ batchId, count }) => [batchId, count]), [["INB-NEW", 1], ["INB-OLD", 2]]);
});

test("aggregates batches into PO lifecycle and honors manual closure", () => {
  const rows = [
    ["B1", "2026-10-01", "PO-A", "Acme", "S1"],
    ["B2", "2026-10-02", "PO-A", "Acme", "S2"],
    ["B3", "2026-10-03", "PO-B", "Beta", "S3"],
  ];
  const assets = [
    { serialNumber: "S1", location: "Stock", siteName: "Intranin" },
    { serialNumber: "S2", location: "Farm A", siteName: "Farm A" },
    { serialNumber: "S3", location: "Farm B", siteName: "Farm B" },
  ];
  const groups = summarizeInboundPOs(rows, assets, [{ poNumber: "PO-A", status: "closed", closedAt: "2026-10-04" }]);
  assert.deepEqual(groups.map(({ poNumber, status, assetCount, stockCount, batchCount }) => [poNumber, status, assetCount, stockCount, batchCount]).sort((a, b) => a[0].localeCompare(b[0])), [
    ["PO-A", "closed", 2, 1, 2], ["PO-B", "closed", 1, 0, 1],
  ]);
  const reopened = summarizeInboundPOs([...rows, ["B4", "2026-10-05", "PO-A", "Acme", "S4"]], [...assets, { serialNumber: "S4", location: "Stock", siteName: "Intranin" }], []);
  assert.equal(reopened[0].status, "active");
  assert.equal(reopened[0].stockCount, 2);
});

test("creates one inbound row per serial and maps legacy assets as untracked", () => {
  const rows = createInboundRows({ batchId: "INB-20261009-001", receivedAt: "2026-10-08T18:30:00.000Z", poNumber: " PO-1 ", supplier: " Acme ", serials: ["S1", "S2"], assetIds: ["A1", "A2"], receivedBy: "buyer", createdAt: "2026-10-09T00:00:00.000Z" });
  assert.deepEqual(rows[0], ["INB-20261009-001", "2026-10-08T18:30:00.000Z", "PO-1", "Acme", "S1", "A1", "buyer", "2026-10-09T00:00:00.000Z"]);
  assert.equal(rows.length, 2);
  assert.deepEqual(inboundMap(rows).get("S2"), { batchId: "INB-20261009-001", receivedAt: "2026-10-08T18:30:00.000Z", poNumber: "PO-1", supplier: "Acme" });
  assert.equal(inboundMap([]).has("OLD-SERIAL"), false);
});

test("persists and reads inbound metadata locally when Google Sheets is unavailable", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "inbound-local-"));
  const filePath = path.join(dir, "inbound.json");
  const rows = createInboundRows({ batchId: "INB-20261009-001", receivedAt: "2026-10-08T18:30:00.000Z", serials: ["LOCAL-SN"], assetIds: ["LOCAL-ASSET"], receivedBy: "buyer" });
  const storage = await appendInboundRows(null, "", rows, filePath, async () => { throw new Error("sheet unavailable"); });
  assert.equal(storage, "local-file");
  const loaded = await readInboundRows(null, "", filePath);
  assert.equal(loaded.length, 1);
  assert.equal(loaded[0][0], "INB-20261009-001");
  assert.equal(inboundMap(loaded).get("LOCAL-SN").batchId, "INB-20261009-001");
  await appendInboundRows(null, "", createInboundRows({ batchId: "INB-20261010-001", receivedAt: "2026-10-09T18:30:00.000Z", serials: ["LOCAL-SN-2"], assetIds: ["LOCAL-ASSET-2"], receivedBy: "buyer" }), filePath, async () => { throw new Error("sheet unavailable"); });
  assert.equal((await readInboundRows(null, "", filePath)).length, 2);
  await fs.rm(dir, { recursive: true, force: true });
});
