const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { normalizeReceivedAt, createBatchId, createInboundRows, inboundMap } = require("../utils/inbound");
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
