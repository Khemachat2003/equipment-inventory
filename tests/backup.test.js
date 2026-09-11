// tests/backup.test.js — unit tests สำหรับ helper ของ versioned backup
const { test, describe } = require("node:test");
const assert = require("node:assert");
const { buildRunId, BACKUP_KEEP_RUNS, ensureBackupTableSchema } = require("../services/backup");

// ── fake pg client (จับ query ที่ถูกเรียก ไม่ต่อ DB จริง) ──
function fakeClient(regResult) {
  const queries = [];
  return {
    queries,
    async query(sql, params) {
      queries.push({ sql, params });
      // จำลอง to_regclass(): คืนชื่อตารางถ้า "มีอยู่" / null ถ้าไม่มี
      return { rows: [{ reg: regResult }], rowCount: 1 };
    },
  };
}

describe("ensureBackupTableSchema (self-heal)", () => {
  test("ตารางมีอยู่ → เช็คด้วย to_regclass แล้ว ALTER เติม backup_run_id", async () => {
    const client = fakeClient("backup_users");
    const ok = await ensureBackupTableSchema(client, "backup_users");
    assert.strictEqual(ok, true);
    assert.ok(client.queries[0].sql.includes("to_regclass"));
    assert.strictEqual(client.queries[0].params[0], "backup_users");
    assert.ok(client.queries[1].sql.includes("ADD COLUMN IF NOT EXISTS backup_run_id"));
  });

  test("ยังไม่มีตาราง (ยังไม่เคย backup) → คืน false และไม่ ALTER", async () => {
    const client = fakeClient(null);
    const ok = await ensureBackupTableSchema(client, "backup_users");
    assert.strictEqual(ok, false);
    assert.strictEqual(client.queries.length, 1); // มีแค่ to_regclass
  });

  test("ชื่อตารางแปลกปลอม/SQL injection → throw (กันชั้นใน)", async () => {
    await assert.rejects(
      () => ensureBackupTableSchema(fakeClient(null), "evil; DROP TABLE users"),
      /Invalid backup table/
    );
    await assert.rejects(
      () => ensureBackupTableSchema(fakeClient(null), "not_backup_table"),
      /Invalid backup table/
    );
  });
});

describe("buildRunId", () => {
  test("รูปแบบ YYYYMMDDHHmmss = 14 หลักตัวเลข", () => {
    assert.match(buildRunId(), /^\d{14}$/);
  });

  test("ใช้วันที่ที่กำหนดได้ (pad เลขเดี่ยวด้วย 0)", () => {
    const d = new Date(2026, 8, 11, 9, 5, 3); // 2026-09-11 09:05:03
    assert.strictEqual(buildRunId(d), "20260911090503");
  });

  test("run id ใหม่ > run id เก่า (เรียงเป็น string ได้ → เทียบ max() ถูกต้อง)", () => {
    const newer = buildRunId(new Date(2026, 8, 11, 18, 0, 0));
    const older = buildRunId(new Date(2026, 8, 11, 9, 0, 0));
    assert.ok(newer > older);
  });
});

describe("BACKUP_KEEP_RUNS", () => {
  test("ค่าเริ่มต้นต้อง >= 1 (เก็บ backup ล่าสุดอย่างน้อย 1 run)", () => {
    assert.ok(BACKUP_KEEP_RUNS >= 1);
  });
});