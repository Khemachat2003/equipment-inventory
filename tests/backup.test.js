// tests/backup.test.js — unit tests สำหรับ helper ของ versioned backup
const { test, describe } = require("node:test");
const assert = require("node:assert");
const { buildRunId, BACKUP_KEEP_RUNS } = require("../services/backup");

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