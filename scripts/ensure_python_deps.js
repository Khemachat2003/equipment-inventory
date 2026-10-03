// scripts/ensure_python_deps.js — auto-install Python deps สำหรับ generate_label.py (ฉลาก PDF)
// เรียกที่ต้น `npm start` ก่อน node server.js เพื่อให้ทำงานทุกครั้งที่ Render boot
// (Render native node runtime ไม่คง system site-packages จาก build phase ไว้ให้ runtime)
//
// พฤติกรรม:
//  1) Fast path: ลอง `python -c "import reportlab, barcode, PIL"` → ถ้าผ่านจบทันที (~1s)
//  2) ถ้าไม่ผ่าน: pip install requirements.txt ตามลำดับ fallback
//     python -m pip → +--break-system-packages (PEP 668) → pip3 → pip3 +--break-system-packages
//  3) ติดตั้งเสร็จ verify ซ้ำ ถ้าทุกวิธีล้ม → log เตือนดัง ๆ แต่ยังปล่อย server บูตต่อ
//     (ฟีเจอร์อื่นไม่พัง — endpoint ฉลากจะ 500 พร้อม error ชัดเจนจาก routes/label.js)

const { execFileSync } = require("child_process");
const path = require("path");
const fs = require("fs");

const REQUIREMENTS = path.join(__dirname, "..", "requirements.txt");
const CHECK_SNIPPET = "import reportlab, barcode, PIL";

function run(cmd, args, timeoutMs) {
  try {
    execFileSync(cmd, args, { stdio: "inherit", timeout: timeoutMs, windowsHide: true });
    return true;
  } catch (_) {
    return false;
  }
}

// หา interpreter ที่ใช้ได้ — เช็คด้วย `-c "pass"` เพราะบน Windows อาจเจอ
// Microsoft Store alias stub ที่คืน exit code แปลก ๆ
function findPython() {
  for (const candidate of ["python", "python3"]) {
    if (run(candidate, ["-c", "pass"], 30000)) return candidate;
  }
  return null;
}

function main() {
  console.log("[ensure-python-deps] checking python deps (reportlab, barcode, PIL)...");

  if (!fs.existsSync(REQUIREMENTS)) {
    console.error(`[ensure-python-deps] ❌ ไม่พบ ${REQUIREMENTS} — ข้ามการติดตั้ง`);
    return;
  }

  const py = findPython();
  if (!py) {
    console.warn(
      "[ensure-python-deps] ⚠️ ไม่พบ python/python3 บนเครื่องนี้ — " +
        "endpoint ฉลาก PDF (/api/label/*) จะใช้ไม่ได้ แต่ server จะบูตต่อตามปกติ"
    );
    return;
  }

  // Fast path — deps ครบอยู่แล้ว (เช่นตอนรัน local หรือ instance ที่เคย install แล้ว)
  if (run(py, ["-c", CHECK_SNIPPET], 60000)) {
    console.log("[ensure-python-deps] ✅ python deps OK (fast path)");
    return;
  }

  console.log("[ensure-python-deps] deps ยังไม่ครบ — เริ่ม pip install (ครั้งแรกของ instance ใช้เวลา ~15-30s)");

  const installs = [
    [py, ["-m", "pip", "install", "--no-cache-dir", "-r", REQUIREMENTS]],
    [py, ["-m", "pip", "install", "--no-cache-dir", "--break-system-packages", "-r", REQUIREMENTS]],
    ["pip3", ["install", "--no-cache-dir", "-r", REQUIREMENTS]],
    ["pip3", ["install", "--no-cache-dir", "--break-system-packages", "-r", REQUIREMENTS]],
  ];

  let installed = false;
  for (const [cmd, args] of installs) {
    console.log(`[ensure-python-deps] trying: ${cmd} ${args.join(" ")}`);
    if (run(cmd, args, 600000)) { installed = true; break; }
    console.warn(`[ensure-python-deps] command failed, ลองวิธีถัดไป...`);
  }

  // verify ซ้ำก่อนปล่อยผ่าน
  if (installed && run(py, ["-c", CHECK_SNIPPET], 60000)) {
    console.log("[ensure-python-deps] ✅ python deps ติดตั้งและ verify ผ่าน");
    return;
  }

  console.error(
    "[ensure-python-deps] ❌❌ ติดตั้ง/verify python deps ไม่สำเร็จทุกวิธี — " +
      "endpoint ฉลาก PDF (/api/label/pdf, /api/label/a4pdf) จะคืน 500 " +
      "ดู error จาก pip ด้านบน แล้วพิจารณาแก้ Build Command ใน Render Dashboard"
  );
  // ไม่ exit(1) — ปล่อยให้ server บูตต่อ ฟีเจอร์อื่นยังใช้ได้
}

main();
