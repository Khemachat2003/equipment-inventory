// scripts/ensure_python_deps.js — auto-install Python deps สำหรับ generate_label.py (ฉลาก PDF)
// เรียกที่ต้น `npm start` ก่อน node server.js เพื่อให้ทำงานทุกครั้งที่ Render boot
//
// ทำไมต้องมีสคริปต์นี้: Render native node runtime ไม่คง system site-packages
// จาก build phase ไว้ให้ runtime → ชั้นหลักคือ buildCommand ใน render.yaml ที่
// ติดตั้งลง ./python_modules (path "ใน repo" คงอยู่ถึง runtime เหมือน node_modules)
// ส่วนสคริปต์นี้เป็นชั้น self-heal ตอน boot:
//
// พฤติกรรม:
//  1) Fast path: ลอง `python -c "import reportlab, barcode, PIL"` (PYTHONPATH ชี้ ./python_modules)
//     → ถ้าผ่านจบทันที (~1s)
//  2) ถ้าไม่ผ่าน: pip install ตามลำดับ fallback — ลง ./python_modules ก่อน 4 วิธี
//     (python -m pip → +--break-system-packages (PEP 668) → pip3 → pip3 +--break-system-packages)
//     แล้วลง system เป็นชั้นสำรองอีก 4 วิธีแบบเดียวกัน
//  3) ติดตั้งเสร็จ verify ซ้ำ ถ้าทุกวิธีล้ม → log เตือนดัง ๆ แต่ยังปล่อย server บูตต่อ
//     (ฟีเจอร์อื่นไม่พัง — endpoint ฉลากจะ 500 พร้อม error ชัดเจนจาก routes/label.js)

const { execFileSync } = require("child_process");
const path = require("path");
const fs = require("fs");

const REQUIREMENTS = path.join(__dirname, "..", "requirements.txt");
const TARGET_DIR = path.join(__dirname, "..", "python_modules"); // path ใน repo — คงอยู่ถึง runtime
const CHECK_SNIPPET = "import reportlab, barcode, PIL";

function run(cmd, args, timeoutMs, env) {
  try {
    execFileSync(cmd, args, { stdio: "inherit", timeout: timeoutMs, windowsHide: true, env });
    return true;
  } catch (_) {
    return false;
  }
}

// env สำหรับ python subprocess — PYTHONPATH ชี้ ./python_modules เพื่อ import จาก repo
// (ถ้ายังไม่มี dir นี้ ก็ปล่อยว่าง — เดี๋ยว import หาจาก system site-packages)
function pyEnv() {
  const env = Object.assign({}, process.env);
  const parts = [];
  if (fs.existsSync(TARGET_DIR)) parts.push(TARGET_DIR);
  if (env.PYTHONPATH) parts.push(env.PYTHONPATH);
  if (parts.length) env.PYTHONPATH = parts.join(path.delimiter);
  return env;
}

function checkImports(py) {
  return run(py, ["-c", CHECK_SNIPPET], 60000, pyEnv());
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

  // Fast path — deps ครบอยู่แล้ว (เช่น ./python_modules จาก build phase หรือ system site-packages)
  if (checkImports(py)) {
    console.log("[ensure-python-deps] ✅ python deps OK (fast path)");
    return;
  }

  console.log("[ensure-python-deps] deps ยังไม่ครบ — เริ่ม pip install (ครั้งแรกของ instance ใช้เวลา ~15-60s)");

  const reqs = ["-r", REQUIREMENTS];
  // ลง ./python_modules ก่อน (คงอยู่ถึง runtime) — system เป็นชั้นสำรอง
  const installs = [
    [py, ["-m", "pip", "install", "--no-cache-dir", "--upgrade", "--target", TARGET_DIR, ...reqs]],
    [py, ["-m", "pip", "install", "--no-cache-dir", "--upgrade", "--break-system-packages", "--target", TARGET_DIR, ...reqs]],
    ["pip3", ["install", "--no-cache-dir", "--upgrade", "--target", TARGET_DIR, ...reqs]],
    ["pip3", ["install", "--no-cache-dir", "--upgrade", "--break-system-packages", "--target", TARGET_DIR, ...reqs]],
    // system (ชั้นสำรอง — install ได้แต่ไม่คงอยู่ระหว่าง build→runtime บน Render)
    [py, ["-m", "pip", "install", "--no-cache-dir", ...reqs]],
    [py, ["-m", "pip", "install", "--no-cache-dir", "--break-system-packages", ...reqs]],
    ["pip3", ["install", "--no-cache-dir", ...reqs]],
    ["pip3", ["install", "--no-cache-dir", "--break-system-packages", ...reqs]],
  ];

  let ok = false;
  for (const [cmd, args] of installs) {
    console.log(`[ensure-python-deps] trying: ${cmd} ${args.join(" ")}`);
    if (run(cmd, args, 600000)) {
      // verify ทันที — ถ้า pip exit 0 แต่ import ยังไม่ผ่าน ให้ลองวิธีถัดไป
      if (checkImports(py)) { ok = true; break; }
      console.warn("[ensure-python-deps] install ผ่านแต่ verify ยังไม่ผ่าน — ลองวิธีถัดไป...");
    } else {
      console.warn("[ensure-python-deps] command failed, ลองวิธีถัดไป...");
    }
  }

  if (ok) {
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
