// vite-plugin-icon-sync.js — ทำให้ icon subset font อัปเดตเองอัตโนมัติ (dev + build)
// ─────────────────────────────────────────────────────────────────────────────
// ที่มา: subset font + iconCodepoints.js เป็นไฟล์ generate จาก
// `python scripts/make_icon_subset.py` — เดิมต้องรันมือทุกครั้งที่เพิ่ม
// <Icon name="..."> ชื่อใหม่ ไม่งั้น icon ไม่แสดง
//
// plugin นี้จัดการให้เอง:
//   1) scan ชื่อ icon จาก src/**/*.{js,jsx} (regex ชุดเดียวกับใน .py — แก้ที่ไหนต้องแก้อีกที่)
//   2) เทียบกับ iconCodepoints.js และ .icon-skip-cache.json (ชื่อ string ที่ไม่ใช่ icon
//      เช่น 'axios' ถูก .py จดไว้ระหว่าง regen ล่าสุด → ไม่เกิด regen ต่อไม่รู้จบ)
//   3) เจอชื่อใหม่ที่ยังไม่ถูก map → รัน make_icon_subset.py ให้เอง
//      - dev: regen เสร็จสั่ง full-reload → เบราว์เซอร์ได้ subset ใหม่ + codepoints ใหม่ทันที
//   4) python ไม่มี/ล้มเหลว → เตือนแล้วปล่อย dev/build ต่อ (Icon.jsx มี full-font fallback ครอบ)
//
// หมายเหตุ deploy: บน Render buildCommand pip install fonttools → ./python_modules
// เราเลยต่อ PYTHONPATH=./python_modules ให้ subprocess เพื่อให้ regen ใช้ได้ตอน build ด้วย

import { spawn, spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const FRONTEND_ROOT = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(FRONTEND_ROOT, '..');
const SRC_DIR = path.join(FRONTEND_ROOT, 'src');
const CODEPOINTS_FILE = path.join(SRC_DIR, 'data', 'iconCodepoints.js');
const SKIP_CACHE_FILE = path.join(FRONTEND_ROOT, '.icon-skip-cache.json');
const SUBSET_SCRIPT = path.join(PROJECT_ROOT, 'scripts', 'make_icon_subset.py');

// ── regex ชุดเดียวกับ PATTERNS + GLOBAL_LITERAL_RE ใน scripts/make_icon_subset.py ──
const PATTERNS = [
  /<Icon[^>]*\bname="([a-z_0-9]+)"/g,
  /\bname=\{["'`]([a-z_0-9]+)["'`]\}/g,
  /\bicon:\s*['"]([a-z_0-9]+)['"]/g,
  /\bicon="([a-z_0-9]+)"/g,
];
const GLOBAL_LITERAL_RE = /['"`]([a-z_0-9]{2,})['"`]/g;

function collect(names, text, re) {
  re.lastIndex = 0;
  let m;
  while ((m = re.exec(text)) !== null) names.add(m[1]);
}

function walkSrc(dir, out = []) {
  if (!fs.existsSync(dir)) return out;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) walkSrc(full, out);
    else if (/\.(jsx|js)$/.test(entry.name)) out.push(full);
  }
  return out;
}

function scanIconNames() {
  const names = new Set();
  for (const file of walkSrc(SRC_DIR)) {
    let text = '';
    try {
      text = fs.readFileSync(file, 'utf8');
    } catch {
      continue;
    }
    for (const re of PATTERNS) collect(names, text, re);
    collect(names, text, GLOBAL_LITERAL_RE);
  }
  return names;
}

// ชื่อ icon ที่ map ไว้แล้ว (จากไฟล์ generate — parse ด้วย regex พอ ไม่ต้อง import จริง)
function readMappedNames() {
  try {
    const text = fs.readFileSync(CODEPOINTS_FILE, 'utf8');
    return new Set([...text.matchAll(/'([a-z_0-9]+)':\s*0x[0-9a-fA-F]+/g)].map((m) => m[1]));
  } catch {
    return new Set();
  }
}

// ชื่อที่ .py ยืนยันแล้วว่า "ไม่ใช่ glyph จริงของฟอนต์" → ไม่ต้อง regen ซ้ำเพราะชื่อพวกนี้
function readSkippedNames() {
  try {
    return new Set(JSON.parse(fs.readFileSync(SKIP_CACHE_FILE, 'utf8')));
  } catch {
    return new Set();
  }
}

function findMissingIcons() {
  const mapped = readMappedNames();
  const skipped = readSkippedNames();
  return [...scanIconNames()].filter((n) => !mapped.has(n) && !skipped.has(n)).sort();
}

function findPython() {
  // เช็คด้วย `-c "pass"` กัน Windows Store alias stub ที่คืน exit code แปลกๆ
  for (const candidate of ['python', 'python3']) {
    const probe = spawnSync(candidate, ['-c', 'pass'], { windowsHide: true, timeout: 15000 });
    if (probe.status === 0) return candidate;
  }
  return null;
}

function runSubsetScript() {
  const py = findPython();
  if (!py) {
    console.warn(
      '[icon-sync] ⚠️ ไม่พบ python บนเครื่องนี้ — ข้ามการ regen subset ' +
        '(icon ใหม่ยังแสดงผ่าน full-font fallback ของ <Icon> อยู่)'
    );
    return Promise.resolve(false);
  }
  const env = { ...process.env };
  // บน Render: fonttools ถูก pip install ลง ./python_modules ตอน build → ต่อ PYTHONPATH ให้ import เจอ
  const pyModules = path.join(PROJECT_ROOT, 'python_modules');
  if (fs.existsSync(pyModules)) {
    env.PYTHONPATH = env.PYTHONPATH ? `${pyModules}${path.delimiter}${env.PYTHONPATH}` : pyModules;
  }
  return new Promise((resolve) => {
    const child = spawn(py, [SUBSET_SCRIPT], {
      cwd: PROJECT_ROOT,
      env,
      windowsHide: true,
      stdio: 'inherit',
    });
    const timer = setTimeout(() => {
      try {
        child.kill();
      } catch {
        /* ปล่อยได้ */
      }
      resolve(false);
    }, 180000);
    child.on('close', (code) => {
      clearTimeout(timer);
      resolve(code === 0);
    });
    child.on('error', () => {
      clearTimeout(timer);
      resolve(false);
    });
  });
}

export default function iconSyncPlugin() {
  let devServer = null;
  let debounceTimer = null;
  let pythonBroken = false; // python ใช้ไม่ได้ → หยุด retry กัน log รก (restart dev server เพื่อลองใหม่)
  let running = false;

  async function syncNow(cause) {
    if (pythonBroken || running) return false;
    running = true;
    try {
      const missing = findMissingIcons();
      if (!missing.length) return false;
      console.log(`[icon-sync] เจอ icon ใหม่ที่ยังไม่อยู่ใน subset (${cause}): ${missing.join(', ')}`);
      console.log('[icon-sync] กำลัง regen subset font + iconCodepoints.js …');
      const ok = await runSubsetScript();
      if (!ok) {
        pythonBroken = true;
        console.warn(
          '[icon-sync] ❌ regen ไม่สำเร็จ — ข้ามไปก่อน (icon ใหม่ยังแสดงผ่าน full-font fallback ของ <Icon>) ' +
            'แก้เรียบร้อยแล้วให้ restart dev server เพื่อเปิด auto-regen อีกครั้ง'
        );
        return false;
      }
      console.log('[icon-sync] ✅ subset font + iconCodepoints.js อัปเดตแล้ว');
      if (devServer) {
        try {
          if (devServer.moduleGraph?.invalidateAll) await devServer.moduleGraph.invalidateAll();
        } catch {
          /* บางเวอร์ชันไม่มี — full-reload อย่างเดียวก็พอ */
        }
        devServer.ws.send({ type: 'full-reload' });
      }
      return true;
    } finally {
      running = false;
    }
  }

  return {
    name: 'icon-sync',
    enforce: 'pre',
    // รับช่วงก่อน build/transform เสมอ → iconCodepoints.js ล่าสุดถูก bundle เข้าไปเสมอ
    async buildStart() {
      await syncNow('startup');
    },
    configureServer(server) {
      devServer = server;
      const onSrcChange = (file) => {
        if (pythonBroken) return;
        const rel = path.relative(SRC_DIR, file).split(path.sep).join('/');
        if (rel.startsWith('..') || !/\.(jsx|js)$/.test(rel)) return;
        if (rel === 'data/iconCodepoints.js') return; // ไฟล์ที่ script generate เอง — ไม่ให้ loop
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(() => {
          syncNow('file change').catch(() => {});
        }, 800);
      };
      server.watcher.on('change', onSrcChange);
      server.watcher.on('add', onSrcChange);
    },
  };
}
