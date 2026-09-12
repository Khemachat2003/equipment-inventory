#!/usr/bin/env python3
"""สร้าง subset ของ Material Symbols Rounded font เฉพาะ icon ที่ใช้ในโปรเจค
ทำให้ font จาก ~5.1 MB เล็กลงเหลือระดับ ~100 KB → โหลดเร็ว ไม่เห็นชื่อ icon แวบ

ใช้: python scripts/make_icon_subset.py
- อ่าน icon ทั้งหมดจาก frontend/src (รูปแบบ <Icon name="..", icon: '..', icon="..")
- ตัดจาก full font ที่ fonts/material-symbols-rounded-full.woff2
- เขียนผลไปที่ frontend/public/fonts/material-symbols-rounded.subset.woff2
ต้องติดตั้ง: pip install fonttools brotli (ดู requirements.txt)
"""
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT / "frontend" / "src"
FULL_FONT = ROOT / "fonts" / "material-symbols-rounded-full.woff2"
OUT_FONT = ROOT / "frontend" / "public" / "fonts" / "material-symbols-rounded.subset.woff2"

# icon ที่ใช้โดยไม่ได้อยู่ใน <Icon name=".."> (spinner ของ LoadingButton เป็นต้น)
EXTRA_ICONS = {
    "progress_activity",
    "autorenew",
    "hourglass_empty",
    "sync",
    "cancel",
}

PATTERNS = [
    re.compile(r'''<Icon[^>]*\bname="([a-z_0-9]+)"'''),
    re.compile(r'''\bname=\{["'`]([a-z_0-9]+)["'`]\}'''),
    re.compile(r'''\bicon:\s*['"]([a-z_0-9]+)['"]'''),
    re.compile(r'''\bicon="([a-z_0-9]+)"'''),
]


def collect_icon_names() -> set:
    names = set(EXTRA_ICONS)
    files = list(SRC_DIR.rglob("*.jsx")) + list(SRC_DIR.rglob("*.js"))
    for f in files:
        text = f.read_text(encoding="utf-8", errors="ignore")
        for pat in PATTERNS:
            names.update(pat.findall(text))
    return names


def main() -> int:
    if not FULL_FONT.exists():
        print(f"❌ ไม่พบ full font: {FULL_FONT}")
        return 1
    names = collect_icon_names()
    # ชื่อ icon ตรงกับ glyph name ใน full font (เช่น 'inventory_2')
    # → ระบุ --glyphs ตรงๆ เพื่อให้ ligature glyph ถูกเก็บมาด้วย
    #   (--text อย่างเดียวทำ closure ของ GSUB ligature ไม่ครบ → icon จะว่างเปล่า!)
    glyphs = ",".join(sorted(names))
    text = " ".join(sorted(names))
    OUT_FONT.parent.mkdir(parents=True, exist_ok=True)

    cmd = [
        sys.executable, "-m", "fontTools.subset",
        str(FULL_FONT),
        f"--glyphs={glyphs}",
        f"--text={text}",
        "--layout-features=liga,calt",
        "--flavor=woff2",
        f"--output-file={OUT_FONT}",
    ]
    subprocess.run(cmd, check=True)

    # ── สร้าง map ชื่อ icon → codepoint จาก cmap ของ full font ──
    # font subset ที่ได้ (woff2) ถูกลบ glyph names ทิ้ง → การแสดงผลผ่านชื่อ
    # (ligature) จึงพัง → Icon component ใช้ codepoint ตรงๆ จากไฟล์นี้แทน
    from fontTools.ttLib import TTFont  # noqa: E402

    full = TTFont(str(FULL_FONT))
    cmap = full.getBestCmap()
    # reverse cmap: glyph name → codepoint
    rev = {g: cp for cp, g in cmap.items()}
    codepoints = {}
    for name in sorted(names):
        cp = rev.get(name)
        if cp is not None:
            codepoints[name] = f"{cp:04x}"

    out_js = ROOT / "frontend" / "src" / "data" / "iconCodepoints.js"
    out_js.parent.mkdir(parents=True, exist_ok=True)
    lines = ["// AUTO-GENERATED โดย scripts/make_icon_subset.py — ห้ามแก้มือ",
             "// map ชื่อ icon → codepoint (มาจาก cmap ของ Material Symbols full font)",
             "export const ICON_CODEPOINTS = {"]
    for name in sorted(codepoints):
        lines.append(f"  '{name}': 0x{codepoints[name]},")
    lines.append("};")
    lines.append(f"// {len(codepoints)}/{len(names)} icons mapped")
    out_js.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"   iconCodepoints.js: {len(codepoints) if 'codepoints' in dir() else '?'}")

    # ── VERIFY สุดท้าย: codepoint ของทุก icon ต้องอยู่ใน cmap ของ subset ──
    sub = TTFont(str(OUT_FONT))
    sub_cmap = sub.getBestCmap()
    missing = [
        name for name, cp in codepoints.items() if int(cp, 16) not in sub_cmap
    ]
    if missing:
        print(f"❌ VERIFY FAILED: {len(missing)} icons missing in subset cmap: {missing[:10]}")
        return 1
    unmapped = sorted(set(names) - set(codepoints))
    if unmapped:
        print(f"⚠️  icons ที่ไม่มี codepoint ใน full font (จะ fallback แสดงชื่อ): {unmapped}")
    size_kb = OUT_FONT.stat().st_size // 1024
    full_kb = FULL_FONT.stat().st_size // 1024
    print(f"✅ VERIFY OK: {len(codepoints)} icons พร้อมใช้ | {size_kb} KB (จาก {full_kb} KB)")
    print("   ถ้าเพิ่ม icon ใหม่ในโค้ด ให้รันสคริปต์นี้ซ้ำเพื่อ regen subset")
    return 0

    size_kb = OUT_FONT.stat().st_size // 1024
    full_kb = FULL_FONT.stat().st_size // 1024
    print(f"✅ subset {len(names)} icons -> {OUT_FONT.name} ({size_kb} KB, จาก {full_kb} KB)")
    print("   ถ้าเพิ่ม icon ใหม่ในโค้ด ให้รันสคริปต์นี้ซ้ำเพื่อ regen subset")
    return 0


if __name__ == "__main__":
    sys.exit(main())
