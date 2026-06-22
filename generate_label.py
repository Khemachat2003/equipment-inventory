#!/usr/bin/env python3
"""
Industrial Barcode Label Generator (Production Grade - Fully Fixed)
==================================================================
- Absolute Grid with fixed margins (no auto-centering)
- Fixed Module Width 0.28mm with safe fallback
- High-Contrast Light/Dark Profile
- Only Barcode + Serial Number
- In-Memory BytesIO Processing (No disk permission errors)
- Transparent mask disabled (Fixes invisible barcode lines)
"""

import argparse
import json
import sys
import os
from io import BytesIO

# Force UTF-8 output on Windows (avoids CP1252 UnicodeEncodeError)
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader

import barcode
from barcode.writer import ImageWriter
from PIL import Image

# ─── Font Setup ───
FONT_MONO = 'Courier-Bold'

def try_register_font():
    paths = [
        '/usr/share/fonts/truetype/thsarabun/THSarabunNew.ttf',
        '/usr/share/fonts/THSarabunNew.ttf',
        os.path.join(os.path.dirname(__file__), 'fonts', 'THSarabunNew.ttf'),
    ]
    for p in paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont('THSarabunNew', p))
                return 'THSarabunNew'
            except Exception:
                pass
    return FONT_MONO

THAI_FONT = try_register_font()

# ─── Constants ───
DEFAULT_MODULE_WIDTH_MM = 0.28
MIN_MODULE_WIDTH_MM = 0.26
QUIET_ZONE_MM = 3.5
BARCODE_DPI = 300  # 300 DPI for print quality

# ─── Barcode Generator (Fixed: Using BytesIO to avoid OS file locks) ───
def make_barcode_image(data, height_mm, module_width_mm=DEFAULT_MODULE_WIDTH_MM, 
                       quiet_zone_mm=QUIET_ZONE_MM, bg='#ffffff', fg='#000000'):
    """
    สร้างบาร์โค้ด CODE128 ลงบนหน่วยความจำโดยตรง ไม่เขียนลงดิสก์
    """
    opts = {
        'module_width': module_width_mm,
        'module_height': height_mm,
        'quiet_zone': quiet_zone_mm,
        'font_size': 0,
        'write_text': False,
        'background': bg,
        'foreground': fg,
        'dpi': BARCODE_DPI,
    }
    
    # ใช้ BytesIO ประมวลผลบนแรมแทนระบบไฟล์ชั่วคราว เพื่อความเสถียรและเสร็จไวขึ้น
    fp = BytesIO()
    code = barcode.get('code128', data, writer=ImageWriter())
    code.write(fp, options=opts)
    fp.seek(0)
    
    img = Image.open(fp).convert('RGB')
    
    # คำนวณขนาดจริง (mm) จากพิกเซลและ DPI
    width_px, height_px = img.size
    width_mm = (width_px / BARCODE_DPI) * 25.4
    height_mm = (height_px / BARCODE_DPI) * 25.4
    
    return img, width_mm, height_mm


def render_label_absolute(
    c,
    abs_x_mm,
    abs_y_mm,
    w_mm,
    h_mm,
    serial,
    theme='light',
    font_size=12,
    page_h_mm=297.0
):
    """
    วาดป้ายสินค้าด้วยพิกัดแม่นยำสูง (Absolute Grid Matrix)
    """
    # แปลงเป็นพิกัด ReportLab (origin ที่มุมล่างซ้าย)
    x = abs_x_mm * mm
    y = (page_h_mm - abs_y_mm - h_mm) * mm
    w = w_mm * mm
    h = h_mm * mm

    inner_margin = 1.8 * mm
    inner_w = w - 2 * inner_margin
    inner_h = h - 2 * inner_margin

    # ─── จัดแบ่งสัดส่วนพื้นที่: บาร์โค้ดด้านบน 60%, ข้อความด้านล่าง 40% ───
    bc_h = inner_h * 0.60
    txt_h = inner_h - bc_h
    
    # พิกัดฐานล่างสุดของพื้นที่บาร์โค้ด
    bc_bottom_y = y + inner_margin + txt_h

    # ─── พื้นหลังและการตั้งค่าธีม (Fixed: พิกัดขอบเขตกล่องขาว) ───
    if theme == 'dark':
        # พื้นหลังดำทั้งดวง
        c.setFillColor(colors.black)
        c.rect(x, y, w, h, fill=1, stroke=0)

        # White Isolation Box สำหรับบาร์โค้ด (กว้างและสูงตรงตามกรอบพื้นที่ทำงาน)
        box_x = x + inner_margin
        box_y = bc_bottom_y
        box_w = inner_w
        box_h = bc_h

        c.setFillColor(colors.white)
        c.rect(box_x, box_y, box_w, box_h, fill=1, stroke=0)
        c.setStrokeColor(colors.HexColor('#dddddd'))
        c.setLineWidth(0.3)
        c.rect(box_x, box_y, box_w, box_h, fill=0, stroke=1)

        bc_bg = '#ffffff'
        bc_fg = '#000000'
        text_color = colors.white

    else:  # light
        c.setFillColor(colors.white)
        c.rect(x, y, w, h, fill=1, stroke=0)
        bc_bg = '#ffffff'
        bc_fg = '#000000'
        text_color = colors.black

    # ─── กรอบป้ายด้านนอกสุด ───
    c.setStrokeColor(colors.HexColor('#666666'))
    c.setLineWidth(0.4)
    c.rect(x, y, w, h, fill=0, stroke=1)

    # ─── เตรียมข้อมูลสำหรับบาร์โค้ด (ตัดเซกเมนต์ท้ายมาทำบาร์โค้ด) ───
    parts = serial.split('-')
    bc_data = '-'.join(parts[-2:]) if len(parts) >= 2 else serial

    # ─── สร้างและวาดบาร์โค้ด ───
    try:
        img, barcode_width_mm, barcode_height_mm = make_barcode_image(
            bc_data,
            bc_h / mm,  # ความสูงที่ต้องการ (mm)
            DEFAULT_MODULE_WIDTH_MM,
            QUIET_ZONE_MM,
            bg=bc_bg,
            fg=bc_fg
        )
        
        # ตรวจสอบว่าบาร์โค้ดกว้างเกินป้ายหรือไม่
        max_width_mm = w_mm - 2 * (inner_margin / mm)
        
        if barcode_width_mm > max_width_mm:
            ratio = max_width_mm / barcode_width_mm
            new_module_width = DEFAULT_MODULE_WIDTH_MM * ratio
            
            if new_module_width < MIN_MODULE_WIDTH_MM:
                print(f'[WARN] Module width would be {new_module_width:.3f}mm, using minimum {MIN_MODULE_WIDTH_MM}mm')
                new_module_width = MIN_MODULE_WIDTH_MM
            
            img2, barcode_width_mm, barcode_height_mm = make_barcode_image(
                bc_data,
                bc_h / mm,
                new_module_width,
                QUIET_ZONE_MM,
                bg=bc_bg,
                fg=bc_fg
            )
            img = img2

        # กำหนดขนาดพิกัดจุดวาดภาพให้อยู่ในกรอบอย่างสมบูรณ์
        dw = min(barcode_width_mm * mm, inner_w)
        dh = bc_h
        dx = x + (w - dw) / 2
        dy = bc_bottom_y

        # Fixed: เปลี่ยนจาก mask='auto' เป็น mask=None เพื่อป้องกันบาร์โค้ดเส้นดำหายโปร่งแสง
        c.drawImage(ImageReader(img), dx, dy, dw, dh, preserveAspectRatio=False, mask=None)
        
    except Exception as e:
        print(f'[WARN] Barcode render error: {e}', file=sys.stderr)

    # ─── เส้นแบ่งกลางระว่าง บาร์โค้ด กับ ข้อความ ───
    c.setStrokeColor(colors.HexColor('#aaaaaa'))
    c.setLineWidth(0.5)
    c.line(x + inner_margin, bc_bottom_y, x + w - inner_margin, bc_bottom_y)

    # ─── ส่วนแสดงผลข้อความ Serial Number (Auto-fit font) ───
    max_text_w = w - 2 * inner_margin
    fit_size = font_size
    c.setFont(THAI_FONT, fit_size)
    while fit_size > 5 and c.stringWidth(serial, THAI_FONT, fit_size) > max_text_w:
        fit_size -= 0.5
    c.setFillColor(text_color)
    c.setFont(THAI_FONT, fit_size)
    c.drawCentredString(x + w / 2, y + inner_margin + (txt_h * 0.35), serial)


def generate_single(serial, w, h, theme, font_size, out):
    """PDF ขนาด label เดียว (สำหรับ Label Printer)"""
    c = canvas.Canvas(out, pagesize=(w * mm, h * mm))
    render_label_absolute(c, 0, 0, w, h, serial, theme, font_size, h)
    c.save()
    print(f'[OK] Single label: {out}')


def generate_a4(data, out):
    """PDF หลายดวงบนหน้า A4 (Absolute Grid, no centering)"""
    items = data.get('items', [])
    if not items:
        print('[ERROR] No items', file=sys.stderr)
        sys.exit(1)

    lw = data.get('labelW', 50)
    lh = data.get('labelH', 25)
    tm = data.get('marginTop', 10)
    bm = data.get('marginBottom', 10)
    lm = data.get('marginLeft', 10)
    rm = data.get('marginRight', 10)
    gx = data.get('gapX', 3)
    gy = data.get('gapY', 3)
    theme = data.get('theme', 'light')
    font_size = data.get('fontSizeSerial', 12)

    A4_W, A4_H = 210.0, 297.0

    # คำนวณขอบเขตตารางป้าย
    cols = max(1, int((A4_W - lm - rm + gx) / (lw + gx)))
    rows = max(1, int((A4_H - tm - bm + gy) / (lh + gy)))
    per_page = cols * rows

    start_x = lm
    start_y = tm

    c = canvas.Canvas(out, pagesize=A4)
    page_h_mm = A4_H

    for idx, item in enumerate(items):
        if idx > 0 and idx % per_page == 0:
            c.showPage()   # ขึ้นหน้าแผ่นใหม่

        pos = idx % per_page
        col = pos % cols
        row = pos // cols

        abs_x = start_x + col * (lw + gx)
        abs_y = start_y + row * (lh + gy)

        render_label_absolute(
            c, abs_x, abs_y, lw, lh,
            item.get('serial', 'SN-UNKNOWN'),
            theme, font_size, page_h_mm
        )

    c.save()
    total_pages = (len(items) + per_page - 1) // per_page
    print(f'[OK] A4 PDF >> {len(items)} labels, {total_pages} pages >> {out}')


def main():
    p = argparse.ArgumentParser(description='Industrial Barcode Label Generator')
    p.add_argument('--mode', choices=['single', 'a4'], default='single')
    p.add_argument('--serial', default='SN-0000001')
    p.add_argument('--width', type=float, default=50)
    p.add_argument('--height', type=float, default=25)
    p.add_argument('--theme', default='light', choices=['light', 'dark'])
    p.add_argument('--font-size-serial', type=float, default=12)
    p.add_argument('--json-input', help='JSON file for A4 mode')
    p.add_argument('--output', required=True)
    args = p.parse_args()

    if args.mode == 'single':
        generate_single(
            serial=args.serial,
            w=args.width,
            h=args.height,
            theme=args.theme,
            font_size=args.font_size_serial,
            out=args.output
        )
    elif args.mode == 'a4':
        if not args.json_input:
            print('[ERROR] --json-input required for A4 mode', file=sys.stderr)
            sys.exit(1)
        with open(args.json_input, encoding='utf-8') as f:
            data = json.load(f)
        generate_a4(data, args.output)
    else:
        print('[ERROR] Unknown mode', file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()