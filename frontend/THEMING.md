# 🎨 THEMING — คู่มือ Frontend Design / UI-UX (ฉบับทีมดูแลต่อ)

> **ไฟล์นี้คือจุดรวมกติกาเดียวสำหรับทุกคนที่แตะ UI/UX**  
> แก้ฟีเจอร์อะไร เพิ่มอะไร ที่เกี่ยวกับหน้าตา → **อ่านไฟล์นี้ก่อน แล้วจดอัปเดต log ท้ายไฟล์ให้เป็นปัจจุบันเสมอ**

---

## 1. ภาพรวมระบบ

- Stack: **React 19 + Vite + Tailwind v4** (+ Chart.js ผ่าน react-chartjs-2)
- output: static files ไปที่ `frontend/dist/` → backend (Express) เสิร์ฟเป็น SPA ที่ root `/`
- ระบบงานจริง 3 ภาษาในตัว: ไทยเป็นหลัก, อังกฤษ (label บางที่)
- มี 2 โหมด: **portal mode** (เข้าแบบ regular user อัตโนมัติ) กับ **admin** (login /login)

---

## 2. รันโปรเจกต์ (ต้องทำก่อนแก้)

```bash
cd frontend
npm install      # ครั้งแรก
npm run dev      # dev server http://localhost:5173 (proxy /api ไป backend)
```

> ⚠️ ต้องรัน backend ด้วยก่อน (`node server.js` ในโฟลเดอร์ราก) ไม่งั้นหน้าเว็บโหลด แต่ข้อมูล API 404
>
> จุดสำคัญ: ระหว่าง dev **ไม่ต้อง build ทุกครั้ง** — Vite hot-reload ทันทีเมื่อเซฟ

---

## 3. แผนที่ไฟล์ — อยากแก้อะไร ไปเปิดไฟล์ไหน

| ต้องการแก้ | ไฟล์ |
|---|---|
| **สี/ธีมทั้งระบบ** (ตัวแปร --xxx) | `src/theme.css` ⭐ |
| base style, font, helper class | `src/index.css` |
| sidebar / topbar / layout / hamburger-mobile | `src/components/layout/Layout.jsx` |
| ไอคอน | `src/components/ui/Icon.jsx` (ใช้ `<Icon name="..." />`) |
| ป้ายสถานะสี | `src/components/ui/StatusBadge.jsx` |
| เลื่อนหน้าแบบแบ่งหน้า | `src/components/ui/Pagination.jsx` |
| หน้าแต่ละหน้า | `src/pages/` (Dashboard, Stock, Asset, Farm, History, Bundle, Report, Settings) |
| หน้า admin | `src/pages/admin/` |
| การ login/role | `src/context/AuthContext.jsx`, `src/pages/Login.jsx` |
| หมวด + ไอคอนเริ่มต้น | `src/data/categories.js` |
| จุดรวม router | `src/App.jsx` |
| **เมนู sidebar + ชื่อหน้า (topbar)** | `src/navigation.js` ⭐ (เพิ่มหน้าใหม่ต้องลงทะเบียนที่นี่) |
| config build | `vite.config.js` |

---

## 4. Design Tokens — แก้ธีมทั้งระบบที่ไฟล์เดียว (`src/theme.css`)

Component ทั้งหมดอ้างสี/ขนาดผ่าน `var(--xxx)` เท่านั้น → **หา theme ทั้งระบบแค่ไฟล์เดียว**

### สีที่เจอบ่อยสุดในงาน

| ตัวแปร | ความหมาย | แก้ที่บรรทัด |
|---|---|---|
| `--blue` / `--blue-d` / `--blue-l` / `--blue-b` | **สีหลักแบรนด์** (ปุ่ม, link, active nav) | 27–30 |
| `--emerald*` | สำเร็จ / เพิ่ม / ของใหม่ | 34–37 |
| `--amber*` | เตือน / กำลังดำเนินการ | 38–41 |
| `--red` / `--red-l` / `--red-b` | ลบ / อันตราย / logout / สถานะเสีย | 42–44 |
| `--purple` / `--purple-l` | สีรอง (badge พิเศษ) | 45–46 |
| `--bg` / `--surface` / `--surface2` | พื้นหลังโฟลเดอร์ / การ์ด | 49–51 |
| `--text` / `--tsub` / `--tmuted` | ข้อความหลัก / รอง / จาง | 54–56 |
| `--g100`…`--g700` | เส้นขอบ slate | 59–65 |
| `--r` / `--r2` / `--r3` / `--r4` | มุมโค้ง | 74–77 |
| `--sb-w` / `--topbar-h` | ง sidebar / สูง topbar | 80–82 |
| `--icon-*` | scale/weight ของไอคอน | 85–117 |

**เมื่อทีม portal มีสีแบรนด์บริษัท → เปลี่ยนค่าแค่ในไฟล์นี้ แล้ว build ใหม่เป็นอันเสร็จ**

---

## 5. กติกาเขียนโค้ด (กระบวนการใครก็ตาม)

1. **ห้าม hardcode สี Hex** — ใช้ `var(--xxx)` เสมอ
   - ถูก: `text-[#0A1628]` ✗ → ถูก: `text-[var(--ink)]` ✓
   - Tailwind ใช้ arbitrary value: `bg-[var(--blue)]`, `border-[var(--g200)]`
2. **ใช้ตัวแปร token ที่มีก่อน; เพิ่มตัวแปรใหม่ใน theme.css ไม่ใช่ฝังค่าใน component**
3. ไอคอนใช้ `<Icon name="material_symbols_name" size="md" weight="regular" />`
   - รายชื่อไอคอน: https://fonts.google.com/icons (Material Symbols Rounded, self-hosted)
4. กล่อง dialog/form: ปุ่มหลัก `bg-[var(--blue)]`, ปุ่มรอง transparent + border
5. ป้ายสถานะใช้ `StatusBadge` (สีตามสถานะในชีต) อย่าเขียนสีเอง
6. **Responsive + Mobile-first** (ดูหัวข้อถัดไป) ห้ามทำตารางเลื่อนแนวนอนบนมือถือ
7. ภาษาไทย/อังกฤษที่มีอยู่ ใช้ต่อตามแบบเดิม (ไทยหลัก, label อังกฤษเฉพาะคำที่เป็นเทคนิค)

---

## 6. Responsive ตามสไตล์ที่มีอยู่ (สำคัญ!)

- **≥ lg (1024px)**: sidebar 240px คงที่ (`lg:ml-[var(--sb-w)]`)
- **< lg**: sidebar กลายเป็น **drawer** (เลื่อนเข้า-ออก + overlay + hamburger)
- **จอใหญ่**: ตารางเต็ม; **จอเล็ก (มือถือ)**: ตารางถูกซ่อน (`hidden md:block`) แล้วเปลี่ยนเป็น **card list** (`md:hidden`) — ตัวอย่างใน `Stock.jsx`, `Asset.jsx`, `Farm.jsx`, `History.jsx`
- จอแคบให้ซ่อนอะไรที่ไม่จำเป็นได้ (`hidden sm:block` สำหรับ time/label)
- เบรกพอยต์: `sm=640, md=768, lg=1024`

---

## 7. Build + Deploy (ตอนงานเสร็จแล้ว)

```bash
cd frontend
npm run build     # ต้องผ่าน ไม่มี error/warning
```

แล้วจากโฟลเดอร์รากของโปรเจกต์:

```bash
git add -A
git commit -m "feat/fix: ..."
git push origin main
```

- ขึ้น Render: **Manual Deploy → Deploy latest commit** (Build command ตั้งไว้แล้ว: `npm run build`)
- เปิดเว็บแล้ว **Ctrl+F5** (กัน stale cache)
- bundle ใหญ่ได้ถึง 700kB ว่าปลาย warnings (`chunkSizeWarningLimit` ใน vite.config.js)

---

## 8. การทดสอบที่ควรเช็คก่อนส่งทุกครั้ง

- [ ] หน้า login/admin งานปกติ (portal user เข้าอัตโนมัติ, admin แยก role)
- [ ] มือถือ: hamburger เปิด sidebar, ตารางกลายเป็น card ไม่ล้นจอ
- [ ] ไม่มี console error (F12)
- [ ] `npm run build` ผ่านสะอาด
- [ ] อัปเดต log ด้านล่างเสร็จแล้ว

---

## 📝 ประวัติปรับปรุง (ทีมทุกคนช่วยอัปเดตด้วย)

| วันที่ | เวอร์ชัน | สิ่งที่แก้ |
|---|---|---|
| 2026-09-08 | — | สร้าง THEMING.md ฉบับแรก (คู่มือทีม design) |
| 2026-09-08 | — | Stock: เพิ่มระบบตะกร้าเบิกหลายรายการ (ปุ่ม + เบิก → cart → ยืนยัน; ของไม่พอ เบิกเท่าที่มี + แจ้ง) |
| 2026-09-08 | — | Stock: เบิกเป็น stepper −/+/จำนวน ตรงตัวแถวแบบ Shopper; ตัดระบบเบิก/คืนแถวเดิม (qty+select+✓) ทิ้ง |
| 2026-09-08 | — | Stock: คืนจาก Site เลือกได้ว่ากี่ชิ้นต่อรายการ (เดิมคืนเต็มจำนวนตลอด) + สรุปยอดรวม |
| 2026-09-08 | — | เพิ่มปุ่ม "สแกน" ใน topbar (ทุกหน้า): กล้องอ่าน Barcode/QR + พิมพ์ Serial ด้วยมือ; แสดงตำแหน่งปัจจุบัน + ประวัติ + โอนย้าย/คืนต่อทันที (dep: @zxing/browser) |
| 2026-09-08 | — | Scanner เปลี่ยนจาก modal เป็นหน้า /scan เต็มจอ + ปุ่มเปิด/ปิดกล้อง; แก้ reopen ไม่ preview (เลิกใช้ releaseAllStreams) |
| 2026-09-08 | — | Topbar แสดงชื่อหน้าตาม path จริง (ศูนย์กลางที่ navigation.js; เดิมค้าง "Dashboard" ทุกหน้า) |
| 2026-09-08 | — | ตัด title/subtitle ซ้ำใน header ของทุกหน้า (topbar เป็นคนถือชื่อหน้าแทน; หน้าเหลือแค่แถวปุ่ม action) |
| 2026-09-08 | — | นาฬิกา topbar real-time (เดิมค้าง 09:00) |
| 2026-09-08 | — | Silent chunk warning (limit 700 kB) |
| 2026-09-08 | — | Fixed Icon import หายใน Dashboard (error Icon is not defined) |
| 2026-09-08 | — | Dashboard: Line chart → rounded Bar chart |
| 2026-09-08 | — | Mobile card views: Stock/Asset/Farm/History + drawer sidebar |
| 2026-09-09 | — | Trace: แปลงจาก legacy public/trace.html → หน้า React public /trace/:serial (ดูประวัติไม่ต้องล็อกอิน; ปุ่มโอนย้ายแสดงเมื่อล็อกอิน) + ลิงก์ Asset/Farm/Bundle ใช้ route ใหม่ + server redirect /trace.html → /trace/:serial |
| 2026-09-09 | — | QR Label: แปลงจาก legacy public/qr.html → หน้า React /qr (ตารางเลือกอุปกรณ์, Label/A4, presets, ดีไซน์, templates, JsBarcode preview, PDF/พิมพ์) + dep jsbarcode ใน npm + redirect /qr.html → /qr |
| 2026-09-09 | — | Mobile responsive: หน้าเลือกอุปกรณ์ /qr เปลี่ยนตารางเป็น Mobile Card List (ข้อมูลครบ ไม่ล้นจอ), Label preview/ตารางห่อ overflow ซ่อนบีบ; เพิ่มปุ่มลัด "ฉลาก" ใน TopBar ข้างปุ่มสแกน (เข้าหน้า /qr ได้ทันทีจากทุกหน้า ไม่ต้องผ่าน asset) |