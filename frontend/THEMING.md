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
| 2026-09-08 | — | Scanner: แก้ missing error state (ReferenceError on render), missing controlsRef/foundRef, camera preview ไม่ขึ้นตอนเปิดใหม่ → ใช้ native getUserMedia + video element + rAF decode loop; ลบ legacy public/scan.html, Settings "เปิด Scanner" → navigate /scan |
| 2026-09-09 | — | Trace: แปลงจาก legacy public/trace.html → หน้า React public /trace/:serial (ดูประวัติไม่ต้องล็อกอิน; ปุ่มโอนย้ายแสดงเมื่อล็อกอิน) + ลิงก์ Asset/Farm/Bundle ใช้ route ใหม่ + server redirect /trace.html → /trace/:serial |
| 2026-09-09 | — | QR Label: แปลงจาก legacy public/qr.html → หน้า React /qr (ตารางเลือกอุปกรณ์, Label/A4, presets, ดีไซน์, templates, JsBarcode preview, PDF/พิมพ์) + dep jsbarcode ใน npm + redirect /qr.html → /qr |
| 2026-09-09 | — | Mobile responsive: หน้าเลือกอุปกรณ์ /qr เปลี่ยนตารางเป็น Mobile Card List (ข้อมูลครบ ไม่ล้นจอ), Label preview/ตารางห่อ overflow ซ่อนบีบ; เพิ่มปุ่มลัด "ฉลาก" ใน TopBar ข้างปุ่มสแกน (เข้าหน้า /qr ได้ทันทีจากทุกหน้า ไม่ต้องผ่าน asset) |
| 2026-09-09 | — | TopBar: ปุ่มลัด สแกน/ฉลาก เปลี่ยนเป็น Segmented Control (กล่องเดียว พื้น neutral เส้นแบ่ง 2 ช่อง) แทนปุ่มทึบสีเขียว+น้ำเงิน, แก้ .print-sheet selector (เติมจุด) — A4 preview/print ตำแหน่งถูกต้อง |
| 2026-09-09 | — | Backup View: แปลง legacy public/backup-view.html → หน้า React /backup (Admin; เลือกตาราง/โหลด/Refresh/Export CSV) + redirect /backup-view.html; ลบ legacy index.html (dead) + audit.html (redirect /audit) + scripts/update-categories.js |
| 2026-09-09 | — | StatPill: component กลาง mini KPI (icon box + label + ตัวเลขใหญ่, tone blue/green/amber/red) ใช้แทน StatChip/StatCard แบบ flat ทั้ง Asset + Farm (Dashboard/Bundle ใช้ card ใหญ่แบบเดียวกันอยู่แล้ว) — ดูพรีเมียม เป็นชุดเดียวกันทั้งระบบ |
| 2026-09-09 | — | Dashboard: redesign พรีเมียม (ใช้ token เดิม) — hero gradient (totalItems + refresh), glass KPI 4 ใบ (Office/Site/เบิก/คืนวันนี้), KPI cards (ฟาร์ม/ติดตั้ง/สต็อก), สถานะอุปกรณ์แบบ progress bar (4 สถานะ), อันดับฟาร์มติดตั้งสูงสุด + แถบ rank; chart ใช้ segmented range control |
| 2026-09-09 | — | เลือกโรงเรือน (HouseID/HouseName) ครบทั้ง 2 flow ของการย้าย: (1) Bundle DeployModal เพิ่มช่องโรงเรือน cascade ตามฟาร์ม + เขียน Asset_List L/M ให้สมาชิกทุกชิ้น (recall ล้างค่า), (2) TransferModal แก้บั๊ก houseId ค้างข้ามฟาร์ม + dropdown แสดง `HouseID · ชื่อ (ประเภท)` + เตือนเมื่อฟาร์มยังไม่มีโรงเรือน; โรงเรือนไม่บังคับ (ถ้าไม่ระบุจะล้างค่าเดิม); แก้บั๊ก routes/asset.js อ่าน Asset_List!A2:Q ให้ครบคอลัมน์ N (wasInBundle เคยเป็น dead code) |
| 2026-09-09 | — | จัด data model ตำแหน่งให้ไม่ซ้ำซ้อน: ไซต์งาน/ฟาร์ม (SiteName H) = ตัวระบุฟาร์ม, **โรงเรือน (L/M) = จุดย่อยระดับอาคาร**, **ตำแหน่งย่อยในไซต์ (Location G) = จุดย่อยลึกลงไปอีก** — ตัด select "Location" ที่อ่านจาก sites ชุดเดียวกันแล้วแย่งค่ากับช่องไซต์ (ทำให้ G เป็น `-` เมื่อลืมเลือก ขณะที่ H เป็นชื่อฟาร์ม = ข้อมูลไม่สอดคล้อง); ช่องใหม่เป็น text + `<datalist>` จาก `GET /api/asset-locations` (ตำแหน่งย่อยที่เคยใช้, ตัด Stock/Intranin ออก) พร้อม disable เมื่ออยู่คลัง/ซ่อม; เพิ่มคอลัมน์ "โรงเรือน" ในตาราง Asset + mobile card + CSV export; หัวคอลัมน์เปลี่ยนเป็นไทย (ตำแหน่งย่อย / ไซต์-ฟาร์ม); ข้อมูลเก่าที่ G ยังเป็นชื่อฟาร์ม กรองไม่ให้แสดงซ้ำที่ public/asset.html + ScanPage |
| 2026-09-29 | — | **รวมตำแหน่งให้ตามรอบ (clarity pass จาก feedback ผู้ใช้)**: ตั้งชื่อชั้นที่ 3 เป็น **"จุดติดตั้ง"** (แทน "ตำแหน่งย่อย") และแสดงเป็น **เส้นทางเต็ม `ฟาร์ม › โรงเรือน › จุดติดตั้ง` + ป้ายย่อยกำกับชั้น** — สร้าง `utils/location.js` (buildLocation/buildBundleLocation, จัดการ Stock/Intranin/ค่าว่าง/ข้อมูลเก่าที่ G เคยเป็นชื่อฟาร์ม) + component `LocationPath` ใช้ร่วมกันทุกหน้า; **Bundle Detail** แสดงตำแหน่งเต็มจากสมาชิก (เดิมมีแต่ชื่อฟาร์มจาก Bundles!E) + เตือนเป็นสีเหลืองเมื่อสมาชิกอยู่คนละจุด (เกิดจากย้ายเฉพาะรายชิ้นหลังย้ายทั้งชุด) พร้อมปุ่ม "ดูตำแหน่งเต็ม" บนการ์ด; **หน้า Asset** เพิ่มคอลัมน์ "อยู่ในชุด" (ป้ายชื่อชุด) + ตำแหน่งเต็มแทน 3 คอลัมน์เดิม, CSV ใช้เส้นทางเต็ม; API `/api/assets` แนบ `bundleName` (join จาก sheet Bundles, cache + ล้างใน clearAssetCache/clearBundleCache) และ `/api/bundles/asset-info` เพิ่ม houseId/houseName; หน้า QR สาธารณะใช้ศัพท์เดียวกัน |
| 2026-10-01 | — | **UX simplify (ตาม feedback "แท็บเยอะ หลายสเต็ป ไม่มีสอน")**: (1) เมนูใหม่ `navigation.js` — 3 ปุ่มงานประจำวัน (หน้าแรก / เบิก–คืนของ / ย้าย–โอนอุปกรณ์) + กลุ่ม "เพิ่มเติม" พับเก็บไว้ (พับ/กางได้, คุมด้วย `UI_FLAGS.simpleMenu`) + ย้ายหน้าข้อมูลทั้งหมดเข้ากลุ่มเพิ่มเติม (route เดิมยังใช้ได้ทุกเส้น แค่ไม่โชว์ในเมนูหลัก), (2) หน้าแรกใหม่ `pages/Home.jsx` แบบ **Workdesk** — ค้นหาชื่อ/รหัส/Serial ได้ทั้งระบบ (asset + stock) แล้วกด action ต่อจากการ์ดผลลัพธ์ได้ทันที (ย้าย→/scan?serial=, ประวัติ→/trace, ฉลาก→/qr?serial=, เบิก→/stock?q=) + การ์ดงานเร็ว 3 ใบ + StatPill สรุปวันนี้ 4 ตัว (ไม่มีกราฟ) + ล่าสุด 5 รายการจาก /api/history, (3) หลักการ **"มีแต่ไม่แสดง"** — สวิตช์กลาง `src/uiConfig.js` (`charts:false` ซ่อนกราฟใน /dashboard, `simpleMenu`, `onboardingTour`) โค้ดเดิมไม่ลบแม้แต่บรรทัด, (4) สอนในแอป — `OnboardingTour.jsx` ทัวร์ 4 สไลด์แสดงครั้งแรกครั้งเดียว (localStorage `intranin-onboarded-v1`) + `pages/Help.jsx` คู่มือ /help 3 งานหลัก + ปุ่ม ❓ ที่แถบบนทุกหน้า + รายการ "วิธีใช้งาน" ใน sidebar, (5) มือถือ — bottom nav 3 ปุ่ม (หน้าแรก/เบิก–คืน/ย้าย–โอน) ไม่ต้องเปิด drawer, (6) Stock/Scan รับ deep-link `?q=` / `?serial=` จากหน้าแรก, (7) เทสต์เพิ่ม `tests/navigation.test.jsx` + `tests/onboarding.test.jsx` |
| 2026-10-01 | — | **Home Workdesk ปรับตาม feedback**: (1) การ์ด "ล่าสุด" เปลี่ยนเป็น **"การโอนย้ายล่าสุด"** — ดึงจาก **Asset_History** (การย้าย/โอนอุปกรณ์รายชิ้น) ผ่าน endpoint ใหม่ `GET /api/asset-history-recent` (backend, cache 60s + invalidate เมื่อ transfer/status เปลี่ยน; Transfer_Log เดิมมีแค่เบิก/คืนของคลัง), แถวโชว์ action + serial + ปลายทาง → + วันที่, ปุ่มด้านการ์ด → /asset, (2) หน้าแรกเพิ่ม **แถวสถิติฟาร์ม/อุปกรณ์รายชิ้น** (จาก /api/dashboard-full): ฟาร์มทั้งหมด · ติดตั้งที่ฟาร์ม · รอติดตั้ง (สต็อก) · รายชิ้นทั้งหมด — ให้ผู้ใช้ใหม่เห็นภาพรวมระบบทันทีที่เข้า |
| 2026-10-01 | — | **หน้า ย้าย/โอนอุปกรณ์ (/scan) ออกแบบใหม่ตาม feedback "user งง + กล้องเปิดเอง"**: (1) **กล้องไม่เปิดเองอีกต่อไป** — ตัด auto-start ตอน mount (เดิม `useEffect→startCam` ทำให้ popup ขอสิทธิ์กล้องเด้งทันทีที่เข้าหน้า) → เปลี่ยนเป็นการ์ด "เปิดกล้องเพื่อสแกนบาร์โค้ด" ปุ่มใหญ่ให้ผู้ใช้กดเอง + cleanup ปิดกล้องเมื่อออกจากหน้า + ปุ่มควบคุม (ปิดกล้อง/พร้อมสแกน) แสดงเฉพาะตอนกล้องเปิด, (2) เพิ่ม **ช่องค้นหาอุปกรณ์ที่จะย้าย** เป็นทางเข้าหลักอันดับ 1 — พิมพ์ชื่อ/รหัส/Serial/assetId (โหลดทะเบียน `/api/assets` ครั้งเดียว กรอง client-side) แสดงผลพร้อม StatusBadge + ตำแหน่งเต็ม (LocationPath) + ปุ่ม **โอนย้าย** (เปิด TransferModal ตรงนั้น) / ประวัติ → ผู้ใช้ไม่ต้องจำ Serial ก็ย้ายได้, (3) ลำดับหน้าใหม่: ค้นหา → กล้อง(ปิด) → พิมพ์ Serial(USB Scanner), การ์ด "สแกนสำเร็จ" เปลี่ยนชื่อเป็น "พบอุปกรณ์" (เพราะมาจากการค้นหาได้ด้วย); deep-link `?serial=` จากหน้าแรกยังทำงาน |
| 2026-10-01 | — | **เพิ่มอุปกรณ์ใหม่ (Asset) โฟลว์เดียว — ทุก user ใช้ได้**: component ใหม่ `components/AddDeviceModal.jsx` — ① เลือก Part แบบ **typeahead** (พิมพ์ชื่อ/รหัส → suggestion จาก Part_Catalog พร้อมจำนวนที่มี; ไม่พบ → กด "สร้าง Part ใหม่" แล้วระบบสร้างให้เองผ่าน /api/add-part) ② จำนวนชิ้นเท่าไหร่ก็ได้ ③ **preview Serial ก่อนเพิ่ม** (SN-PART-DDMMYYYY-XXXX) + ส่วนขั้นสูงพับเก็บ (สถานะ/ไซต์/จุดติดตั้ง/ผู้รับผิดชอบ มี default ครบ) → submit เรียก /api/bulk-add-asset (สร้าง Serial+AssetID+Asset_History อัตโนมัติ) แล้วโชว์ **Serial จริงที่สร้าง**; **backend กัน Serial ซ้ำเด็ดขาด (all-or-nothing)** — `bulk-add-asset` ตรวจ Serial ชุดใหม่ทุกตัวกับทั้ง Asset_List + กันซ้ำในชุดเอง ถ้าซ้ำแม้แต่ตัวเดียว → 400 และไม่เขียนอะไรลงชีต; จุดวางปุ่ม: หน้าแรก (ใต้แถบค้นหา) · หน้าสแกน (ค้นไม่พบ → "เพิ่มอุปกรณ์ใหม่ {คำค้น}" prefill + เสร็จแล้ว resolve Serial แรกให้เลย) · แทนปุ่ม "เพิ่ม Asset"+"เพิ่มหลายชิ้น" ในหน้า Asset (ลด 2 modals เดิมเหลือ 1 โฟลว์); ส่วน Stock (add-item) ยังไม่แตะตามขอบเขตที่ตกลง |
| 2026-10-01 | — | **Home ปรับตาม feedback รอบ 2**: (1) ลบข้อความทักทาย ("สวัสดี... portal_user 👋 / วันนี้ต้องทำอะไรครับ") ออกจากหน้าแรก, (2) **ปุ่มลอย (FAB)** แทน segmented control บนแถบบน — สแกน (วงกลมเขียวใหญ่ + ring ขาว) และ ฉลาก (วงกลมน้ำเงิน) อยู่มุมขวาล่างแบบปุ่มติดต่อบนเว็บทั่วไป + label โชว์ตอน hover; ปุ่ม ❓ วิธีใช้งาน เอาออกจากแถบบน (ยังอยู่ใน sidebar), (3) หน้าเพิ่มอุปกรณ์ — เพิ่มปุ่ม **"เพิ่มอุปกรณ์อื่นต่อ"** บนหน้าสำเร็จ (เคลียร์ฟอร์มให้เพิ่มต่อได้ทันทีโดยไม่ต้องปิด-เปิดใหม่), (4) **สรุปภาพรวม redesign แบบพรีเมียม** (รักษาธีม token เดิม): hero gradient ink + glass KPI 4 ใบ (ในคลัง/Site/เบิกวันนี้/คืนวันนี้) + ปุ่มรีเฟรช แยกการ์ดขาว "อุปกรณ์และฟาร์ม" เป็นแถวสถิติแบบสะอาด (ฟาร์ม/ติดตั้งที่ฟาร์ม/รอติดตั้ง/รายชิ้นทั้งหมด) — เลย์เอาต์ 2 คอลัมน์บนจอใหญ่ (HeroTile/MiniStat helpers ใหม่) |
| 2026-10-06 | — | **เพิ่มฟาร์ม/โรงเรือนได้ทุกคน + โอนย้ายง่ายขึ้น + พิมพ์ฉลากง่ายขึ้น (แก้ 3 จุดที่ user ติด)**: **(A)** `POST /api/add-farm-site` + `/api/add-farm-house` เปิดให้ผู้ใช้ที่ล็อกอิน**ทุกคน**เพิ่มได้ (เดิม admin เท่านั้น → user โดน 403 จนติดขั้นตอนโอนย้าย) + **กันเพิ่มซ้ำ** siteId/siteName (case-insensitive) และ houseId ในฟาร์มเดิม → **409** ข้อความไทย + บันทึก **Audit_Log** "เพิ่มฟาร์ม/เพิ่มโรงเรือน" (เดิมไม่มี audit เลย) + helpers ใหม่ `services/farmGuard.js` (pure) พร้อมเทส `tests/farm-permission.test.js` (ยิง route chain จริงผ่าน fake sheets: user เพิ่มได้ / 401 / admin ผ่าน / ซ้ำ 409); หน้า Farm เอา gate admin ออกจากปุ่ม + ฟอร์มเหลือฟิลด์หลัก (ชื่อฟาร์ม+ประเภท / ฟาร์ม+ชื่อโรงเรือน) รหัส auto-suggest จากชื่อ/ไล่เลข `SITE-H01` (`utils/farmId.js` + เทส) ส่วนจังหวัด/ผู้จัดการ/ประเภทโรงเรือน/ความจุ/หมายเหตุ พับเข้า "ตัวเลือกเพิ่มเติม" — โค้ดเดิมไม่ลบ; component ใหม่ **`InlineFarmAdd.jsx`** (ฟอร์มย่อเพิ่มฟาร์ม/โรงเรือน + error 409 แสดงในแผง) ฝังข้าง dropdown ใน **TransferModal** และ **Bundle DeployModal** — เพิ่มแล้ว dropdown refresh + **เลือกค่าใหม่ให้เลย** ไม่ต้องออกจาก modal; ข้อความ dead-end "ฟาร์มนี้ยังไม่มีโรงเรือน — ติดต่อผู้ดูแล" เปลี่ยนเป็นปุ่ม "+ เพิ่มโรงเรือนทันที" (กดได้ทั้ง 2 modal); **(B)** TransferModal รีดีไซน์ **"เลือกปลายทางก่อน"** — ปุ่มใหญ่ 3 ปุ่ม (🚜 ไปฟาร์ม / 🏠 คืนคลังกลาง / 🔧 ส่งซ่อม) แทน dropdown "ประเภทการโอนย้าย", สถานะ + ประเภทฟาร์ม **auto** ตามปลายทาง/ฟาร์มที่เลือก, แถบสรุป "จาก X → ฟาร์ม › โรงเรือน › จุดติดตั้ง" ก่อนกดยืนยัน, ส่วน ประเภทการโอนย้าย 4 แบบ/สถานะ/ประเภทสัตว์/หมายเหตุ ย้ายเข้า "ตัวเลือกเพิ่มเติม" (พับได้) — **payload `/api/transfer-asset` ไม่เปลี่ยนทุกบิต** (ใช้ร่วม Asset/Bundle/Farm/Scan/Trace เดิม) + เทส `frontend/src/tests/transferModal.test.jsx`; **(C)** QrPage โครงใหม่ 3 ขั้น: เลือกอุปกรณ์ → **ขนาด 3 ปุ่มใหญ่** (เล็ก 38×20 / กลาง 50×25 / ใหญ่ 60×30 + คำนวณดวง/แผ่น A4 อัตโนมัติ) → **ปุ่มพิมพ์ใหญ่ + ดาวน์โหลด PDF**; ตัวเลือกละเอียดทั้งหมด (โหมด/ขนาด mm/margin/gap/สีพื้น/badge/เทมเพลต) ย้ายเข้า **"ตั้งค่าขั้นสูง"** พับเก็บ default + **จำค่าที่ใช้ล่าสุดอัตโนมัติ** (localStorage `ems_label_last_v1`) — ไม่แตะ backend label / generate_label.py |