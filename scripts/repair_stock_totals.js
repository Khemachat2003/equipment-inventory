// scripts/repair_stock_totals.js
// ซ่อม Stock_Master คอลัมน์ F (ทั้งหมด) ให้ = Office + Site ทุกรายการ
// หลักการ:
//  - รายการที่มีแถวในทั้ง Stock_Office และ Stock_Site → เขียน F = office + site
//  - รายการที่ไม่มีแถวใน Office/Site → ข้าม + รายงานชื่อไว้ (กันลบข้อมูลเดิมทิ้ง)
//
// วิธีรัน:  node scripts/repair_stock_totals.js
// หมายเหตุ: เขียนข้อมูลจริงลง Google Sheets — ควรตรวจ output แบบ dry-run ตาม post-run

require('dotenv').config();
const { getSheetsClient, SPREADSHEET_ID } = require('../services/sheets');

const DRY_RUN = process.argv.includes('--dry-run');

async function main() {
  if (DRY_RUN) console.log('🔎 DRY-RUN — จะยังไม่เขียนลงสเปรดชีต\n');
  const sheets = await getSheetsClient();

  const masterRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'Stock_Master!A2:I',
  });
  const officeRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'Stock_Office!A2:C',
  });
  const siteRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'Stock_Site!A2:C',
  });

  const masterData = masterRes.data.values || [];
  const officeData = officeRes.data.values || [];
  const siteData = siteRes.data.values || [];

  const officeByCode = new Map(officeData.map((r) => [r[0], parseInt(r[2] || 0) || 0]));
  const siteByCode = new Map(siteData.map((r) => [r[0], parseInt(r[2] || 0) || 0]));

  const updates = [];
  const skipped = [];
  let fixed = 0;

  for (let i = 0; i < masterData.length; i++) {
    const row = masterData[i];
    const code = row[0];
    if (!code) continue;
    const officeQty = officeByCode.has(code) ? officeByCode.get(code) : null;
    const siteQty = siteByCode.has(code) ? siteByCode.get(code) : null;

    // ถ้าไม่มีแถวใน Office/Site เลย → ข้าม (ป้องกัน data loss) แล้วรายงาน
    if (officeQty == null && siteQty == null) {
      skipped.push(`${code} (${row[1] || ''})`);
      continue;
    }

    const expected = (officeQty || 0) + (siteQty || 0);
    const current = parseInt(row[5] || 0);
    if (current !== expected) {
      updates.push({
        range: `Stock_Master!F${i + 2}`,
        values: [[expected]],
      });
      fixed++;
      console.log(`🛠  ${code}\tF ${current} → ${expected} (Office ${officeQty || 0} + Site ${siteQty || 0})`);
    }
  }

  if (updates.length > 0) {
    if (DRY_RUN) {
      console.log(`\n🔎 (dry-run) จะเขียน ${updates.length} เซลล์ — ยังไม่ได้เขียน`);
    } else {
      // เขียนทีละช่วงสั้นๆ (batchUpdate รับ payload <5MB) กัน timeout
      for (let i = 0; i < updates.length; i += 200) {
        const batch = updates.slice(i, i + 200);
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: SPREADSHEET_ID,
          requestBody: { valueInputOption: 'RAW', data: batch },
        });
      }
    }
  }

  console.log(`\n✅ แก้ไขแล้ว ${fixed} รายการ`);
  if (skipped.length) {
    console.log(`⚠️  ข้าม ${skipped.length} รายการที่ไม่มีแถวใน Office/Site (ต้องจัดการมือ):`);
    skipped.forEach((s) => console.log('   -', s));
  }
  if (fixed === 0) console.log('ℹ️  ไม่มีรายการที่ต้องซ่อม — ทุกตัวตรงอยู่แล้ว');
}

main().catch((err) => {
  console.error('❌ Repair error:', err);
  process.exit(1);
});