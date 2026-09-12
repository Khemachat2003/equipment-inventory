// scripts/repair_misplaced_stock_rows.js
// ซ่อม Stock_Master แถวที่ถูก append ผิดคอลัมน์ (bug ของ values.append ช่วง "A:I" —
// เริ่มเขียนแถวใหม่ที่คอลัมน์ I แทน A) เช่น แถว T089 ที่ code ไปอยู่คอลัมน์ I..Q
//
// วิธีรัน:
//   node scripts/repair_misplaced_stock_rows.js            # เขียนจริง
//   node scripts/repair_misplaced_stock_rows.js --dry-run  # ดูแผนเฉย ๆ ยังไม่เขียน
//
// สิ่งที่ทำ:
//   1. ค้นหาแถวใน Stock_Master ที่คอลัมน์ A ว่าง แต่มีข้อมูลหลุดไปอยู่ที่ I..Q
//   2. ย้าย (relocate) กลับมา A..I ของแถวเดิม + ล้าง I..Q ที่ทิ้งไว้
//   3. ล้าง "ข้อความหลง" ใน Stock_Office / Stock_Site (คอลัมน์ D..Z) เฉพาะแถวที่ A..C ว่าง
//      และคอลัมน์ D..Z ของแถวที่เพิ่งย้าย code กลับ (คอลัมน์เกินหัวตาราง A..C = stray)

require('dotenv').config();
const { getSheetsClient, SPREADSHEET_ID } = require('../services/sheets');

const DRY_RUN = process.argv.includes('--dry-run');
const IMG_BASE = 'https://cdn.jsdelivr.net/gh/Khemachat2003/stock-image/images';
const COLS = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i));

async function main() {
  if (DRY_RUN) console.log('🔎 DRY-RUN — จะยังไม่เขียนลงสเปรดชีต\n');
  const sheets = await getSheetsClient();

  const masterRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'Stock_Master!A1:Z600',
  });
  const masterRows = masterRes.data.values || [];

  const updates = [];
  const clears = [];
  const relocated = [];

  for (let i = 0; i < masterRows.length; i++) {
    const row = masterRows[i];
    const code = row[0];
    const shiftedCode = row[8]; // I (ถ้า A ว่าง แต่ I มีค่า → แถวเขียนเพี้ยน)
    if (code) continue; // แถวปกติ
    if (!shiftedCode) continue; // ไม่มีอะไรหลุด

    const rowNo = i + 1;
    const name = row[9] || ''; // J
    const total = parseInt(row[13]) || 0; // N
    const ext = row[16] || ''; // Q
    const imageFormula = ext ? `=IMAGE("${IMG_BASE}/${shiftedCode}.${ext}")` : '';

    console.log(`🛠  relocate แถว ${rowNo}: code "${shiftedCode}" หลุดที่ I..Q → ย้ายกลับ A..I`);
    console.log(`   A=${JSON.stringify(shiftedCode)}  B=${JSON.stringify(name)}  F=${total}  G=${total}  I=${JSON.stringify(ext)}`);

    updates.push({
      range: `Stock_Master!A${rowNo}:I${rowNo}`,
      values: [[shiftedCode, name, imageFormula, '', '', total, total, '', ext]],
    });
    // ล้าง A..Z ของแถวเดิมที่เหลือ (กันข้อมูลเพี้ยนตกค้าง เช่น Q="png" ถูกต้องแล้วใน col I ใหม่)
    clears.push({ range: `Stock_Master!J${rowNo}:Z${rowNo}`, values: [[...Array(17).fill('')]] });
    relocated.push({ code: shiftedCode, row: rowNo });
  }

  // ล้างข้อความหลงใน Stock_Office / Stock_Site (คอลัมน์ D..Z ที่ไม่มีหัวตาราง)
  for (const sheetName of ['Stock_Office', 'Stock_Site']) {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A1:Z400`,
    });
    const rows = res.data.values || [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i] || [];
      const aCode = row[0];
      const strayCells = [];
      for (let c = 3; c < 26; c++) { // D..Z
        if (row[c] !== '' && row[c] != null) strayCells.push(`${COLS[c]}${i + 1}=${JSON.stringify(row[c])}`);
      }
      if (!strayCells.length) continue;
      // เฉพาะถ้าเป็นแถวนอกตาราง (A..C ว่าง) หรือแถวของ code ที่เพิ่งย้าย → ล้างได้
      const isOrphan = !aCode && !row[1] && !row[2];
      const isRelocated = relocated.some((r) => r.code === aCode);
      if (isOrphan || isRelocated) {
        console.log(`🧹 clear stray ${sheetName}: แถว ${i + 1} -> ${strayCells.join(', ')}`);
        clears.push({ range: `${sheetName}!D${i + 1}:Z${i + 1}`, values: [[...Array(23).fill('')]] });
      } else {
        console.log(`⚠️  พบ cell นอกตาราง ${sheetName} แถว ${i + 1}: ${strayCells.join(', ')} — ยังไม่แตะ (A/B/C มีข้อมูล กรุณาตรวจเอง)`);
      }
    }
  }

  const all = [...updates, ...clears];
  console.log(`\n✅ แผนการซ่อม: ย้าย ${updates.length} แถว, ล้าง ${clears.length} ช่วง`);

  if (all.length === 0) {
    console.log('ℹ️  ไม่พบแถวที่พิมพ์เพี้ยน — ไม่มีการเปลี่ยนแปลง');
    return;
  }

  if (DRY_RUN) {
    console.log('\n🔎 (dry-run) ยังไม่ได้เขียนอะไรลงสเปรดชีต — พิมพ์ --dry-run ออกเพื่อใช้งานจริง');
    return;
  }

  for (let i = 0; i < all.length; i += 100) {
    const batch = all.slice(i, i + 100);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'USER_ENTERED', data: batch },
    });
  }
  console.log('✅ เขียนลงสเปรดชีตเรียบร้อย');
}

main().catch((err) => {
  console.error('❌ Repair error:', err);
  process.exit(1);
});