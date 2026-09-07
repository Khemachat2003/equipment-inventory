require('dotenv').config({ path: 'D:/Stock Equipment/equipment-inventory/.env' });
const { getSheetsClient, SPREADSHEET_ID } = require('D:/Stock Equipment/equipment-inventory/services/sheets');

const BY_NAME = {
  'Sensors & Weather': ['เซ็นเซอร์วัดฝุ่น', 'เซ็นเซอร์แอมโมเนีย NH3', 'เซ็นเซอร์ซัลเฟอร์ไดออกไซด์ SO2', 'เซ็นเซอร์ก๊าซไฮโดรเจนซัลไฟด์', 'Weather Station Flip Ultrasonic', 'MI-Volt-Protection'],
  'Controllers & Microcontrollers': ['radxa dragon q6a', 'Pi Control Flip', 'Fan control_PIC30F', 'บอร์ด Hydroponics'],
  'Networking & Comms': ['RouterSim', 'Microtik Switch Hub 5 Port', 'Powerline Module'],
  'Power Supplies': ['Power Supply Mean Well 12V', 'Power Supply Mean Well 5V', 'Power Supply Delta 24V'],
  'Power Protection & Switching': ['เบรกเกอร์ AC2P220V63A', 'เบรกเกอร์ NF125', 'แมกเนติก S-T100', 'Passthrough'],
  'Display & Peripherals': ['โมดูลจอ LCD 20x4 + I2C', 'Radxa Camera M8 219'],
  'IO & Automation Modules': ['Modbus RTU Relay 8-CH'],
};
const NAME_TO_CAT = {};
Object.entries(BY_NAME).forEach(([cat, names]) => names.forEach((n) => { NAME_TO_CAT[n] = cat; }));

(async () => {
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Part_Catalog!A2:B' });
  const rows = res.data.values || [];
  const updates = [];
  let filled = 0;
  rows.forEach((r, i) => {
    if (r[0] === undefined) return;
    const name = (r[1] || '').trim();
    const cat = NAME_TO_CAT[name] || '';
    if (cat) filled++;
    updates.push({ range: `Part_Catalog!C${i + 2}`, values: [[cat]] });
  });
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { valueInputOption: 'USER_ENTERED', data: updates },
  });
  console.log(`✔ อัปเดต Part_Catalog คอลัมน์ C แล้ว ${filled} แถว (จาก ${rows.length})`);
})().catch((e) => { console.error('❌', e.message); process.exit(1); });