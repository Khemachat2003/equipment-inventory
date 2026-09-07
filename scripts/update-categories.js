require('dotenv').config({ path: 'D:/Stock Equipment/equipment-inventory/.env' });
const { getSheetsClient, SPREADSHEET_ID } = require('D:/Stock Equipment/equipment-inventory/services/sheets');

const MAPPING = {
  // Sensors & Weather
  'เซ็นเซอร์วัดฝุ่น': 'Sensors & Weather',
  'เซ็นเซอร์แอมโมเนีย NH3': 'Sensors & Weather',
  'เซ็นเซอร์ซัลเฟอร์ไดออกไซด์ SO2': 'Sensors & Weather',
  'เซ็นเซอร์ก๊าซไฮโดรเจนซัลไฟด์': 'Sensors & Weather',
  'Weather Station Flip Ultrasonic': 'Sensors & Weather',
  'MI-Volt-Protection': 'Sensors & Weather',
  // Controllers & Microcontrollers
  'radxa dragon q6a': 'Controllers & Microcontrollers',
  'Pi Control Flip': 'Controllers & Microcontrollers',
  'Fan control_PIC30F': 'Controllers & Microcontrollers',
  'บอร์ด Hydroponics': 'Controllers & Microcontrollers',
  // Networking & Comms
  'RouterSim': 'Networking & Comms',
  'Microtik Switch Hub 5 Port': 'Networking & Comms',
  'Powerline Module': 'Networking & Comms',
  // Power Supplies
  'Power Supply Mean Well 12V': 'Power Supplies',
  'Power Supply Mean Well 5V': 'Power Supplies',
  'Power Supply Delta 24V': 'Power Supplies',
  // Power Protection & Switching
  'เบรกเกอร์ AC2P220V63A': 'Power Protection & Switching',
  'เบรกเกอร์ NF125': 'Power Protection & Switching',
  'แมกเนติก S-T100': 'Power Protection & Switching',
  'Passthrough': 'Power Protection & Switching',
  // Display & Peripherals
  'โมดูลจอ LCD 20x4 + I2C': 'Display & Peripherals',
  'Radxa Camera M8 219': 'Display & Peripherals',
  // IO & Automation Modules
  'Modbus RTU Relay 8-CH': 'IO & Automation Modules',
};

const CATEGORIES = [
  ['Sensors & Weather', 'เซนเซอร์ & สภาพอากาศ', 'sensors'],
  ['Controllers & Microcontrollers', 'คอนโทรลเลอร์ & ไมโครคอนโทรลเลอร์', 'memory'],
  ['Networking & Comms', 'เครือข่าย & การสื่อสาร', 'router'],
  ['Power Supplies', 'แหล่งจ่ายไฟ', 'power'],
  ['Power Protection & Switching', 'อุปกรณ์ป้องกัน & สวิตชิ่ง', 'shield'],
  ['Display & Peripherals', 'จอแสดงผล & อุปกรณ์ต่อพ่วง', 'monitor'],
  ['IO & Automation Modules', 'โมดูล IO & ระบบอัตโนมัติ', 'settings_input_component'],
];

(async () => {
  const sheets = await getSheetsClient();
  const ss = spreadsheetId = SPREADSHEET_ID;

  // 1) header
  await sheets.spreadsheets.values.update({
    spreadsheetId: ss, range: 'Asset_List!R1', valueInputOption: 'USER_ENTERED',
    requestBody: { values: [['Category']] },
  });
  console.log('✔ header R1 = Category');

  // 2) fill R column
  const assetRes = await sheets.spreadsheets.values.get({ spreadsheetId: ss, range: 'Asset_List!A2:C' });
  const rows = assetRes.data.values || [];
  const updates = [];
  let filled = 0;
  rows.forEach((r, i) => {
    if (r[0] === undefined) return;
    const name = (r[2] || '').trim();
    const cat = MAPPING[name] || '';
    if (cat) filled++;
    updates.push({ range: `Asset_List!R${i + 2}`, values: [[cat]] });
  });
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: ss,
    requestBody: { valueInputOption: 'USER_ENTERED', data: updates },
  });
  console.log(`✔ เติม Category แล้ว ${filled} แถว (จาก ${rows.length} แถว)`);

  // 3) ensure Categories sheet exists
  const meta = await sheets.spreadsheets.get({ spreadsheetId: ss });
  const titles = meta.data.sheets.map((s) => s.properties.title);
  if (!titles.includes('Categories')) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: ss,
      requestBody: { requests: [{ addSheet: { properties: { title: 'Categories' } } }] },
    });
    console.log('✔ สร้าง sheet Categories');
  }
  await sheets.spreadsheets.values.update({
    spreadsheetId: ss, range: 'Categories!A1:C', valueInputOption: 'USER_ENTERED',
    requestBody: { values: [['Name', 'Label', 'Icon'], ...CATEGORIES] },
  });
  console.log('✔ เขียน Categories (7 หมวด)');

  // verify
  const chk = await sheets.spreadsheets.values.get({ spreadsheetId: ss, range: 'Asset_List!R1:R' });
  const cs = (chk.data.values || []).slice(1).filter((r) => r[0] && r[0].trim()).length;
  console.log(`✔ ตรวจสอบ: R เติมแล้ว ${cs} แถว`);
})().catch((e) => { console.error('❌', e.message); process.exit(1); });