import { useState } from 'react';
import Icon from '../components/ui/Icon.jsx';

export default function Report() {
  const [title, setTitle] = useState('');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [reportType, setReportType] = useState('all');
  const [vehicle, setVehicle] = useState('');
  const [locations, setLocations] = useState('');
  const [employeeCount, setEmployeeCount] = useState('');
  const [employees, setEmployees] = useState('');
  const [err, setErr] = useState('');

  function submit() {
    setErr('');
    if (!startDate || !endDate) { setErr('กรุณาเลือกช่วงวันที่'); return; }
    const data = {
      title: title || 'รายงานประวัติการเบิก–คืนอุปกรณ์',
      locations,
      vehicle: vehicle || '-',
      startDate,
      endDate,
      employeeCount: employeeCount || '0',
      employees,
      reportType,
    };
    const form = document.createElement('form');
    form.method = 'POST';
    form.action = '/api/export-history';
    form.target = '_blank';
    Object.keys(data).forEach((k) => {
      const i = document.createElement('input');
      i.type = 'hidden';
      i.name = k;
      i.value = data[k];
      form.appendChild(i);
    });
    document.body.appendChild(form);
    form.submit();
    document.body.removeChild(form);
  }

  return (
    <div className="max-w-3xl space-y-4">
      <div>
        <div className="text-lg font-bold text-[var(--text)]">รายงาน</div>
        <div className="text-[12px] text-[var(--tmuted)]">Export ข้อมูลเป็น PDF</div>
      </div>

      <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
        <div className="p-5 border-b border-[var(--g100)]">
          <div className="flex items-center gap-2 text-[15px] font-bold text-[var(--text)]">
            <span className="text-[var(--blue)]"><Icon name="description" /></span> Export รายงาน PDF
          </div>
          <div className="text-[13px] text-[var(--tmuted)]">กรอกข้อมูลให้ครบแล้วกด Export</div>
        </div>

        <div className="p-5 grid grid-cols-1 sm:grid-cols-2 gap-4">
          <F label="หัวข้อรายงาน" span>
            <input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="รายงานประจำเดือน..." className={inp} />
          </F>
          <F label="วันที่เริ่ม *">
            <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} className={inp} />
          </F>
          <F label="วันที่สิ้นสุด *">
            <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} className={inp} />
          </F>
          <F label="ประเภทรายการ">
            <select value={reportType} onChange={(e) => setReportType(e.target.value)} className={inp}>
              <option value="all">เบิกและคืน</option>
              <option value="borrow">เฉพาะเบิก</option>
              <option value="return">เฉพาะคืน</option>
            </select>
          </F>
          <F label="ยานพาหนะ">
            <input value={vehicle} onChange={(e) => setVehicle(e.target.value)} placeholder="เช่น รถยนต์" className={inp} />
          </F>
          <F label="จำนวนพนักงาน">
            <input type="number" value={employeeCount} onChange={(e) => setEmployeeCount(e.target.value)} placeholder="5" className={inp} />
          </F>
          <F label="รายชื่อสถานที่ (แต่ละบรรทัด)" span>
            <textarea value={locations} onChange={(e) => setLocations(e.target.value)} rows={3} placeholder={'โรงงาน A\nโกดัง B'} className={inp + ' resize-y'} />
          </F>
          <F label="รายชื่อพนักงาน (แต่ละบรรทัด)" span>
            <textarea value={employees} onChange={(e) => setEmployees(e.target.value)} rows={3} placeholder={'สมชาย\nสมหญิง'} className={inp + ' resize-y'} />
          </F>
        </div>

        {err && <div className="flex items-center gap-1.5 px-5 pb-4 text-[13px] text-[var(--red)]"><Icon name="warning" size="sm" /> {err}</div>}

        <div className="p-5 pt-0">
          <button onClick={submit} className="w-full flex items-center justify-center gap-2 py-3 rounded-xl bg-[var(--blue)] text-white text-[14px] font-semibold hover:bg-[var(--blue-d)]">
            <Icon name="description" size="sm" /> Export PDF
          </button>
        </div>
      </div>

      <div className="flex items-center gap-1.5 px-4 py-3 rounded-xl bg-[var(--blue-l)] border border-[var(--blue-b)] text-[12px] text-[var(--blue)]">
        <Icon name="lightbulb" size="sm" /> PDF จะเปิดในแท็บใหม่ — สามารถใช้ปุ่มเลือกประเภทรายการเพื่อกรองเฉพาะรายการเบิก หรือรายการคืน
      </div>
    </div>
  );
}

function F({ label, children, span }) {
  return (
    <div className={span ? 'sm:col-span-2 space-y-1' : 'space-y-1'}>
      <label className="block text-[12px] font-medium text-[var(--tsub)]">{label}</label>
      {children}
    </div>
  );
}

const inp = 'w-full h-9 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px] focus:outline-none focus:border-[var(--blue)]';