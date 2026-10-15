// AddDeviceModal — เพิ่มอุปกรณ์ใหม่ (ทะเบียนรายชิ้น + Serial) โฟลว์เดียว ใช้ได้ทุก user
// การเลือก Part: พิมพ์แล้วมี suggestion — เจอ Part เดิมกดเลือก, ไม่เจอกด "สร้าง Part ใหม่"
// Serial สร้างโดย backend (SN-PART-DDMMYYYY-XXXX) + backend ตรวจซ้ำแบบ all-or-nothing
// (preview ในฟอร์มเป็นเพียง "คำเดาก่อนส่ง" — ตัวจริงคืนจาก backend ตอน success)
//
// Episode 2 — Fast-Track Multi-Step Guided Flow: หลังรับเข้าสำเร็จ (Step 1) เดินต่อใน Modal เดียว
//   Step 2 จัดชุดติดตั้ง (เลือกชุดเดิม / สร้างชุดใหม่ทันทีผ่าน CreateBundleModal)
//   Step 3 โอนย้ายไปฟาร์ม/โรงเรือน/คลัง (เพิ่มฟาร์ม/โรงเรือนใหม่ได้ทันทีใน Modal ด้วย InlineFarmAdd)
//   Step 4 พิมพ์ฉลาก QR (Direct-to-Print จาก Episode 1) — ทุกขั้นถัดจาก Step 1 ไม่บังคับ ข้ามได้

import { useEffect, useMemo, useState } from 'react';
import axios from 'axios';
import Icon from './ui/Icon.jsx';
import { useBusy, BusyOverlay } from './ui/Busy.jsx';
import { showToast, ToastHost } from './ui/Toast.jsx';
import CreateBundleModal from './CreateBundleModal.jsx';
import { FarmInlineAdd, HouseInlineAdd } from './InlineFarmAdd.jsx';

const ain = 'w-full h-9 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px] focus:outline-none focus:border-[var(--blue)]';

function toLocalDateTimeValue(date) {
  const local = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 16);
}

function formatReceivedAt(value) {
  if (!value) return 'ไม่ระบุวัน';
  return new Intl.DateTimeFormat('th-TH', { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(value));
}

function sanitizePartNumber(v) {
  return (v || '').trim().toUpperCase().replace(/\s+/g, '-');
}

// ── Episode 2: แถบขั้นตอนของ Fast-Track Flow (Step 1 รับเข้า → 2 จัดชุด → 3 โอนย้าย → 4 พิมพ์ฉลาก) ──
const FLOW_STEPS = [
  { n: 1, label: 'รับเข้าอุปกรณ์' },
  { n: 2, label: 'จัดชุดติดตั้ง' },
  { n: 3, label: 'โอนย้ายปลายทาง' },
  { n: 4, label: 'พิมพ์ฉลาก QR' },
];

function FlowStepper({ current }) {
  return (
    <div className="flex items-center gap-1.5 flex-wrap">
      {FLOW_STEPS.map((s) => {
        const done = s.n < current;
        const active = s.n === current;
        return (
          <span
            key={s.n}
            className={`inline-flex items-center gap-1 px-2 py-1 rounded-full text-[11px] font-semibold border ${active ? 'bg-[var(--blue)] text-white border-[var(--blue)]' : done ? 'bg-[var(--emerald-l)] text-[var(--emerald-d)] border-[var(--emerald-b)]' : 'bg-[var(--surface2)] text-[var(--tmuted)] border-[var(--g200)]'}`}
          >
            {done ? <Icon name="check" size="xs" /> : <span className="w-3 text-center">{s.n}</span>}
            {s.label}
          </span>
        );
      })}
    </div>
  );
}

// หัวแผงของแต่ละ Step — บอกชัดว่า "ไม่บังคับ ข้ามได้" + ปุ่มย้อนกลับ (สำหรับขั้นที่เลยมาแล้ว)
function StepPanel({ title, onBack, children }) {
  return (
    <div className="rounded-xl border border-[var(--g200)] bg-[var(--surface2)] p-3 space-y-2.5">
      <div className="flex items-center justify-between gap-2">
        <div className="text-[12px] font-bold text-[var(--text)]">{title} <span className="font-medium text-[var(--tmuted)]">(ไม่บังคับ — ข้ามได้)</span></div>
        {onBack && <button onClick={onBack} className="text-[11px] font-semibold text-[var(--blue)] shrink-0">← ย้อนกลับ</button>}
      </div>
      {children}
    </div>
  );
}

export default function AddDeviceModal({ open, presetName = '', presetAdvanced = false, onClose, onDone }) {
  const [query, setQuery] = useState('');
  const [selected, setSelected] = useState(null); // { partNumber, partName, isNew }
  const [qty, setQty] = useState('1');
  const [advanced, setAdvanced] = useState(false);
  const [status, setStatus] = useState('ใช้งานได้');
  const [siteName, setSiteName] = useState('Intranin');
  const [location, setLocation] = useState('Stock');
  const [user, setUser] = useState('');
  const [receivedAt, setReceivedAt] = useState(() => toLocalDateTimeValue(new Date()));
  const [poNumber, setPoNumber] = useState('');
  const [supplier, setSupplier] = useState('');
  const [batchMode, setBatchMode] = useState('new');
  const [batchId, setBatchId] = useState('');
  const [batchQuery, setBatchQuery] = useState('');
  const [availableBatches, setAvailableBatches] = useState([]);
  const [batchesLoading, setBatchesLoading] = useState(false);
  const [batchListLoaded, setBatchListLoaded] = useState(false);
  const [sites, setSites] = useState([]);
  const [parts, setParts] = useState([]);
  const [result, setResult] = useState(null); // { serials, added, partCreated }
  const [newPart, setNewPart] = useState(null); // โหมดกรอก Part ใหม่ { partNumber, partName }
  const busy = useBusy();

  // ── Episode 2: state ของ Fast-Track Multi-Step Flow ──────────────────────────
  // flowStep: 0 = ยังอยู่ฟอร์มรับเข้า (Step 1) | 2/3/4 = ขั้นตอนเสริมหลังรับเข้าสำเร็จ (ข้ามได้ทุกขั้น)
  const [flowStep, setFlowStep] = useState(0);
  const [bundleOptions, setBundleOptions] = useState([]); // ชุดติดตั้งที่มีอยู่ (โหลดเมื่อเข้า Step 2)
  const [bundlesLoaded, setBundlesLoaded] = useState(false);
  const [bundlesLoading, setBundlesLoading] = useState(false);
  const [bundleQuery, setBundleQuery] = useState('');
  const [pickedBundleId, setPickedBundleId] = useState('');
  const [createBundleOpen, setCreateBundleOpen] = useState(false); // สร้างชุดใหม่ทันที (nested modal)
  // Step 3 — ปลายทาง
  const [destMode, setDestMode] = useState('stock'); // 'stock' = คงเก็บคลังกลาง | 'farm' = ฟาร์ม/โรงเรือน
  const [farmId, setFarmId] = useState('');
  const [houseId, setHouseId] = useState('');
  const [houses, setHouses] = useState([]);
  const [loadingHouses, setLoadingHouses] = useState(false);
  const [installPoint, setInstallPoint] = useState('');
  const [extraFarms, setExtraFarms] = useState([]); // ฟาร์มที่เพิ่งสร้างจากฟอร์มย่อ — เติม dropdown ทันที
  const [farmAddOpen, setFarmAddOpen] = useState(false);
  const [houseAddOpen, setHouseAddOpen] = useState(false);
  const [stepNote, setStepNote] = useState(null); // ผลลัพธ์ของขั้นตอนล่าสุด { tone: 'ok'|'err', text }
  const [transferFailed, setTransferFailed] = useState([]); // Serial ที่ย้ายไม่สำเร็จ (Step 3)

  useEffect(() => {
    if (!open) return;
    setQuery(presetName || '');
    setSelected(null);
    setQty('1');
    setStatus('ใช้งานได้');
    setSiteName('Intranin');
    setLocation('Stock');
    setUser('');
    setReceivedAt(toLocalDateTimeValue(new Date()));
    setPoNumber('');
    setSupplier('');
    setBatchMode('new');
    setBatchId('');
    setBatchQuery('');
    setBatchListLoaded(false);
    // "รับเข้า" เปิดมาพร้อมแผนล็อต/PO ขยายไว้ล่วงหน้า (presetAdvanced)
    setAdvanced(Boolean(presetAdvanced));
    setResult(null);
    setNewPart(null);
    // Episode 2 — เคลียร์ Fast-Track Flow ทุกครั้งที่เปิดใหม่
    setFlowStep(0);
    setBundleOptions([]);
    setBundlesLoaded(false);
    setBundlesLoading(false);
    setBundleQuery('');
    setPickedBundleId('');
    setCreateBundleOpen(false);
    setDestMode('stock');
    setFarmId('');
    setHouseId('');
    setHouses([]);
    setLoadingHouses(false);
    setInstallPoint('');
    setExtraFarms([]);
    setFarmAddOpen(false);
    setHouseAddOpen(false);
    setStepNote(null);
    setTransferFailed([]);
    axios.get('/api/farm-sites').then(({ data }) => setSites(data || [])).catch(() => {});
    axios.get('/api/part-catalog').then(({ data }) => setParts(data || [])).catch(() => {});
    setAvailableBatches([]);
    setBatchListLoaded(false);
  }, [open, presetName, presetAdvanced]);

  const selectedBatch = availableBatches.find((batch) => batch.batchId === batchId);
  const filteredBatches = useMemo(() => {
    const term = batchQuery.trim().toLowerCase();
    if (!term) return availableBatches;
    return availableBatches.filter((batch) => [batch.batchId, batch.poNumber, batch.supplier].some((value) => String(value || '').toLowerCase().includes(term)));
  }, [availableBatches, batchQuery]);

  async function chooseExistingBatchMode() {
    setBatchMode('existing');
    if (batchListLoaded || batchesLoading) return;
    setBatchesLoading(true);
    try {
      const { data } = await axios.get('/api/inbound-batches');
      setAvailableBatches(data || []);
      setBatchListLoaded(true);
    } catch {
      showToast('โหลดรายการล็อตไม่สำเร็จ กรุณาลองอีกครั้ง', { type: 'err' });
    } finally {
      setBatchesLoading(false);
    }
  }

  // ═══════════════════════════════════════════════════════════════════
  // Episode 2 — Fast-Track Multi-Step Guided Flow (Step 2 → 3 → 4)
  // ทุกขั้น "ไม่บังคับ" — กดข้ามได้ตามความเร่งด่วนหน้างาน และย้อนกลับมาแก้ได้
  // ═══════════════════════════════════════════════════════════════════
  const flowSerials = result?.serials || [];
  const flowAssetIds = result?.assetIds || (Array.isArray(result?.assets) ? result.assets.map((a) => a.assetId) : []);

  const filteredBundleOptions = useMemo(() => {
    const term = bundleQuery.trim().toLowerCase();
    if (!term) return bundleOptions;
    return bundleOptions.filter((b) => [b.bundleId, b.bundleName, b.description].some((v) => String(v || '').toLowerCase().includes(term)));
  }, [bundleOptions, bundleQuery]);
  const pickedBundle = bundleOptions.find((b) => b.bundleId === pickedBundleId);

  // ฟาร์มทั้งหมดสำหรับ Step 3 = ที่โหลดไว้ตั้งแต่เปิด Modal + ที่เพิ่งสร้างจากฟอร์มย่อ (ไม่ซ้ำ)
  const flowFarms = [...(sites || []), ...extraFarms].filter(
    (s, i, arr) => s?.siteId && s.siteId !== 'Intranin' && s.siteName && arr.findIndex((x) => x.siteId === s.siteId) === i
  );
  const selectedFarm = flowFarms.find((s) => s.siteId === farmId);
  const flowHouseName = houses.find((h) => h.houseId === houseId)?.houseName || '';

  // เข้า Step 2 → โหลดชุดติดตั้งที่มีอยู่ให้เลือก (event-driven: เรียกจากจุดที่ย้ายเข้า Step 2 เท่านั้น)
  function enterStep2() {
    setFlowStep(2);
    if (bundlesLoaded || bundlesLoading) return;
    setBundlesLoading(true);
    axios.get('/api/bundles')
      .then(({ data }) => { setBundleOptions(data || []); setBundlesLoaded(true); })
      .catch(() => { setBundlesLoaded(true); })
      .finally(() => setBundlesLoading(false));
  }

  // Step 3 — โหลดโรงเรือนของฟาร์มที่เลือก (เปลี่ยนฟาร์ม = ล้างโรงเรือนเดิมเสมอ เหมือน DeployModal)
  function refreshHouses(fid) {
    if (!fid) { setHouses([]); return; }
    setLoadingHouses(true);
    axios.get(`/api/farm-houses/${encodeURIComponent(fid)}`)
      .then(({ data }) => setHouses(data || []))
      .catch(() => setHouses([]))
      .finally(() => setLoadingHouses(false));
  }

  function chooseFarm(siteId) {
    setFarmId(siteId);
    setHouseId('');
    setHouseAddOpen(false);
    refreshHouses(siteId);
  }

  // Step 2 — เพิ่มอุปกรณ์ที่เพิ่งรับเข้าเข้า "ชุดที่เลือก" (bulk: ทั้งหมดในคำสั่งเดียว)
  async function confirmAddToBundle() {
    if (!pickedBundleId) return;
    if (!flowAssetIds.length) return showToast('ไม่พบรหัสอุปกรณ์ (Asset ID) ของชิ้นที่เพิ่งรับเข้า — ไม่สามารถเพิ่มเข้าชุดได้', { type: 'warn' });
    await busy.run('กำลังเพิ่มเข้าชุด...', async () => {
      try {
        const { data } = await axios.post(`/api/bundles/${encodeURIComponent(pickedBundleId)}/assets/bulk`, { assetIds: flowAssetIds });
        if (!data.success) throw new Error(data.error || 'เพิ่มเข้าชุดไม่สำเร็จ');
        setStepNote({ tone: 'ok', text: `เพิ่ม ${flowAssetIds.length} ชิ้นเข้าชุด ${pickedBundle?.bundleName || ''} (${pickedBundleId}) แล้ว` });
        setFlowStep(3);
      } catch (e) {
        showToast('เพิ่มเข้าชุดไม่สำเร็จ: ' + (e.response?.data?.error || e.message), { type: 'err' });
      }
    });
  }

  // Step 2 — สร้างชุดใหม่ทันทีจากฟอร์มย่อ (CreateBundleModal รันเลขรหัสให้) แล้วเพิ่มอุปกรณ์เข้าชุดนั้นเลย
  async function handleCreateBundleNow(form) {
    await busy.run('กำลังสร้างชุดใหม่...', async () => {
      try {
        const { data } = await axios.post('/api/bundles', form);
        if (data && data.error) throw new Error(data.error);
        let addedNote = `สร้างชุด ${form.bundleId} — ${form.bundleName} แล้ว (ยังไม่ได้เพิ่มอุปกรณ์เข้าชุด)`;
        if (flowAssetIds.length) {
          const r2 = await axios.post(`/api/bundles/${encodeURIComponent(form.bundleId)}/assets/bulk`, { assetIds: flowAssetIds });
          if (!r2.data.success) throw new Error(r2.data.error || 'เพิ่มอุปกรณ์เข้าชุดไม่สำเร็จ');
          addedNote = `สร้างชุด ${form.bundleId} — ${form.bundleName} และเพิ่ม ${flowAssetIds.length} ชิ้นเข้าชุดแล้ว`;
        }
        setBundleOptions((prev) => [...prev, { bundleId: form.bundleId, bundleName: form.bundleName, description: form.description || '', status: form.status || 'In Stock', location: 'Stock', assetIds: [...flowAssetIds] }]);
        setPickedBundleId(form.bundleId);
        setCreateBundleOpen(false);
        setStepNote({ tone: 'ok', text: addedNote });
        setFlowStep(3);
      } catch (e) {
        showToast('สร้างชุดไม่สำเร็จ: ' + (e.response?.data?.error || e.message), { type: 'err' });
      }
    });
  }

  // Step 3 — เพิ่มฟาร์ม/โรงเรือนใหม่จากฟอร์มย่อ → เติม dropdown + เลือกให้เลย (ไม่ต้องออกจาก Modal)
  function handleFlowFarmAdded(site) {
    const f = { siteId: site.siteId, siteName: site.siteName, farmType: site.farmType || '' };
    setExtraFarms((prev) => (prev.some((x) => x.siteId === f.siteId) ? prev : [...prev, f]));
    chooseFarm(f.siteId);
    setFarmAddOpen(false);
  }

  function handleFlowHouseAdded(house) {
    if (house?.houseId) setHouseId(house.houseId);
    setHouseAddOpen(false);
    refreshHouses(farmId);
  }

  // Step 3 — ยืนยันโอนย้ายรายชิ้นที่เพิ่งรับเข้า ไปยังปลายทางที่เลือก (payload เดียวกับ TransferModal)
  async function confirmFlowTransfer() {
    if (destMode === 'stock') {
      setStepNote({ tone: 'ok', text: 'คงเก็บที่คลังกลาง (Intranin · Stock) ตามที่รับเข้า' });
      setFlowStep(4);
      return;
    }
    if (!farmId) return showToast('เลือกฟาร์ม/ไซต์ปลายทางก่อน', { type: 'warn' });
    await busy.run('กำลังโอนย้าย...', async () => {
      const fromLocation = pickedBundle
        ? `${pickedBundle.bundleName || pickedBundleId} (${pickedBundleId}) — ${pickedBundle.location || 'Stock'}`
        : `${siteName} (${location.trim() || '-'})`;
      const succeeded = [];
      const failed = [];
      for (const s of flowSerials) {
        try {
          const { data } = await axios.post('/api/transfer-asset', {
            serialNumber: s,
            action: 'ย้ายตำแหน่งอุปกรณ์',
            status,
            location: installPoint.trim() || '-',
            siteName: selectedFarm?.siteName || '',
            remark: '',
            farmType: selectedFarm?.farmType || '-',
            animalType: '-',
            houseId: houseId || '-',
            houseName: houseId ? (flowHouseName || '-') : '-',
            user: user.trim(),
            fromLocation,
          });
          if (!data.success) throw new Error(data.error || 'โอนย้ายไม่สำเร็จ');
          succeeded.push(s);
        } catch (e) {
          failed.push({ serialNumber: s, error: e.response?.data?.error || e.message || 'โอนย้ายไม่สำเร็จ' });
        }
      }
      if (failed.length) {
        setTransferFailed(failed.map((f) => f.serialNumber));
        setStepNote({ tone: 'err', text: `ย้ายสำเร็จ ${succeeded.length} ชิ้น · ไม่สำเร็จ ${failed.length} ชิ้น (${failed.map((f) => f.serialNumber).join(', ')}) — กดยืนยันโอนย้ายอีกครั้งเพื่อลองใหม่` });
        return;
      }
      setTransferFailed([]);
      setStepNote({ tone: 'ok', text: `ย้าย ${flowSerials.length} ชิ้นไป ${selectedFarm?.siteName || ''}${houseId ? ` › ${flowHouseName || houseId}` : ''} แล้ว` });
      setFlowStep(4);
    });
  }

  const pn = selected
    ? selected.partNumber
    : newPart
      ? sanitizePartNumber(newPart.partNumber)
      : sanitizePartNumber(query);
  const qtyNum = Math.max(1, parseInt(qty) || 0);

  const suggestions = useMemo(() => {
    if (selected) return null;
    const kk = query.trim().toLowerCase();
    if (!kk) return null;
    const hits = parts
      .filter((p) =>
        (p.partNumber || '').toLowerCase().includes(kk) ||
        (p.partName || '').toLowerCase().includes(kk)
      )
      .slice(0, 6);
    const exact = parts.some((p) => sanitizePartNumber(p.partNumber) === pn && pn);
    return { hits, showCreate: pn.length >= 2 && !exact };
  }, [query, selected, parts, pn]);

  // Serial preview — คำเดาก่อนส่ง (ตัวจริง backend สร้าง + ตรวจซ้ำ all-or-nothing)
  const serialPreview = useMemo(() => {
    if (!pn || !qtyNum) return [];
    const now = new Date();
    const ds = String(now.getDate()).padStart(2, '0') + String(now.getMonth() + 1).padStart(2, '0') + now.getFullYear();
    const prefix = `SN-${pn}-${ds}`;
    return Array.from({ length: Math.min(qtyNum, 4) }, (_, i) => `${prefix}-${String(i + 1).padStart(4, '0')}`);
  }, [pn, qtyNum]);

  async function confirmNewPart() {
    const pnum = sanitizePartNumber(newPart.partNumber);
    const pname = (newPart.partName || '').trim();
    if (!pnum) return showToast('กรอก Part Number ก่อน (ตัวย่อภาษาอังกฤษ เช่น SENWT)', { type: 'warn' });
    if (!/^[A-Z0-9][A-Z0-9._-]*$/.test(pnum)) {
      return showToast('Part Number ใช้ได้เฉพาะ A-Z 0-9 . _ - เท่านั้น (ห้ามช่องว่าง/ภาษาไทย) — เพราะจะถูกใช้สร้าง Serial', { type: 'warn' });
    }
    if (parts.some((p) => sanitizePartNumber(p.partNumber) === pnum)) {
      return showToast('Part Number นี้มีอยู่แล้วในระบบ — กดเลือกจากรายการด้านบนแทน', { type: 'warn' });
    }
    setSelected({ partNumber: pnum, partName: pname || pnum, isNew: true });
    setNewPart(null);
  }

  async function submit() {
    if (batchMode === 'existing' && !selectedBatch) return showToast('กรุณาเลือกล็อตเดิมก่อน', { type: 'warn' });
    if (!selected) return showToast('เลือกอุปกรณ์จากรายการ หรือกด "สร้าง Part ใหม่" ก่อน', { type: 'warn' });
    if (!qtyNum || qtyNum < 1) return showToast('ระบุจำนวนชิ้นให้ถูกต้อง', { type: 'warn' });
    const partNumber = selected.partNumber;
    const partName = selected.partName || partNumber;
    await busy.run('กำลังเพิ่มอุปกรณ์...', async () => {
      try {
        // ① Part ใหม่ → สร้างใน Part_Catalog ก่อน (bulk-add จะ +total ให้เองเมื่อ Part มีอยู่)
        if (selected.isNew) {
          const r1 = await axios.post('/api/add-part', { partNumber, partName, unit: 'ชิ้น' });
          if (r1.data && r1.data.error) throw new Error('สร้าง Part ไม่สำเร็จ: ' + r1.data.error);
        }
        // ② สร้าง Serial + ทะเบียนรายชิ้น (backend ตรวจซ้ำ all-or-nothing — ซ้ำตัวเดียวไม่เขียนเลย)
        const r2 = await axios.post('/api/bulk-add-asset', {
          partNumber,
          partName,
          qty: qtyNum,
          status,
          siteName,
          location: location.trim() || '-',
          user: user.trim(),
          receivedAt: new Date(receivedAt).toISOString(),
          poNumber: poNumber.trim(),
          supplier: supplier.trim(),
          ...(batchMode === 'existing' ? { batchId } : {}),
        });
        if (!r2.data.success) throw new Error(r2.data.error || 'เพิ่ม Asset ไม่สำเร็จ');
        setResult({ serials: r2.data.serials || [], assetIds: r2.data.assetIds || (Array.isArray(r2.data.assets) ? r2.data.assets.map((a) => a.assetId) : []), added: r2.data.added || qtyNum, partCreated: !!selected.isNew, batchId: r2.data.batchId, receivedAt: r2.data.receivedAt, poNumber: r2.data.poNumber, supplier: r2.data.supplier, inboundStorage: r2.data.inboundStorage });
        // Episode 2 — รับเข้าสำเร็จ → เข้าสู่ Fast-Track Flow ที่ Step 2 (จัดชุด — ข้ามได้)
        enterStep2();
      } catch (e) {
        showToast('เกิดข้อผิดพลาด: ' + (e.response?.data?.error || e.message), { type: 'err' });
      }
    });
  }

  function startAnother() {
    // เพิ่มอุปกรณ์ถัดไปทันที — เคลียร์ฟอร์มให้ใหม่ (คง modal เปิดอยู่)
    setResult(null);
    setSelected(null);
    setQuery('');
    setQty('1');
    setStatus('ใช้งานได้');
    setSiteName('Intranin');
    setLocation('Stock');
    setUser('');
    setReceivedAt(toLocalDateTimeValue(new Date()));
    setPoNumber('');
    setSupplier('');
    setBatchMode('new');
    setBatchId('');
    setBatchQuery('');
    setBatchListLoaded(false);
    setAdvanced(false);
    // Episode 2 — กลับไปเริ่มที่ฟอร์ม (Step 1) และเคลียร์ Fast-Track Flow ให้สะอาด
    setFlowStep(0);
    setBundleOptions([]);
    setBundlesLoaded(false);
    setBundlesLoading(false);
    setBundleQuery('');
    setPickedBundleId('');
    setCreateBundleOpen(false);
    setDestMode('stock');
    setFarmId('');
    setHouseId('');
    setHouses([]);
    setLoadingHouses(false);
    setInstallPoint('');
    setFarmAddOpen(false);
    setHouseAddOpen(false);
    setStepNote(null);
    setTransferFailed([]);
  }

  // Episode 1 — Direct-to-Print: เปิดหน้าพิมพ์ฉลาก QR (/qr) ของอุปกรณ์ที่เพิ่งเพิ่มทันที
  // ส่ง Serial ทั้งชุดผ่าน deep-link ?serials=... → หน้า /qr เลือกไว้ให้พร้อมพิมพ์เลย
  // (เปิดแท็บใหม่ตาม pattern ปุ่มฉลากทุกจุดในระบบ — หน้าเดิมทำงานต่อได้ทันที)
  function openLabelPrint() {
    const serials = (result?.serials || []).filter(Boolean);
    if (!serials.length) return;
    window.open(`/qr?serials=${encodeURIComponent(serials.join(','))}`, '_blank', 'noopener');
  }

  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center p-4 bg-black/40 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
      <div className="w-full max-w-lg bg-[var(--surface)] rounded-2xl shadow-xl my-8" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-[var(--g100)]">
          <div className="flex items-center gap-2">
            <span className="flex items-center justify-center w-8 h-8 rounded-lg bg-[var(--blue-l)] text-[var(--blue)]"><Icon name="add" size="sm" /></span>
            <span className="text-[15px] font-bold text-[var(--text)]">เพิ่มอุปกรณ์ใหม่</span>
          </div>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)]"><Icon name="close" size="sm" /></button>
        </div>

        {result ? (
          /* ── สำเร็จ — โชว์ Serial จริงที่ backend สร้าง (ตรวจซ้ำแล้ว) ── */
          <div className="p-5">
            <div className="flex flex-col items-center gap-2 py-3">
              <span className="w-12 h-12 flex items-center justify-center rounded-full bg-[var(--emerald-l)] text-[var(--emerald-d)]"><Icon name="check_circle" size="2xl" /></span>
              <div className="text-[15px] font-bold text-[var(--text)]">เพิ่ม {result.added} ชิ้นสำเร็จ</div>
              {result.partCreated && <div className="text-[12px] text-[var(--tmuted)]">สร้าง Part "{selected.partNumber}" ให้อัตโนมัติแล้ว</div>}
            </div>
            {result.batchId && <div className="text-center text-[13px] font-semibold text-[var(--blue)]">ล็อต {result.batchId} · รับเข้า {formatReceivedAt(result.receivedAt)}</div>}
            {(result.poNumber || result.supplier) && <div className="text-center text-[12px] text-[var(--tsub)]">{[result.poNumber && `PO ${result.poNumber}`, result.supplier].filter(Boolean).join(' · ')}</div>}
            {result.inboundStorage === 'local-file' && <div role="status" className="rounded-lg bg-amber-50 px-3 py-2 text-[12px] text-amber-800">บันทึกข้อมูลล็อตไว้ใน Localhost แล้ว (Google Sheets ยังไม่พร้อม)</div>}
            <div className="px-3 py-2.5 rounded-lg bg-[var(--blue-l)] border border-[var(--blue-b)]">
              <div className="text-[11px] font-bold text-[var(--blue)] mb-1">Serial ที่สร้าง (ตรวจซ้ำโดยระบบแล้ว):</div>
              <div className="max-h-40 overflow-y-auto space-y-0.5">
                {result.serials.map((s) => <div key={s} className="font-mono text-[12px] text-[var(--blue)]">{s}</div>)}
              </div>
            </div>
            {/* Episode 1 — Direct-to-Print: ปุ่มพิมพ์ฉลากทันทีหลังบันทึกสำเร็จ (ไม่ต้องไปหน้าพิมพ์ฉลากแล้วไล่หา Serial ใหม่) */}
            {result.serials.length > 0 && (
              <button
                onClick={openLabelPrint}
                title="เปิดหน้าพิมพ์ฉลาก QR พร้อมเลือก Serial ที่เพิ่งสร้างไว้ให้เลย"
                className={`w-full flex items-center justify-center gap-1.5 px-4 py-2.5 rounded-lg bg-[var(--blue)] text-white text-[13px] font-bold hover:bg-[var(--blue-d)] mt-4 ${flowStep === 4 ? 'ring-2 ring-[var(--blue)] ring-offset-2 ring-offset-[var(--surface)]' : ''}`}
              >
                <Icon name="print" size="sm" /> พิมพ์ฉลาก QR ทันที ({result.serials.length} ดวง)
              </button>
            )}

            {/* ── Episode 2: Fast-Track Multi-Step Flow — จัดชุด → โอนย้าย → พิมพ์ฉลาก (ทุกขั้น ข้ามได้) ── */}
            <div className="mt-4 pt-3 border-t border-[var(--g100)] space-y-2.5">
              <FlowStepper current={flowStep} />
              {stepNote && (
                <div role="status" className={`px-3 py-2 rounded-lg text-[12px] font-medium border ${stepNote.tone === 'err' ? 'bg-[var(--red-l)] border-[var(--red-b)] text-[var(--red)]' : 'bg-[var(--emerald-l)] border-[var(--emerald-b)] text-[var(--emerald-d)]'}`}>
                  {stepNote.text}
                </div>
              )}

              {/* Step 2 — จัดชุดติดตั้ง: เลือกชุดที่มีอยู่ / สร้างชุดใหม่ทันที / ข้าม */}
              {flowStep === 2 && (
                <StepPanel title="② จัดชุดติดตั้ง">
                  <div className="space-y-2">
                    <input value={bundleQuery} onChange={(e) => setBundleQuery(e.target.value)} placeholder="ค้นหาชุดติดตั้ง (รหัส/ชื่อ)" className={ain} aria-label="ค้นหาชุดติดตั้ง" />
                    <select value={pickedBundleId} onChange={(e) => setPickedBundleId(e.target.value)} className={ain} aria-label="เลือกชุดติดตั้ง">
                      <option value="">{bundlesLoading ? 'กำลังโหลดชุด...' : '— เลือกชุดติดตั้ง —'}</option>
                      {filteredBundleOptions.map((b) => <option key={b.bundleId} value={b.bundleId}>{b.bundleId} · {b.bundleName} · {b.assetIds?.length || 0} ชิ้น</option>)}
                    </select>
                    {!bundlesLoading && bundlesLoaded && filteredBundleOptions.length === 0 && (
                      <div className="text-[12px] text-[var(--tmuted)]">ไม่พบชุดที่ค้นหา — หรือกด "สร้างชุดใหม่ทันที" ด้านล่าง (รหัสชุดรันเลขให้อัตโนมัติ)</div>
                    )}
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                    <button onClick={() => setCreateBundleOpen(true)} className="min-h-11 inline-flex items-center justify-center gap-1.5 px-3 rounded-lg bg-[var(--emerald)] text-white text-[12px] font-semibold hover:bg-[var(--emerald-d)] disabled:opacity-50">
                      <Icon name="add" size="sm" /> สร้างชุดใหม่ทันที
                    </button>
                    <button onClick={confirmAddToBundle} disabled={!pickedBundleId || busy.busy} className="min-h-11 inline-flex items-center justify-center gap-1.5 px-3 rounded-lg bg-[var(--blue)] text-white text-[12px] font-semibold hover:bg-[var(--blue-d)] disabled:opacity-50">
                      <Icon name="folder_open" size="sm" /> เพิ่ม {flowSerials.length} ชิ้นเข้าชุดนี้
                    </button>
                  </div>
                  <button onClick={() => setFlowStep(3)} className="w-full text-[12px] font-semibold text-[var(--tsub)] hover:text-[var(--text)]">ข้าม — ไม่จัดชุดตอนนี้ →</button>
                </StepPanel>
              )}

              {/* Step 3 — โอนย้ายไปฟาร์ม/โรงเรือน/คลัง: ปลายทาง + เพิ่มฟาร์ม/โรงเรือนใหม่ได้ทันทีใน Modal นี้ */}
              {flowStep === 3 && (
                <StepPanel title="③ โอนย้ายไปยังฟาร์ม/โรงเรือน/คลัง" onBack={enterStep2}>
                  {pickedBundleId && (
                    <div className="flex items-start gap-1.5 px-2.5 py-1.5 rounded-lg bg-[var(--amber-l)] border border-[var(--amber-b)] text-[11px] text-[var(--amber-d)]">
                      <Icon name="info" size="sm" />
                      <span>อุปกรณ์อยู่ในชุด {pickedBundleId} แล้ว — การย้ายรายชิ้นจะทำให้ตำแหน่งไม่ตรงกับชุด (ติดตามต่อในหน้าชุดติดตั้ง)</span>
                    </div>
                  )}
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                    <button type="button" onClick={() => setDestMode('stock')} className={`min-h-11 rounded-lg border px-3 text-[12px] font-semibold ${destMode === 'stock' ? 'border-[var(--blue)] bg-[var(--surface)] text-[var(--blue)]' : 'border-[var(--g200)] text-[var(--tsub)]'}`}>คลังกลาง — เก็บต่อที่เดิม</button>
                    <button type="button" onClick={() => setDestMode('farm')} className={`min-h-11 rounded-lg border px-3 text-[12px] font-semibold ${destMode === 'farm' ? 'border-[var(--blue)] bg-[var(--surface)] text-[var(--blue)]' : 'border-[var(--g200)] text-[var(--tsub)]'}`}>ฟาร์ม/โรงเรือน — ส่งปลายทางเลย</button>
                  </div>
                  {destMode === 'farm' && (
                    <>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                        <select value={farmId} onChange={(e) => chooseFarm(e.target.value)} className={ain} aria-label="ฟาร์มปลายทาง">
                          <option value="">{flowFarms.length ? '— เลือกฟาร์ม/ไซต์ปลายทาง —' : 'ยังไม่มีฟาร์มในระบบ'}</option>
                          {flowFarms.map((s) => <option key={s.siteId} value={s.siteId}>{s.siteName} ({s.farmType || '-'})</option>)}
                        </select>
                        <select value={houseId} onChange={(e) => setHouseId(e.target.value)} disabled={!farmId || loadingHouses} className={ain} aria-label="โรงเรือนปลายทาง">
                          <option value="">{loadingHouses ? 'กำลังโหลดโรงเรือน...' : '— ไม่ระบุโรงเรือน —'}</option>
                          {houses.map((h) => <option key={h.houseId} value={h.houseId}>{h.houseId} · {h.houseName}</option>)}
                        </select>
                      </div>
                      {/* ปุ่มทางด่วน — เพิ่มฟาร์ม/โรงเรือนใหม่ได้ทันทีโดยไม่ต้องยกเลิกหรือสลับออกจากหน้าเพิ่มอุปกรณ์
                          (ยังไม่เลือกฟาร์ม → เปิดฟอร์มเพิ่มฟาร์ม | เลือกฟาร์มแล้ว → เปิดฟอร์มเพิ่มโรงเรือนของฟาร์มนั้น) */}
                      <button onClick={() => (farmId ? setHouseAddOpen(!houseAddOpen) : setFarmAddOpen(!farmAddOpen))} className="inline-flex items-center gap-1 px-3 min-h-9 rounded-lg border border-[var(--g300)] bg-[var(--surface)] text-[12px] font-semibold text-[var(--blue)] hover:bg-[var(--blue-l)]">
                        <Icon name="add" size="sm" /> เพิ่มฟาร์ม/โรงเรือนใหม่
                      </button>
                      {farmAddOpen && <FarmInlineAdd mode="panel" open onOpenChange={setFarmAddOpen} onAdded={handleFlowFarmAdded} />}
                      {houseAddOpen && <HouseInlineAdd mode="panel" open onOpenChange={setHouseAddOpen} siteId={farmId} houses={houses} onAdded={handleFlowHouseAdded} />}
                      {farmId && !loadingHouses && houses.length === 0 && (
                        <div className="text-[11px] text-[var(--tmuted)]">ฟาร์มนี้ยังไม่มีโรงเรือน — กดปุ่ม "เพิ่มฟาร์ม/โรงเรือนใหม่" เพื่อสร้างทันที</div>
                      )}
                      <input value={installPoint} onChange={(e) => setInstallPoint(e.target.value)} placeholder="จุดติดตั้ง (ไม่บังคับ) เช่น ชั้น 2, โซนให้อาหาร" className={ain} />
                    </>
                  )}
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                    <button onClick={confirmFlowTransfer} disabled={busy.busy || (destMode === 'farm' && !farmId)} className="min-h-11 inline-flex items-center justify-center gap-1.5 px-3 rounded-lg bg-[var(--blue)] text-white text-[12px] font-semibold hover:bg-[var(--blue-d)] disabled:opacity-50">
                      <Icon name="local_shipping" size="sm" /> {transferFailed.length ? `ลองย้ายใหม่ (${transferFailed.length} ชิ้น)` : `ยืนยันโอนย้าย (${flowSerials.length} ชิ้น)`}
                    </button>
                    <button onClick={() => setFlowStep(4)} className="min-h-11 inline-flex items-center justify-center gap-1.5 px-3 rounded-lg border border-[var(--g300)] text-[12px] font-semibold text-[var(--tsub)]">
                      ข้าม — เก็บที่เดิมไว้ก่อน →
                    </button>
                  </div>
                </StepPanel>
              )}

              {/* Step 4 — พิมพ์ฉลาก QR (Direct-to-Print จาก Episode 1): สรุปก่อนจบ */}
              {flowStep === 4 && (
                <StepPanel title="④ พิมพ์ฉลาก QR" onBack={() => setFlowStep(3)}>
                  <div className="text-[12px] text-[var(--tsub)]">
                    สรุป: ชุด {pickedBundleId ? <b className="text-[var(--text)]">{pickedBundleId}</b> : '— ไม่จัดชุด'} · ปลายทาง {destMode === 'farm' && farmId && transferFailed.length === 0 ? <b className="text-[var(--text)]">{selectedFarm?.siteName || ''}{houseId ? ` › ${flowHouseName || houseId}` : ''}</b> : 'คลังกลาง'}
                  </div>
                  <button onClick={openLabelPrint} className="w-full min-h-11 inline-flex items-center justify-center gap-1.5 px-3 rounded-lg bg-[var(--blue)] text-white text-[13px] font-bold hover:bg-[var(--blue-d)]">
                    <Icon name="print" size="sm" /> เปิดหน้าพิมพ์ฉลาก QR ({flowSerials.length} ดวง)
                  </button>
                  <div className="text-[11px] text-[var(--tmuted)]">เปิดหน้า /qr พร้อมเลือก Serial ที่เพิ่งรับเข้าไว้ให้พิมพ์เลย — หรือกดปุ่มพิมพ์ด้านบน (สรุปผล) ได้ทุกเมื่อ</div>
                </StepPanel>
              )}
            </div>
            <div className="flex justify-end gap-2 mt-3">
              <button
                onClick={startAnother}
                title="เคลียร์ฟอร์มเพื่อเพิ่มอุปกรณ์ถัดไปทันที"
                className="flex items-center gap-1.5 px-4 py-2 rounded-lg border border-[var(--blue-b)] bg-[var(--blue-l)] text-[var(--blue)] text-[13px] font-semibold hover:bg-[var(--blue)] hover:text-white"
              >
                <Icon name="add" size="sm" /> เพิ่มอุปกรณ์อื่นต่อ
              </button>
              <button onClick={() => onDone(result)} className="flex items-center gap-1.5 px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)]">
                <Icon name="check" size="sm" /> เสร็จ
              </button>
            </div>
          </div>
        ) : (
          /* ── ฟอร์ม ── */
          <>
          <div className="p-5 space-y-3">
            {/* ① อุปกรณ์อะไร — พิมพ์แล้วเลือกจาก suggestion หรือสร้างใหม่ */}
            <div className="space-y-1">
              <label className="block text-[12px] font-medium text-[var(--tsub)]">① อุปกรณ์อะไร? (พิมพ์ชื่อหรือ Part Number)</label>
              {newPart ? (
                /* ฟอร์มสร้าง Part ใหม่ — แยก "ชื่ออุปกรณ์" กับ "Part Number (ตัวย่ออังกฤษ)" ออกจากกัน
                   เพราะ Part Number จะถูกใช้สร้าง Serial (SN-<PART>-...) ต้องเป็น A-Z 0-9 เท่านั้น */
                <div className="p-3 rounded-lg border border-[var(--emerald-b)] bg-[var(--emerald-l)] space-y-2">
                  <div className="text-[12px] font-bold text-[var(--emerald-d)]">สร้าง Part ใหม่ — กรอกให้ครบ 2 ช่อง</div>
                  <div className="space-y-1">
                    <label className="block text-[11px] font-medium text-[var(--tsub)]">ชื่ออุปกรณ์ (โชว์ในระบบ — ภาษาไทยได้)</label>
                    <input
                      value={newPart.partName}
                      onChange={(e) => setNewPart({ ...newPart, partName: e.target.value })}
                      placeholder="เช่น เซนเซอร์วัดน้ำ"
                      className={ain}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="block text-[11px] font-medium text-[var(--tsub)]">Part Number (ตัวย่อ — ใช้ใน Serial)</label>
                    <input
                      value={newPart.partNumber}
                      onChange={(e) => setNewPart({ ...newPart, partNumber: e.target.value })}
                      placeholder="เช่น SENWT"
                      className={ain + ' font-mono uppercase'}
                    />
                    <div className="text-[10px] text-[var(--tmuted)]">
                      ตัวอย่าง Serial ที่จะได้: SN-{sanitizePartNumber(newPart.partNumber) || 'PART'}-01102026-0001
                    </div>
                  </div>
                  <div className="flex justify-end gap-2">
                    <button onClick={() => setNewPart(null)} className="px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)]">ยกเลิก</button>
                    <button onClick={confirmNewPart} className="px-3 py-1.5 rounded-lg bg-[var(--emerald)] text-white text-[12px] font-semibold hover:bg-[var(--emerald-d)]">ยืนยัน Part นี้</button>
                  </div>
                </div>
              ) : selected ? (
                <div className="flex items-center justify-between gap-2 px-3 h-9 rounded-lg bg-[var(--blue-l)] border border-[var(--blue-b)]">
                  <span className="text-[13px] font-semibold text-[var(--blue)] truncate">
                    {selected.partNumber} — {selected.partName}
                    {selected.isNew && <span className="ml-1.5 px-1.5 py-px rounded bg-[var(--emerald-l)] text-[var(--emerald-d)] text-[10px] font-bold">ใหม่</span>}
                  </span>
                  <button onClick={() => { setSelected(null); setQuery(''); }} className="text-[11px] text-[var(--tmuted)] hover:text-[var(--red)] shrink-0">เปลี่ยน</button>
                </div>
              ) : (
                <div className="relative">
                  <input
                    value={query}
                    onChange={(e) => setQuery(e.target.value)}
                    placeholder="เช่น เซนเซอร์อุณหภูมิ หรือ SEN.TEMP-RS485"
                    className={ain}
                    autoFocus
                  />
                  {suggestions && !newPart && (
                    <div className="absolute z-10 mt-1 w-full rounded-xl border border-[var(--g200)] bg-[var(--surface)] shadow-[var(--sh-md)] overflow-hidden max-h-56 overflow-y-auto">
                      {suggestions.hits.map((p) => (
                        <button
                          key={p.partNumber}
                          onClick={() => setSelected({ partNumber: sanitizePartNumber(p.partNumber), partName: p.partName || p.partNumber, isNew: false })}
                          className="w-full text-left px-3 py-2 hover:bg-[var(--surface2)] flex items-center justify-between gap-2"
                        >
                          <span className="min-w-0">
                            <span className="text-[13px] font-medium text-[var(--text)] block truncate">{p.partName || p.partNumber}</span>
                            <span className="text-[11px] font-mono text-[var(--tmuted)]">{p.partNumber}</span>
                          </span>
                          <span className="text-[11px] text-[var(--tmuted)] shrink-0">มี {p.totalQty || 0} ชิ้น</span>
                        </button>
                      ))}
                      {suggestions.showCreate && (
                        <button
                          onClick={() => setNewPart({ partName: query.trim(), partNumber: '' })}
                          className="w-full text-left px-3 py-2 hover:bg-[var(--emerald-l)] flex items-center gap-2 border-t border-[var(--g100)]"
                        >
                          <Icon name="add_circle" size="sm" />
                          <span className="text-[13px] font-semibold text-[var(--emerald-d)]">สร้าง Part ใหม่ "{query.trim()}" — กรอก Part Number ต่อ</span>
                        </button>
                      )}
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* ② จำนวนชิ้น */}
            <div className="space-y-1">
              <label className="block text-[12px] font-medium text-[var(--tsub)]">② จะเพิ่มกี่ชิ้น?</label>
              <input type="number" min="1" value={qty} onChange={(e) => setQty(e.target.value)} className={ain} />
            </div>

            {/* ③ Serial preview — คำเดาก่อนส่ง (ตัวจริง backend ตรวจซ้ำให้) */}
            {pn && qtyNum > 0 && (
              <div className="px-3 py-2 rounded-lg bg-[var(--blue-l)] border border-[var(--blue-b)] text-[11px] text-[var(--blue)]">
                <div className="font-bold mb-1">Serial ที่จะได้ (ระบบสร้างให้ + ตรวจซ้ำอัตโนมัติ):</div>
                {serialPreview.map((s) => <div key={s} className="font-mono">{s}</div>)}
                {qtyNum > serialPreview.length && <div>... รวมทั้งหมด {qtyNum} ชิ้น</div>}
                <div className="mt-1 text-[10px] opacity-75">เลขรันต่อท้ายต่อจากที่มีอยู่จริงของวันนี้ — ระบบจะไม่สร้าง Serial ซ้ำเด็ดขาด</div>
              </div>
            )}

            {/* ขั้นสูง — พับเก็บ มี default ให้ครบ ไม่ต้องแตะก็เพิ่มได้ */}
            <button onClick={() => setAdvanced(!advanced)} className="text-[12px] font-semibold text-[var(--blue)] flex items-center gap-1">
              <Icon name="chevron_right" size="xs" className={advanced ? 'rotate-90' : ''} /> รายละเอียดเพิ่ม (สถานะ · ตำแหน่ง · ผู้รับผิดชอบ · ข้อมูลรับเข้า)
            </button>
            {advanced && (
              <div className="space-y-3 pl-3 border-l-2 border-[var(--g100)]">
                <div className="rounded-xl border border-[var(--blue-b)] bg-[var(--blue-l)] p-3 space-y-3">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                    <button type="button" onClick={() => { setBatchMode('new'); setBatchId(''); }} className={`min-h-11 rounded-lg border px-3 text-[12px] font-semibold ${batchMode === 'new' ? 'border-[var(--blue)] bg-[var(--surface)] text-[var(--blue)]' : 'border-[var(--g200)] text-[var(--tsub)]'}`}>สร้างล็อตใหม่</button>
                    <button type="button" onClick={chooseExistingBatchMode} className={`min-h-11 rounded-lg border px-3 text-[12px] font-semibold ${batchMode === 'existing' ? 'border-[var(--blue)] bg-[var(--surface)] text-[var(--blue)]' : 'border-[var(--g200)] text-[var(--tsub)]'}`}>ต่อล็อตเดิม</button>
                  </div>
                  {batchMode === 'existing' && (
                    <div className="space-y-2">
                      <input value={batchQuery} onChange={(e) => setBatchQuery(e.target.value)} placeholder="ค้น Batch ID / PO / Supplier" className={ain + ' min-h-11'} />
                      <select value={batchId} onChange={(e) => {
                        const next = availableBatches.find((batch) => batch.batchId === e.target.value);
                        setBatchId(next?.batchId || '');
                        if (next?.receivedAt) setReceivedAt(toLocalDateTimeValue(new Date(next.receivedAt)));
                        setPoNumber(next?.poNumber || '');
                        setSupplier(next?.supplier || '');
                      }} className={ain + ' min-h-11'}>
                        <option value="">{batchesLoading ? 'กำลังโหลดล็อต...' : 'เลือก Batch ID'}</option>
                        {filteredBatches.map((batch) => <option key={batch.batchId} value={batch.batchId}>{batch.batchId} · PO {batch.poNumber || '-'} · {batch.count} ชิ้น · {formatReceivedAt(batch.receivedAt)}</option>)}
                      </select>
                      {!batchesLoading && filteredBatches.length === 0 && <div className="text-[12px] text-[var(--tmuted)]">ไม่พบล็อตที่ค้นหา</div>}
                    </div>
                  )}
                  <div className="text-[12px] font-semibold text-[var(--blue)]">ข้อมูลรับเข้า · สร้างล็อตให้อัตโนมัติ</div>
                  <div className="space-y-1"><label className="block text-[12px] font-medium text-[var(--tsub)]">วันที่และเวลารับเข้า</label><input type="datetime-local" value={receivedAt} disabled={batchMode === 'existing' && !!selectedBatch} onChange={(e) => setReceivedAt(e.target.value)} className={ain + ' min-h-11'} /></div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div className="space-y-1"><label className="block text-[12px] font-medium text-[var(--tsub)]">เลข PO (ไม่บังคับ)</label><input value={poNumber} disabled={batchMode === 'existing' && !!selectedBatch} onChange={(e) => setPoNumber(e.target.value)} maxLength={100} className={ain + ' min-h-11'} /></div>
                    <div className="space-y-1"><label className="block text-[12px] font-medium text-[var(--tsub)]">Supplier (ไม่บังคับ)</label><input value={supplier} disabled={batchMode === 'existing' && !!selectedBatch} onChange={(e) => setSupplier(e.target.value)} maxLength={100} className={ain + ' min-h-11'} /></div>
                  </div>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div className="space-y-1">
                    <label className="block text-[12px] font-medium text-[var(--tsub)]">สถานะ</label>
                    <select value={status} onChange={(e) => setStatus(e.target.value)} className={ain}>
                      <option>ใช้งานได้</option><option>สำรอง</option><option>ส่งซ่อม</option><option>ชำรุด/สูญหาย</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="block text-[12px] font-medium text-[var(--tsub)]">ไซต์งาน</label>
                    <select value={siteName} onChange={(e) => setSiteName(e.target.value)} className={ain}>
                      <option value="Intranin">คลังกลาง (บริษัท Intranin)</option>
                      {sites.filter((s) => s.siteName && s.siteName !== 'Intranin').map((s) => <option key={s.siteId} value={s.siteName}>{s.siteName}</option>)}
                    </select>
                  </div>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div className="space-y-1">
                    <label className="block text-[12px] font-medium text-[var(--tsub)]">จุดติดตั้ง</label>
                    <select value={location} onChange={(e) => setLocation(e.target.value)} className={ain}>
                      <option value="Stock">คลังกลาง</option>
                      {sites.filter((s) => s.siteName && s.siteName !== 'Intranin').map((s) => <option key={s.siteId} value={s.siteName}>{s.siteName}</option>)}
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="block text-[12px] font-medium text-[var(--tsub)]">ผู้รับผิดชอบ</label>
                    <input value={user} onChange={(e) => setUser(e.target.value)} placeholder="เว้นว่าง = ชื่อผู้ใช้ของคุณ" className={ain} />
                  </div>
                </div>
              </div>
            )}
          </div>
          <div className="flex justify-end gap-2 px-5 py-4 border-t border-[var(--g100)] bg-[var(--surface2)] rounded-b-2xl">
            <button onClick={onClose} className="px-4 py-2 rounded-lg border border-[var(--g300)] text-[13px] text-[var(--tsub)]">ยกเลิก</button>
            <button onClick={submit} disabled={busy.busy || !selected} className="flex items-center gap-1.5 px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)] disabled:opacity-50">
              <Icon name="add" size="sm" /> เพิ่ม {qtyNum || 0} ชิ้น
            </button>
          </div>
          </>
        )}
        <ToastHost />
        {/* Episode 2 — สร้างชุดใหม่ทันทีจาก Step 2 (nested modal — wrapper กันคลิกทะลุปิด Modal หลัก) */}
        {result && createBundleOpen && (
          <div onClick={(e) => e.stopPropagation()}>
            <CreateBundleModal
              existingIds={bundleOptions.map((b) => b.bundleId)}
              saving={busy.busy}
              onClose={() => setCreateBundleOpen(false)}
              onSubmit={handleCreateBundleNow}
            />
          </div>
        )}
        <BusyOverlay label={busy.busyLabel} />
      </div>
    </div>
  );
}
