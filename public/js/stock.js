// public/js/stock.js
async function loadStock() {
  try {
    const r = await fetch("/api/stock");
    stockData = await r.json();
    renderStockTable(stockData);
  } catch (e) {}
}

function renderStockTable(data) {
  const tb = document.getElementById("stockBody");
  if (!data.length) {
    tb.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:40px;color:var(--tmuted)">ไม่มีข้อมูล</td></tr>`;
    return;
  }
  tb.innerHTML = data.map(item => {
    const ext = item.ext || "jpg";
    const img = `https://cdn.jsdelivr.net/gh/Khemachat2003/stock-image@main/images/${item.code}.${ext}?v=${IMAGE_CACHE_VERSION}`;
    const oq = parseInt(item.office) || 0;
    return `<tr ${oq < 1 ? 'class="row-out"' : ""}>
      <td><span style="font-family:monospace;font-weight:600">${item.code}</span></td>
      <td><img src="${img}" width="46" height="46" loading="lazy" style="object-fit:cover;border-radius:8px;border:1px solid var(--g200)" onerror="this.src='/image/noimage.jpg'"></td>
      <td style="font-weight:500">${item.name}</td>
      <td>${item.total}</td>
      <td>${oq < 1 ? '<span class="out-text">🔴 หมด</span>' : oq}</td>
      <td>${item.site}</td>
      <td><div class="sc-row">
        <input type="number" min="1" id="qty-${item.code}" class="qty-in" placeholder="จำนวน">
        <select id="type-${item.code}" class="typ-sel"><option value="เบิก">เบิก</option><option value="คืน">คืน</option></select>
        <button class="btn btn-blue btn-sm" onclick="event.stopPropagation();submitRowTransfer('${item.code}','${item.name}')">✔</button>
        <button class="btn btn-out btn-sm" onclick="event.stopPropagation();editTotal('${item.code}',${item.total})">✏️</button>
      </div></td>
    </tr>`;
  }).join("");
}

function filterTable() {
  const k = document.getElementById("search").value.toLowerCase();
  renderStockTable(stockData.filter(i => i.code.toLowerCase().includes(k) || i.name.toLowerCase().includes(k)));
}

async function editTotal(code, cur) {
  const n = prompt("แก้ไขจำนวนทั้งหมด:", cur);
  if (n === null) return;
  const r = await (await fetch("/api/update-total", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ code, newTotal: parseInt(n) })
  })).json();
  if (r.error) alert(r.error);
  else { await loadStock(); await loadDashboard(); }
}

async function addItem() {
  const code = document.getElementById("newCode").value.trim(),
    name = document.getElementById("newName").value.trim(),
    qty = parseInt(document.getElementById("newQty").value),
    file = document.getElementById("newImage").files[0];
  if (!code || !name || !qty || !file) { alert("กรอกข้อมูลและเลือกรูปให้ครบ"); return; }
  const ext = file.name.split(".").pop().toLowerCase();
  if (!["jpg", "jpeg", "png"].includes(ext)) { alert("รองรับ JPG/PNG เท่านั้น"); return; }
  const reader = new FileReader();
  reader.onload = async e => {
    try {
      const up = await (await fetch("/upload-image", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fileName: code + "." + ext, base64: e.target.result })
      })).json();
      if (up.success) {
        await fetch("/api/add-item", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ code, name, total: qty, office: 0, site: 0, ext })
        });
        closeModal("addStockModal");
        ["newCode", "newName", "newQty"].forEach(id => document.getElementById(id).value = "");
        await loadStock();
        await loadDashboard();
        alert("เพิ่มอุปกรณ์สำเร็จ");
      } else alert("อัปโหลดรูปไม่สำเร็จ");
    } catch (e) { alert("เกิดข้อผิดพลาด"); }
  };
  reader.readAsDataURL(file);
}

async function submitRowTransfer(code, name) {
  const qty = parseInt(document.getElementById(`qty-${code}`).value),
    type = document.getElementById(`type-${code}`).value;
  if (!qty || qty <= 0) { alert("กรอกจำนวนให้ถูกต้อง"); return; }
  const r = await (await fetch("/api/transfer", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ code, name, qty, type })
  })).json();
  if (r.error) alert(r.error);
  else { await loadStock(); await loadDashboard(); await loadHistory(); }
}

// RETURN
async function _openReturnModal() {
  try {
    const d = await (await fetch("/api/get-site-items")).json();
    document.getElementById("returnList").innerHTML = (d.items || []).map(i =>
      `<label class="ret-item">
        <input type="checkbox" data-code="${i.code}" data-name="${i.name}" data-qty="${i.qty}">
        <span>${i.name}</span>
        <span class="ret-qty">${i.qty} ชิ้น</span>
      </label>`
    ).join("") || `<div style="text-align:center;padding:30px;color:var(--tmuted)">📦 ไม่มีอุปกรณ์ใน Site</div>`;
    openModal("returnModal");
  } catch (e) { alert("โหลดข้อมูลไม่ได้"); }
}

function selectAllReturn() {
  document.querySelectorAll("#returnList input[type='checkbox']").forEach(c => c.checked = true);
}

function unselectAllReturn() {
  document.querySelectorAll("#returnList input[type='checkbox']").forEach(c => c.checked = false);
}

async function confirmSelectedReturn() {
  const items = [...document.querySelectorAll("#returnList input:checked")].map(c => ({ code: c.dataset.code, qty: parseInt(c.dataset.qty) }));
  if (!items.length) { alert("กรุณาเลือกรายการ"); return; }
  const d = await (await fetch("/api/return-selected-site", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ items })
  })).json();
  if (d.success) { closeModal("returnModal"); loadStock(); } else alert("เกิดข้อผิดพลาด");
}