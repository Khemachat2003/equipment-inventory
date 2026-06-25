// public/js/history.js
async function loadHistory() {
  const s = document.getElementById("historyStart").value;
  const e = document.getElementById("historyEnd").value;
  let url = "/api/history";
  if (s || e) url += `?start=${s}&end=${e}`;
  const tb = document.getElementById("historyBody");
  try {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const h = await res.json();
    if (!h.length) {
      tb.innerHTML = `<tr><td colspan="8" style="text-align:center;padding:40px;color:var(--tmuted)">📭 ไม่พบข้อมูลในช่วงเวลานี้</td></tr>`;
      return;
    }
    const ts = t => t === "เบิก" ? `<span class="badge bg-blue">${t}</span>` :
      t === "คืน" ? `<span class="badge bg-green">${t}</span>` :
      `<span class="badge bg-gray">${t}</span>`;
    tb.innerHTML = h.map(x => `<tr>
      <td>${x.date || "-"}</td>
      <td><span style="font-family:monospace">${x.code || "-"}</span></td>
      <td>${x.name || "-"}</td>
      <td>${ts(x.type)}</td>
      <td style="text-align:center">${x.qty || "-"}</td>
      <td>${x.from || "-"}</td>
      <td>${x.to || "-"}</td>
      <td>${x.user || "-"}</td>
    </tr>`).join("");
  } catch (err) {
    console.error("History error:", err);
    tb.innerHTML = `<tr><td colspan="8" style="text-align:center;padding:40px;color:var(--red)">⚠️ ไม่สามารถโหลดประวัติได้ (${err.message})</td></tr>`;
  }
}

function clearHistoryFilter() {
  document.getElementById("historyStart").value = "";
  document.getElementById("historyEnd").value = "";
  loadHistory();
}

function exportPDF() {
  const s = document.getElementById("startDate").value,
    e = document.getElementById("endDate").value;
  if (!s || !e) { alert("กรุณาเลือกช่วงวันที่"); return; }
  const data = {
    title: document.getElementById("title").value || "รายงาน",
    locations: document.getElementById("locations").value,
    vehicle: document.getElementById("vehicle").value || "-",
    startDate: s,
    endDate: e,
    employeeCount: document.getElementById("employeeCount").value || "0",
    employees: document.getElementById("employees").value,
    reportType: document.getElementById("reportType").value
  };
  const form = document.createElement("form");
  form.method = "POST";
  form.action = "/api/export-history";
  form.target = "_blank";
  Object.keys(data).forEach(k => {
    const i = document.createElement("input");
    i.type = "hidden";
    i.name = k;
    i.value = data[k];
    form.appendChild(i);
  });
  document.body.appendChild(form);
  form.submit();
  document.body.removeChild(form);
}

function openDatabase() {
  window.open("https://docs.google.com/spreadsheets/d/1xAqS4dwT91fGVqTp2b3z6VWlXug28ilUHYVJ_tHe3QE", "_blank");
}