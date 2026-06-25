// public/js/common.js
function openModal(id) {
  document.getElementById(id).classList.add("open");
}
function closeModal(id) {
  document.getElementById(id).classList.remove("open");
}
document.querySelectorAll(".mo").forEach(el => {
  el.addEventListener("click", e => {
    if (e.target === el) el.classList.remove("open");
  });
});

function getStatusBadge(s = "") {
  const t = s.toLowerCase();
  if (t.includes("ซ่อม")) return `<span class="badge bg-red">${s}</span>`;
  if (t.includes("สำรอง")) return `<span class="badge bg-amber">${s}</span>`;
  if (t.includes("ชำรุด") || t.includes("สูญ")) return `<span class="badge bg-gray">${s}</span>`;
  return `<span class="badge bg-green">${s}</span>`;
}

function getInitials(name = "") {
  const p = name.trim().split(" ").filter(Boolean);
  if (!p.length) return "?";
  return p.length > 1 ? (p[0][0] + p[p.length - 1][0]).toUpperCase() : p[0].slice(0, 2).toUpperCase();
}

function toast(msg, type = 'ok') {
  const el = document.getElementById('toast');
  if (!el) return;
  el.textContent = msg;
  el.style.borderColor = type === 'ok' ? 'rgba(16,185,129,.4)' : 'rgba(239,68,68,.4)';
  el.style.color = type === 'ok' ? 'var(--green)' : 'var(--red)';
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 2500);
}