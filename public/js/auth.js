// public/js/auth.js
async function login() {
  const btn = document.getElementById("loginBtn");
  btn.textContent = "กำลังเข้าสู่ระบบ...";
  btn.disabled = true;
  document.getElementById("loginError").textContent = "";
  try {
    const u = document.getElementById("username").value.trim();
    const p = document.getElementById("password").value.trim();
    if (!u || !p) {
      document.getElementById("loginError").textContent = "กรุณากรอก Username และ Password";
      return;
    }
    const res = await fetch("/api/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      credentials: "include",
      body: JSON.stringify({ username: u, password: p })
    });
    const r = await res.json();
    if (r.success) {
      showApp();
    } else {
      document.getElementById("loginError").textContent = r.error || "Username หรือ Password ไม่ถูกต้อง";
    }
  } catch (e) {
    document.getElementById("loginError").textContent = "ไม่สามารถเชื่อมต่อได้";
  } finally {
    btn.textContent = "เข้าสู่ระบบ";
    btn.disabled = false;
  }
}

async function logout() {
  try {
    await fetch("/api/logout", { method: "POST", credentials: "include" });
  } catch (e) {}
  location.reload();
}

function showApp() {
  document.getElementById("loginPage").style.display = "none";
  document.getElementById("mainApp").style.display = "block";
  initApp();
}

async function checkAuth() {
  try {
    const d = await (await fetch("/api/check-auth", { credentials: "include" })).json();
    if (d.loggedIn) showApp();
  } catch (e) {}
  document.body.style.visibility = "visible";
}

async function loadCurrentUser() {
  try {
    const res = await fetch("/api/me", { credentials: "include" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const d = await res.json();
    const name = d.username || "User";
    document.getElementById("currentUser").textContent = name;
    document.getElementById("settingsUser").textContent = name;
    const av = document.getElementById("userAvatar");
    if (av) av.textContent = name[0]?.toUpperCase() || "U";
  } catch (e) {
    console.warn("Load user fallback:", e);
    document.getElementById("currentUser").textContent = "User";
    document.getElementById("settingsUser").textContent = "User";
  }
}