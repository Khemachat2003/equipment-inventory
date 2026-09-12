import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { BrowserRouter } from 'react-router-dom'
import './index.css'
import App from './App.jsx'
import { AuthProvider } from './context/AuthContext.jsx'

// ── Material Symbols: แสดง icon เมื่อ font พร้อม ──
// index.html เปิดด้วย class "icons-loading" (ซ่อน .msi ไว้) → ถอด class
// เมื่อ font subset โหลดเสร็จ หรือหลัง timeout 8s (กัน font พังแล้ว icon หายตลอด)
const revealIcons = () => document.documentElement.classList.remove('icons-loading');
if (document.fonts?.load) {
  Promise.race([
    document.fonts.load('400 1em "Material Symbols Rounded"', 'inventory_2'),
    new Promise((resolve) => setTimeout(resolve, 8000)),
  ]).then(revealIcons).catch(revealIcons);
} else {
  setTimeout(revealIcons, 2000);
}

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <BrowserRouter>
      <AuthProvider>
        <App />
      </AuthProvider>
    </BrowserRouter>
  </StrictMode>,
)
