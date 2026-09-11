import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { defineConfig } from 'vite'

// Build output goes into ./dist (separate from the backend's public/).
// SPA ถูก mount ที่ root "/" — base จึงเป็น "/" (asset path = /assets/...)
export default defineConfig({
  base: '/',
  plugins: [react(), tailwindcss()],
  // Vitest (จัดการโดย npm run test) — ใช้ jsdom สำหรับ render component
  test: {
    environment: 'jsdom',
    globals: true,
  },
  server: {
    proxy: {
      '/api': {
        target: 'http://localhost:3000',
        changeOrigin: true,
      },
      '/auth': {
        target: 'http://localhost:3000',
        changeOrigin: true,
      },
    },
  },
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    // ยกจาก 700 → 1200 กัน warning "chunk size" ของ bundle หลัก (ไม่กระทบ production)
    // ถ้าอนาคต bundle ใหญ่ขึ้นมากควรแยก chunk ด้วย manualChunks แทนการยก limit
    chunkSizeWarningLimit: 1200,
  },
})
