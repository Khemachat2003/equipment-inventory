import React from 'react';
import { Routes, Route, Navigate } from 'react-router-dom';
import Layout from './components/layout/Layout.jsx';
import ScanPage from './pages/ScanPage.jsx';
import TracePage from './pages/TracePage.jsx';
import QrPage from './pages/QrPage.jsx';
import Login from './pages/Login.jsx';
import Dashboard from './pages/Dashboard.jsx';
import Stock from './pages/Stock.jsx';
import Asset from './pages/Asset.jsx';
import Bundle from './pages/Bundle.jsx';
import Farm from './pages/Farm.jsx';
import History from './pages/History.jsx';
import Report from './pages/Report.jsx';
import Settings from './pages/Settings.jsx';
import AuditLog from './pages/admin/AuditLog.jsx';
import AdminTools from './pages/admin/AdminTools.jsx';
import UserManagement from './pages/admin/UserManagement.jsx';
import { useAuth } from './context/AuthContext.jsx';

function Protected({ children }) {
  const { user, loading } = useAuth();
  if (loading) return <FullLoader />;
  if (!user) return <Navigate to="/login" replace />;
  return children;
}

// Admin-only route guard — ถ้าไม่ใช่ admin redirect ไปหน้าแรก
function AdminRoute({ children }) {
  const { user, loading } = useAuth();
  if (loading) return <FullLoader />;
  if (!user) return <Navigate to="/login" replace />;
  if (user.role !== 'admin') return <Navigate to="/" replace />;
  return children;
}

function FullLoader() {
  return (
    <div className="min-h-screen flex items-center justify-center bg-[var(--bg)]">
      <div className="flex flex-col items-center gap-3 text-[var(--tmuted)]">
        <span className="w-8 h-8 border-2 border-[var(--g300)] border-t-[var(--blue)] rounded-full animate-spin" />
        <span className="text-sm">กำลังโหลด...</span>
      </div>
    </div>
  );
}

export default function App() {
  const { user, loading } = useAuth();

  // If logged in and hitting /login, go to dashboard
  if (loading) return <FullLoader />;

  return (
    <Routes>
      {/* portal user (role:user) ยังเปิดหน้า login ได้ เพื่อให้ admin login แยก */}
      <Route
        path="/login"
        element={user && user.role === 'admin' ? <Navigate to="/" replace /> : <Login />}
      />

      <Route
        element={
          <Protected>
            <Layout />
          </Protected>
        }
      >
        <Route path="/" element={<Dashboard />} />
        <Route path="/stock" element={<Stock />} />
        <Route path="/asset" element={<Asset />} />
        <Route path="/bundle" element={<Bundle />} />
        <Route path="/farm" element={<Farm />} />
        <Route path="/history" element={<History />} />
        <Route path="/report" element={<Report />} />
        <Route path="/scan" element={<ScanPage />} />
        <Route path="/qr" element={<QrPage />} />

        {/* Admin-only pages */}
        <Route
          path="/admin-tools"
          element={
            <AdminRoute>
              <AdminTools />
            </AdminRoute>
          }
        />
        <Route
          path="/audit"
          element={
            <AdminRoute>
              <AuditLog />
            </AdminRoute>
          }
        />
        <Route
          path="/users"
          element={
            <AdminRoute>
              <UserManagement />
            </AdminRoute>
          }
        />
        <Route
          path="/settings"
          element={
            <AdminRoute>
              <Settings />
            </AdminRoute>
          }
        />
      </Route>

      {/* Public trace page — เปิดได้ไม่ต้องล็อกอิน (ประวัติอ้างอิงจาก QR sticker / ลิงก์ trace) */}
      <Route path="/trace/:serial" element={<TracePage />} />
      <Route path="/trace" element={<TracePage />} />

      {/* path เก่า/ไม่รู้จัก (เช่น /app/stock ที่ใช้ก่อน migrate) → เด้งหน้าแรก */}
      <Route path="*" element={<Navigate to="/" replace />} />
    </Routes>
  );
}