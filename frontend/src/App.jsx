import React, { lazy, Suspense } from 'react';
import { Routes, Route, Navigate } from 'react-router-dom';
import Layout from './components/layout/Layout.jsx';
import { useAuth } from './context/AuthContext.jsx';

// Lazy load หน้าต่างๆ — แต่ละหน้าเป็น chunk แยก โหลดเฉพาะที่เข้า (ลด bundle หลัก / แก้ warning chunk size)
const ScanPage = lazy(() => import('./pages/ScanPage.jsx'));
const TracePage = lazy(() => import('./pages/TracePage.jsx'));
const QrPage = lazy(() => import('./pages/QrPage.jsx'));
const Login = lazy(() => import('./pages/Login.jsx'));
const Dashboard = lazy(() => import('./pages/Dashboard.jsx'));
const Stock = lazy(() => import('./pages/Stock.jsx'));
const Asset = lazy(() => import('./pages/Asset.jsx'));
const Bundle = lazy(() => import('./pages/Bundle.jsx'));
const Farm = lazy(() => import('./pages/Farm.jsx'));
const History = lazy(() => import('./pages/History.jsx'));
const Report = lazy(() => import('./pages/Report.jsx'));
const Settings = lazy(() => import('./pages/Settings.jsx'));
const AuditLog = lazy(() => import('./pages/admin/AuditLog.jsx'));
const AdminTools = lazy(() => import('./pages/admin/AdminTools.jsx'));
const BackupView = lazy(() => import('./pages/admin/BackupView.jsx'));
const UserManagement = lazy(() => import('./pages/admin/UserManagement.jsx'));

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
    <Suspense fallback={<FullLoader />}>
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
        <Route
          path="/backup"
          element={
            <AdminRoute>
              <BackupView />
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
    </Suspense>
  );
}