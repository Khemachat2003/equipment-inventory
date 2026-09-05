import React from 'react';
import { Routes, Route, Navigate } from 'react-router-dom';
import Layout from './components/layout/Layout.jsx';
import Placeholder from './pages/Placeholder.jsx';
import Login from './pages/Login.jsx';
import Dashboard from './pages/Dashboard.jsx';
import Stock from './pages/Stock.jsx';
import { useAuth } from './context/AuthContext.jsx';

function Protected({ children }) {
  const { user, loading } = useAuth();
  if (loading) return <FullLoader />;
  if (!user) return <Navigate to="/login" replace />;
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
      <Route path="/login" element={user ? <Navigate to="/" replace /> : <Login />} />

      <Route
        element={
          <Protected>
            <Layout />
          </Protected>
        }
      >
        <Route path="/" element={<Dashboard />} />
        <Route path="/stock" element={<Stock />} />

        {/* Pages still served from the original static HTML — redirect to them */}
        <Route path="/asset" element={<Navigate to="/asset.html" replace />} />
        <Route path="/bundle" element={<Navigate to="/index.html" replace />} />
        <Route path="/farm" element={<Navigate to="/index.html" replace />} />
        <Route path="/history" element={<Navigate to="/index.html" replace />} />
        <Route path="/report" element={<Navigate to="/index.html" replace />} />
        <Route path="/settings" element={<Navigate to="/index.html" replace />} />
      </Route>
    </Routes>
  );
}
