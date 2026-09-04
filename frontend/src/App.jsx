import { Routes, Route, Navigate } from 'react-router-dom';
import Layout from './components/layout/Layout.jsx';
import Placeholder from './pages/Placeholder.jsx';

export default function App() {
  return (
    <Routes>
      <Route element={<Layout />}>
        {/* Pages already migrated to React */}
        <Route path="/" element={<Placeholder title="Dashboard" note="หน้านี้ยังใช้ API เดิม — กำลังย้าย" />} />

        {/* Pages still served from the original static HTML — redirect to them */}
        <Route path="/stock" element={<Navigate to="/index.html" replace />} />
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
