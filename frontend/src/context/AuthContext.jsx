import { createContext, useContext, useEffect, useState, useCallback } from 'react';
import axios from 'axios';

const AuthContext = createContext(null);

axios.defaults.withCredentials = true;

export function AuthProvider({ children }) {
  const [user, setUser] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [portalHomeUrl, setPortalHomeUrl] = useState('/');
  const [portalMode, setPortalMode] = useState(false);

  // Check session on mount (cookie-based auth via /api/check-auth + /api/me)
  const checkAuth = useCallback(async () => {
    try {
      const { data } = await axios.get('/api/check-auth');
      if (data.loggedIn) {
        const { data: me } = await axios.get('/api/me');
        setUser({ username: me.username || data.username, role: me.role || data.role || 'user', portal: !!(me.portal || data.portal) });
      } else {
        setUser(null);
      }
    } catch (e) {
      setUser(null);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    checkAuth();
    axios.get('/api/portal-config').then(({ data }) => {
      setPortalMode(!!data.portalMode);
      setPortalHomeUrl(data.homeUrl || '/');
    }).catch(() => {});
  }, [checkAuth]);

  const login = useCallback(async (username, password) => {
    setError('');
    try {
      const { data } = await axios.post('/api/login', { username: username.trim(), password: password.trim() });
      if (data.success) {
        const { data: me } = await axios.get('/api/me');
        setUser({ username: me.username || username, role: me.role || data.role || 'user' });
        return { ok: true };
      }
      setError(data.error || 'Username หรือ Password ไม่ถูกต้อง');
      return { ok: false, error: data.error };
    } catch (e) {
      setError('ไม่สามารถเชื่อมต่อได้');
      return { ok: false, error: 'ไม่สามารถเชื่อมต่อได้' };
    }
  }, []);

  const logout = useCallback(async () => {
    try {
      await axios.post('/api/logout');
    } catch (e) {}
    setUser(null);
  }, []);

  return (
    <AuthContext.Provider value={{ user, loading, error, login, logout, checkAuth, portalHomeUrl, portalMode }}>
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  return useContext(AuthContext);
}
