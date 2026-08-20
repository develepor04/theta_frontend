import React, { useEffect } from 'react';
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import { Toaster } from 'react-hot-toast';
import useStore from './store/useStore';
import { authService } from './services/api';
import LoginPage from './pages/LoginPage';
import SignUpPage from './pages/SignUpPage';
import ForgotPasswordPage from './pages/ForgotPasswordPage';
import FirstLoginPasswordPage from './pages/FirstLoginPasswordPage';
import Dashboard from './pages/Dashboard';
import History from './pages/History';
import Output from './pages/Output';
import Chat from './pages/Chat';
import PredictionWhatIF from './pages/PredictionWhatIF2';
import Knowledge from './pages/Knowledge';
import IntelligenceHub from './pages/IntelligenceHub';
import PredictiveAnalysis from './pages/PredictiveAnalysis';
import BaseInsights from './pages/BaseInsights';
import ThetaEngage from './pages/ThetaEngage';
import Profile from './pages/Profile';
import Subscription from './pages/Subscription';
import CompanyAdminPage from './pages/CompanyAdminPage';
import SuperAdminPage from './pages/SuperAdminPage';
import AuthTransfer from './pages/AuthTransfer';
import ProjectIntelligenceDashboard from './pages/ProjectIntelligenceDashboard';
import ThetaSharePage from './pages/ThetaSharePage';

function App() {
  const { isAuthenticated, user, mustChangePassword, authReady, setAuthReady, setUser, logout } = useStore();

  const loginRedirect = () => {
    const next = `${window.location.pathname}${window.location.search}`;
    const safe = next.startsWith('/') && !next.startsWith('//') && next !== '/login';
    return safe ? `/login?next=${encodeURIComponent(next)}` : '/login';
  };

  const renderProtected = (element) => {
    if (!isAuthenticated) return <Navigate to={loginRedirect()} />;
    if (mustChangePassword || user?.must_change_password) return <Navigate to="/first-login-password" />;
    if (
      user?.role === 'super_admin' &&
      !window.location.pathname.startsWith('/theta-share/') &&
      !window.location.pathname.startsWith('/s/')
    ) {
      return <Navigate to="/super-admin" />;
    }
    return element;
  };

  const renderSubscriptionRoute = () => {
    if (!isAuthenticated) return <Navigate to="/login" />;
    if (mustChangePassword || user?.must_change_password) return <Navigate to="/first-login-password" />;
    return <Subscription />;
  };

  const renderOverviewRoute = () => {
    if (!isAuthenticated) return <Navigate to="/login" />;
    if (mustChangePassword || user?.must_change_password) return <Navigate to="/first-login-password" />;
    if (user?.role === 'super_admin') return <Navigate to="/super-admin" />;
    if (user?.company_name?.toLowerCase() === 'descon') return <Navigate to="/reports" />;
    return <ProjectIntelligenceDashboard />;
  };

  const renderKnowledgeRoute = () => {
    if (!isAuthenticated) return <Navigate to="/login" />;
    if (mustChangePassword || user?.must_change_password) return <Navigate to="/first-login-password" />;
    if (!authReady || !user) return null; // still loading, wait

    const allowed = ['admin', 'company_admin', 'super_admin'].includes(user.role);

    return allowed ? <Knowledge /> : <Navigate to="/dashboard" />;
  };

  const renderCompanyAdminRoute = () => {
    if (!isAuthenticated) return <Navigate to="/login" />;
    if (mustChangePassword || user?.must_change_password) return <Navigate to="/first-login-password" />;
    if (!authReady || !user) return null;
    const allowed = ['company_admin', 'admin', 'super_admin'].includes(user.role);
    return allowed ? <CompanyAdminPage /> : <Navigate to="/dashboard" />;
  };

  const renderSuperAdminRoute = () => {
    if (!isAuthenticated) return <Navigate to="/login" />;
    if (mustChangePassword || user?.must_change_password) return <Navigate to="/first-login-password" />;
    if (!authReady || !user) return null;
    return user.role === 'super_admin' ? <SuperAdminPage /> : <Navigate to="/dashboard" />;
  };

  const renderFirstLoginRoute = () => {
    if (!isAuthenticated) return <Navigate to="/login" />;
    if (mustChangePassword || user?.must_change_password) return <FirstLoginPasswordPage />;
    return <Navigate to="/dashboard" />;
  };

  // MSAL is initialized once in main.jsx before React mounts (avoids
  // interaction_in_progress from duplicate init + StrictMode double-effects).

  useEffect(() => {
    // Load user data if authenticated
    const loadUserData = async () => {
      if (isAuthenticated) {
        try {
          const userData = await authService.getCurrentUser();
          setUser(userData);
        } catch (error) {
          console.error('Failed to load user data:', error);
          logout();
        } finally {
          setAuthReady(true);
        }
      } else {
        setAuthReady(true);
      }
    };

    loadUserData();
  }, [isAuthenticated, setAuthReady, setUser, logout]);

  if (!authReady) {
    return (
      <div style={{ minHeight: '100vh', display: 'grid', placeItems: 'center', fontFamily: 'sans-serif' }}>
        Loading...
      </div>
    );
  }

  return (
    <Router 
    >
      <Toaster
        position="top-right"
        toastOptions={{
          duration: 4000,
          style: {
            background: '#fff',
            color: '#323338',
            padding: '16px',
            borderRadius: '12px',
            boxShadow: '0 8px 24px rgba(0, 0, 0, 0.12)',
            fontSize: '14px',
            fontWeight: '500',
          },
          success: {
            iconTheme: {
              primary: '#00ca72',
              secondary: '#fff',
            },
          },
          error: {
            iconTheme: {
              primary: '#e44258',
              secondary: '#fff',
            },
          },
        }}
      />
      <Routes>
        <Route path="/" element={<LoginPage />} />
        <Route path="/login" element={<LoginPage />} />
        <Route path="/auth-transfer" element={<AuthTransfer />} />
        <Route path="/signup" element={<SignUpPage />} />
        <Route path="/forgot-password" element={<ForgotPasswordPage />} />
        <Route path="/reset-password" element={<ForgotPasswordPage />} />
        <Route
          path="/first-login-password"
          element={renderFirstLoginRoute()}
        />

        {/* Protected Routes */}
        <Route
          path="/overview"
          element={renderOverviewRoute()}
        />
        <Route
          path="/dashboard"
          element={renderProtected(<Dashboard />)}
        />
        <Route
          path="/s/:fileId/:linkToken"
          element={<ThetaSharePage />}
        />
        <Route
          path="/theta-share/:shareToken"
          element={<ThetaSharePage />}
        />
        <Route
          path="/reports"
          element={renderProtected(<History />)}
        />
        <Route
          path="/base-insights"
          element={renderProtected(<BaseInsights />)}
        />
       
        <Route
          path="/chat"
          element={renderProtected(<Chat />)}
        />
        <Route
          path="/predictive-whatif"
          element={renderProtected(<PredictionWhatIF />)}
        />
        <Route
          path="/knowledge-base"
          element={renderKnowledgeRoute()}
        />
        <Route
          path="/output/:id"
          element={renderProtected(<Output />)}
        />
        <Route
          path="/intelligence-hub"
          element={renderProtected(<IntelligenceHub />)}
        />
        <Route
          path="/predictive-analysis"
          element={renderProtected(<PredictiveAnalysis />)}
        />
        <Route
          path="/theta-engage"
          element={renderProtected(<ThetaEngage />)}
        />
        <Route
          path="/profile"
          element={renderProtected(<Profile />)}
        />
        <Route
          path="/subscription"
          element={renderSubscriptionRoute()}
        />
        <Route
          path="/company-admin"
          element={renderCompanyAdminRoute()}
        />
        <Route
          path="/super-admin"
          element={renderSuperAdminRoute()}
        />
        {/* Fallback */}
        <Route path="*" element={<Navigate to="/login" />} />
      </Routes>
    </Router>
  );
}

export default App;
