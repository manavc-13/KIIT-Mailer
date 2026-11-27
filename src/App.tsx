// Routes and auth guards. Every page except /login sits behind the login wall;
// /admin/* additionally requires the admin role. Users with a temporary password
// are held on /change-password until they set their own.
import { lazy, Suspense } from 'react';
import { Navigate, Outlet, Route, Routes, useLocation } from 'react-router-dom';
import { Loader2 } from 'lucide-react';
import { useAuth } from '@/lib/auth';
import { AppShell } from '@/components/AppShell';
import LoginPage from '@/pages/LoginPage';
import ChangePasswordPage from '@/pages/ChangePasswordPage';

const ComposePage = lazy(() => import('@/pages/ComposePage'));
const SendProgressPage = lazy(() => import('@/pages/SendProgressPage'));
const TemplatesPage = lazy(() => import('@/pages/TemplatesPage'));
const HistoryPage = lazy(() => import('@/pages/HistoryPage'));
const JobDetailPage = lazy(() => import('@/pages/JobDetailPage'));
const UsersPage = lazy(() => import('@/pages/admin/UsersPage'));
const SendersPage = lazy(() => import('@/pages/admin/SendersPage'));
const AuditPage = lazy(() => import('@/pages/admin/AuditPage'));
const SuppressionsPage = lazy(() => import('@/pages/admin/SuppressionsPage'));

export function FullPageSpinner() {
  return (
    <div className="grid h-full place-items-center">
      <Loader2 className="size-6 animate-spin text-muted-foreground" />
    </div>
  );
}

function RequireAuth() {
  const { user, loading } = useAuth();
  const loc = useLocation();
  if (loading) return <FullPageSpinner />;
  if (!user) return <Navigate to="/login" replace state={{ from: loc.pathname + loc.search }} />;
  if (user.mustChangePassword && loc.pathname !== '/change-password') return <Navigate to="/change-password" replace />;
  return <Outlet />;
}

function RequireAdmin() {
  const { isAdmin } = useAuth();
  return isAdmin ? <Outlet /> : <Navigate to="/compose" replace />;
}

export default function App() {
  return (
    <Suspense fallback={<FullPageSpinner />}>
      <Routes>
        <Route path="/login" element={<LoginPage />} />
        <Route element={<RequireAuth />}>
          <Route element={<AppShell />}>
            <Route index element={<Navigate to="/compose" replace />} />
            <Route path="/change-password" element={<ChangePasswordPage />} />
            <Route path="/compose" element={<ComposePage />} />
            <Route path="/send/:jobId" element={<SendProgressPage />} />
            <Route path="/templates" element={<TemplatesPage />} />
            <Route path="/history" element={<HistoryPage />} />
            <Route path="/history/:jobId" element={<JobDetailPage />} />
            <Route element={<RequireAdmin />}>
              <Route path="/admin/users" element={<UsersPage />} />
              <Route path="/admin/senders" element={<SendersPage />} />
              <Route path="/admin/audit" element={<AuditPage />} />
              <Route path="/admin/suppressions" element={<SuppressionsPage />} />
            </Route>
          </Route>
        </Route>
        <Route path="*" element={<Navigate to="/compose" replace />} />
      </Routes>
    </Suspense>
  );
}
