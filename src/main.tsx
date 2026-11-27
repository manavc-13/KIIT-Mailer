import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { BrowserRouter } from 'react-router-dom';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { Toaster } from 'sonner';
import { TooltipProvider } from '@/components/ui/tooltip';
import { AuthProvider } from '@/lib/auth';
import '@/lib/theme';
import App from './App';
import './index.css';

// After a new deployment, a tab still running the previous version asks for JS chunk files
// that no longer exist. Reload once to pick up the new version instead of showing a broken page.
// The sessionStorage flag prevents a reload loop if the chunk is genuinely unavailable.
const RELOAD_FLAG = 'km_chunk_reload_at';
window.addEventListener('vite:preloadError', (event) => {
  let last = 0;
  try { last = Number(sessionStorage.getItem(RELOAD_FLAG) || 0); } catch { /* storage unavailable */ }
  if (Date.now() - last < 60_000) return; // already reloaded recently: let the error surface
  event.preventDefault();
  try { sessionStorage.setItem(RELOAD_FLAG, String(Date.now())); } catch { /* ignore */ }
  window.location.reload();
});

const queryClient = new QueryClient({
  defaultOptions: {
    queries: { refetchOnWindowFocus: false, retry: 1 },
  },
});

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <QueryClientProvider client={queryClient}>
      <BrowserRouter>
        <AuthProvider>
          <TooltipProvider delayDuration={300}>
            <App />
            <Toaster position="bottom-right" richColors closeButton />
          </TooltipProvider>
        </AuthProvider>
      </BrowserRouter>
    </QueryClientProvider>
  </StrictMode>,
);
