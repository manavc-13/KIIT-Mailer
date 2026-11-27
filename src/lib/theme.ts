// Light/dark appearance for the app UI (not the emails). Persisted per browser.
import { useEffect, useState } from 'react';

type Appearance = 'light' | 'dark';
const KEY = 'km_appearance';

function initial(): Appearance {
  try {
    const v = localStorage.getItem(KEY);
    if (v === 'light' || v === 'dark') return v;
  } catch { /* storage unavailable */ }
  return 'light';
}

export function applyAppearance(a: Appearance) {
  document.documentElement.classList.toggle('dark', a === 'dark');
}

export function useTheme() {
  const [theme, setTheme] = useState<Appearance>(initial);
  useEffect(() => {
    applyAppearance(theme);
    try { localStorage.setItem(KEY, theme); } catch { /* ignore */ }
  }, [theme]);
  return { theme, toggle: () => setTheme(t => (t === 'dark' ? 'light' : 'dark')) };
}

applyAppearance(initial());
