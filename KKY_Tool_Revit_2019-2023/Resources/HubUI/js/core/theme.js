import { post } from './bridge.js';

const THEME_KEY = 'kky_theme';

function getPreferredTheme() {
  const storedTheme = localStorage.getItem(THEME_KEY);
  if (storedTheme === 'dark' || storedTheme === 'light') return storedTheme;

  const bootTheme = document.documentElement.dataset.theme;
  if (bootTheme === 'dark' || bootTheme === 'light') return bootTheme;

  return 'dark';
}

function notifyTheme(theme) {
  try { post('ui:theme-changed', { theme }); } catch {}
}

export function applyTheme(t) {
  const theme = t === 'light' ? 'light' : 'dark';
  document.documentElement.dataset.theme = theme;
  localStorage.setItem(THEME_KEY, theme);
  notifyTheme(theme);
}
export function toggleTheme() { const cur = getPreferredTheme(); applyTheme(cur === 'dark' ? 'light' : 'dark'); }
export function initTheme() { applyTheme(getPreferredTheme()); }
