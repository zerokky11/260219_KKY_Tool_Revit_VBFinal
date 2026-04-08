const THEME_KEY = 'kky_theme';

function getPreferredTheme() {
  const storedTheme = localStorage.getItem(THEME_KEY);
  if (storedTheme === 'dark' || storedTheme === 'light') return storedTheme;

  const bootTheme = document.documentElement.dataset.theme;
  if (bootTheme === 'dark' || bootTheme === 'light') return bootTheme;

  return 'dark';
}

export function applyTheme(t) { document.documentElement.dataset.theme = t; localStorage.setItem(THEME_KEY, t); }
export function toggleTheme() { const cur = getPreferredTheme(); applyTheme(cur === 'dark' ? 'light' : 'dark'); }
export function initTheme() { applyTheme(getPreferredTheme()); }
