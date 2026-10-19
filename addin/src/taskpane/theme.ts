export type ThemePreference = 'auto' | 'light' | 'dark';

const THEME_STORAGE_KEY = 'excelai-theme-preference';

/**
 * Calculates perceived luminance of a hex color (#RRGGBB).
 * Returns true if the background is dark (luminance < 128).
 */
export function isColorDark(hex: string): boolean {
  const clean = hex.replace('#', '').trim();
  if (clean.length === 3) {
    const r = parseInt(clean[0]! + clean[0]!, 16);
    const g = parseInt(clean[1]! + clean[1]!, 16);
    const b = parseInt(clean[2]! + clean[2]!, 16);
    return r * 0.299 + g * 0.587 + b * 0.114 < 128;
  }
  if (clean.length === 6) {
    const r = parseInt(clean.slice(0, 2), 16);
    const g = parseInt(clean.slice(2, 4), 16);
    const b = parseInt(clean.slice(4, 6), 16);
    return r * 0.299 + g * 0.587 + b * 0.114 < 128;
  }
  return false;
}

/**
 * Determines whether dark mode should be active based on user preference,
 * Office.context.officeTheme, or system prefers-color-scheme.
 */
export function resolveIsDark(preference: ThemePreference): boolean {
  if (preference === 'dark') return true;
  if (preference === 'light') return false;

  // 'auto' mode: check Office host theme first
  if (typeof Office !== 'undefined' && Office.context?.officeTheme?.bodyBackgroundColor) {
    return isColorDark(Office.context.officeTheme.bodyBackgroundColor);
  }

  // Fallback to system OS prefers-color-scheme
  if (typeof window !== 'undefined' && window.matchMedia) {
    return window.matchMedia('(prefers-color-scheme: dark)').matches;
  }

  return false;
}

export function getStoredThemePreference(): ThemePreference {
  if (typeof localStorage === 'undefined') return 'auto';
  const val = localStorage.getItem(THEME_STORAGE_KEY);
  if (val === 'light' || val === 'dark' || val === 'auto') return val;
  return 'auto';
}

export function setStoredThemePreference(pref: ThemePreference): void {
  if (typeof localStorage === 'undefined') return;
  localStorage.setItem(THEME_STORAGE_KEY, pref);
}

/**
 * Applies the dark or light theme attribute to <html> and <body>.
 */
export function applyTheme(isDark: boolean): void {
  if (typeof document === 'undefined') return;
  const root = document.documentElement;
  if (isDark) {
    root.setAttribute('data-theme', 'dark');
    root.classList.add('dark');
  } else {
    root.setAttribute('data-theme', 'light');
    root.classList.remove('dark');
  }
}
