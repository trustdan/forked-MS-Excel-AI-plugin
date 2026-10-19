import { createRoot } from 'react-dom/client';
import { App } from './App.js';
import './styles.css';
// Shared runtime: the task pane page hosts the custom functions too.
import '../functions/functions.js';

import { applyTheme, getStoredThemePreference, resolveIsDark } from './theme.js';

// Apply stored/detected theme immediately before render
applyTheme(resolveIsDark(getStoredThemePreference()));

let rootMounted = false;

function mountApp(): void {
  if (rootMounted) return;
  const container = document.getElementById('root');
  if (!container) return;
  rootMounted = true;
  try {
    applyTheme(resolveIsDark(getStoredThemePreference()));
    createRoot(container).render(<App />);
  } catch (err) {
    container.innerHTML = `
      <div style="padding: 20px; font-family: sans-serif; color: #dc2626;">
        <h3>Error Rendering Taskpane</h3>
        <pre style="white-space: pre-wrap; font-size: 12px;">${String(err)}</pre>
      </div>
    `;
  }
}

// 1. If Office.js is loaded, wait for Office.onReady
if (typeof Office !== 'undefined' && typeof Office.onReady === 'function') {
  Office.onReady()
    .then(() => {
      mountApp();
    })
    .catch((err) => {
      console.warn('Office.onReady failed:', err);
      mountApp();
    });
}

// 2. Fallback timer so the screen is never permanently blank if Office.onReady delays or hangs
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    setTimeout(mountApp, 250);
  });
} else {
  setTimeout(mountApp, 250);
}

// 3. Catch any global window errors to prevent silent blank screens
window.addEventListener('error', (event) => {
  console.error('[taskpane uncaught error]', event.error || event.message);
  const container = document.getElementById('root');
  if (container && !rootMounted) {
    container.innerHTML = `
      <div style="padding: 20px; font-family: sans-serif; color: #dc2626;">
        <h3>Initialization Error</h3>
        <p>${event.message}</p>
        <pre style="white-space: pre-wrap; font-size: 11px;">${event.filename}:${event.lineno}</pre>
      </div>
    `;
  }
});
