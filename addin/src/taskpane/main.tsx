import { createRoot } from 'react-dom/client';
import { App } from './App.js';
import './styles.css';
// Shared runtime: the task pane page hosts the custom functions too.
import '../functions/functions.js';

import { applyTheme, getStoredThemePreference, resolveIsDark } from './theme.js';

// Apply stored/detected theme immediately before render
applyTheme(resolveIsDark(getStoredThemePreference()));

/* global Office */
Office.onReady(() => {
  applyTheme(resolveIsDark(getStoredThemePreference()));
  const container = document.getElementById('root');
  if (container) createRoot(container).render(<App />);
});

