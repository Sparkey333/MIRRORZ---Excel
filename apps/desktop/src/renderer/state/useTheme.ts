/**
 * Applying the theme to the document.
 *
 * The palette is written as CSS custom properties on the root element rather
 * than swapping a stylesheet, because the canvas grid needs the same values as
 * strings and a stylesheet swap would leave it painting last week's colours.
 */

import { useEffect, useMemo, useState } from 'react';
import { paletteFor, paletteVariables, resolveTheme, type Palette, type ThemePreference } from '../model/theme.js';

/** Track the OS preference, updating when the user changes it mid-session. */
export function useSystemPrefersDark(): boolean {
  const [prefersDark, setPrefersDark] = useState(() => {
    if (typeof window === 'undefined' || !window.matchMedia) return false;
    return window.matchMedia('(prefers-color-scheme: dark)').matches;
  });

  useEffect(() => {
    if (typeof window === 'undefined' || !window.matchMedia) return;
    const query = window.matchMedia('(prefers-color-scheme: dark)');
    const onChange = (e: MediaQueryListEvent): void => setPrefersDark(e.matches);
    // addEventListener is not in older WebKit's MediaQueryList, and Electron
    // pins its Chromium, so the modern path is the only one needed.
    query.addEventListener('change', onChange);
    return () => query.removeEventListener('change', onChange);
  }, []);

  return prefersDark;
}

export function useTheme(preference: ThemePreference): Palette {
  const prefersDark = useSystemPrefersDark();
  const palette = useMemo(
    () => paletteFor(resolveTheme(preference, prefersDark)),
    [preference, prefersDark],
  );

  useEffect(() => {
    if (typeof document === 'undefined') return;
    const root = document.documentElement;
    for (const [key, value] of Object.entries(paletteVariables(palette))) {
      root.style.setProperty(key, value);
    }
    root.dataset['mzTheme'] = palette.name;
    root.style.colorScheme = palette.name;
  }, [palette]);

  return palette;
}
