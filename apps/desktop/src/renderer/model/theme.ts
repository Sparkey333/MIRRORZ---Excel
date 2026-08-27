/**
 * Light and dark palettes.
 *
 * Defined as data rather than as two stylesheets because the grid is drawn on a
 * canvas: the renderer needs the same colours as literal strings, and keeping
 * one source for both is the only way the grid lines and the chrome around them
 * stay the same grey. The CSS side consumes them as custom properties.
 *
 * Following the OS by default is the right default and a toggle is still
 * required, because "dark everywhere except the thing I stare at all day" is a
 * real preference and an application that refuses to hear it is annoying.
 */

export type ThemePreference = 'system' | 'light' | 'dark';
export type ThemeName = 'light' | 'dark';

export interface Palette {
  name: ThemeName;
  /** Application chrome. */
  bg: string;
  bgRaised: string;
  bgSunken: string;
  border: string;
  borderStrong: string;
  text: string;
  textMuted: string;
  accent: string;
  accentText: string;
  danger: string;
  warning: string;
  success: string;
  /** Grid surface, shared with the canvas renderer. */
  gridBg: string;
  gridLine: string;
  gridHeaderBg: string;
  gridHeaderText: string;
  selectionFill: string;
  selectionBorder: string;
  /** Formula syntax highlighting. */
  synFunction: string;
  synRef: string;
  synString: string;
  synNumber: string;
  synOperator: string;
  synError: string;
  synName: string;
  /** Bracket depth colours, cycled. */
  bracketDepths: string[];
}

export const LIGHT: Palette = {
  name: 'light',
  bg: '#f6f7f9',
  bgRaised: '#ffffff',
  bgSunken: '#eceef2',
  border: '#d8dce3',
  borderStrong: '#b4bbc6',
  text: '#1b1f27',
  textMuted: '#5d6672',
  accent: '#1f6feb',
  accentText: '#ffffff',
  danger: '#c0392b',
  warning: '#a86a00',
  success: '#1a7f43',
  gridBg: '#ffffff',
  gridLine: '#e1e4ea',
  gridHeaderBg: '#f0f2f5',
  gridHeaderText: '#414956',
  selectionFill: 'rgba(31, 111, 235, 0.10)',
  selectionBorder: '#1f6feb',
  synFunction: '#7c3aed',
  synRef: '#0b7285',
  synString: '#1a7f43',
  synNumber: '#b35309',
  synOperator: '#5d6672',
  synError: '#c0392b',
  synName: '#1b1f27',
  bracketDepths: ['#1f6feb', '#b35309', '#1a7f43', '#7c3aed'],
};

export const DARK: Palette = {
  name: 'dark',
  bg: '#16181d',
  bgRaised: '#1e2127',
  bgSunken: '#101216',
  border: '#2c3038',
  borderStrong: '#434a56',
  text: '#e6e9ef',
  textMuted: '#9aa3b2',
  accent: '#4c8dff',
  accentText: '#0b0d11',
  danger: '#ff6b5e',
  warning: '#e0a33a',
  success: '#4fc07d',
  gridBg: '#1a1d22',
  gridLine: '#2b2f37',
  gridHeaderBg: '#22262d',
  gridHeaderText: '#b6becb',
  selectionFill: 'rgba(76, 141, 255, 0.16)',
  selectionBorder: '#4c8dff',
  synFunction: '#c39bff',
  synRef: '#5ec8d8',
  synString: '#7fd39a',
  synNumber: '#f2a65a',
  synOperator: '#9aa3b2',
  synError: '#ff6b5e',
  synName: '#e6e9ef',
  bracketDepths: ['#4c8dff', '#f2a65a', '#7fd39a', '#c39bff'],
};

export function resolveTheme(preference: ThemePreference, systemPrefersDark: boolean): ThemeName {
  if (preference === 'system') return systemPrefersDark ? 'dark' : 'light';
  return preference;
}

export function paletteFor(name: ThemeName): Palette {
  return name === 'dark' ? DARK : LIGHT;
}

/** Palette as CSS custom properties, applied to the document root. */
export function paletteVariables(palette: Palette): Record<string, string> {
  const vars: Record<string, string> = {};
  for (const [key, value] of Object.entries(palette)) {
    if (key === 'name') continue;
    if (Array.isArray(value)) {
      value.forEach((v, i) => {
        vars[`--mz-bracket-${i}`] = v;
      });
      continue;
    }
    vars[`--mz-${kebab(key)}`] = value as string;
  }
  return vars;
}

function kebab(s: string): string {
  return s.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`);
}

/** Cycle the bracket colours so nesting deeper than the palette still alternates. */
export function bracketColor(palette: Palette, depth: number): string {
  const colors = palette.bracketDepths;
  return colors[((depth % colors.length) + colors.length) % colors.length]!;
}

/** The next preference for the toolbar's three-state toggle. */
export function nextPreference(current: ThemePreference): ThemePreference {
  return current === 'system' ? 'light' : current === 'light' ? 'dark' : 'system';
}
