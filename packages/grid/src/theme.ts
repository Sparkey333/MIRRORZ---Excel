/**
 * Grid palettes.
 *
 * Every colour the painter uses comes from here; nothing is hard-coded in
 * render.ts. That is what makes a dark palette a data change rather than a
 * rewrite, and it is why the two palettes below can be swapped at runtime
 * without re-laying anything out.
 *
 * The dark palette is not the light one inverted. Excel's own dark mode gets
 * this wrong in places: inverting a grid line that reads as a hairline on white
 * produces a bright cage on black. Here the dark grid line is *lighter* than the
 * background by a small amount rather than the same contrast ratio flipped, and
 * cell fills authored for a white sheet are left alone, because silently
 * recolouring a user's yellow highlight would be a lie about their data.
 */

export interface GridTheme {
  readonly name: string;
  /** True when cell fills authored for a white sheet need a lighter default text. */
  readonly dark: boolean;

  readonly background: string;
  readonly gridLine: string;
  /** Divider between a frozen pane and the scrolling region. */
  readonly frozenLine: string;

  readonly headerBackground: string;
  readonly headerBackgroundSelected: string;
  readonly headerBackgroundActive: string;
  readonly headerText: string;
  readonly headerTextSelected: string;
  readonly headerLine: string;
  /** The heavier rule under the column headers and right of the row headers. */
  readonly headerEdge: string;
  readonly cornerBackground: string;

  readonly cellText: string;
  readonly cellTextMuted: string;
  readonly errorText: string;

  readonly selectionFill: string;
  readonly selectionBorder: string;
  readonly activeCellBorder: string;
  /** Fill handle at the bottom-right of the selection. */
  readonly fillHandle: string;

  readonly defaultFontFamily: string;
  readonly defaultFontSize: number;
}

export const LIGHT_THEME: GridTheme = Object.freeze({
  name: 'light',
  dark: false,
  background: '#ffffff',
  gridLine: '#d0d7de',
  frozenLine: '#8f9aa6',
  headerBackground: '#f2f3f5',
  headerBackgroundSelected: '#d8e4ee',
  headerBackgroundActive: '#a6c8e4',
  headerText: '#3b4149',
  headerTextSelected: '#12263a',
  headerLine: '#c3c9d0',
  headerEdge: '#9aa3ad',
  cornerBackground: '#eceef1',
  cellText: '#16191d',
  cellTextMuted: '#6b7280',
  errorText: '#b42318',
  selectionFill: 'rgba(33, 115, 190, 0.12)',
  selectionBorder: '#2173be',
  activeCellBorder: '#1b5e96',
  fillHandle: '#1b5e96',
  defaultFontFamily: 'Calibri, "Segoe UI", system-ui, sans-serif',
  defaultFontSize: 11,
});

export const DARK_THEME: GridTheme = Object.freeze({
  name: 'dark',
  dark: true,
  background: '#1b1d21',
  gridLine: '#33373d',
  frozenLine: '#6b7684',
  headerBackground: '#25282d',
  headerBackgroundSelected: '#31414f',
  headerBackgroundActive: '#3f5f7d',
  headerText: '#b8bec7',
  headerTextSelected: '#e6ecf3',
  headerLine: '#3a3f46',
  headerEdge: '#4a5058',
  cornerBackground: '#212429',
  cellText: '#e6e8ec',
  cellTextMuted: '#9098a4',
  errorText: '#f9776a',
  selectionFill: 'rgba(96, 165, 230, 0.16)',
  selectionBorder: '#60a5e6',
  activeCellBorder: '#8cc3f5',
  fillHandle: '#8cc3f5',
  defaultFontFamily: 'Calibri, "Segoe UI", system-ui, sans-serif',
  defaultFontSize: 11,
});

export const THEMES: Readonly<Record<string, GridTheme>> = Object.freeze({
  light: LIGHT_THEME,
  dark: DARK_THEME,
});

/** Build a variant without restating a whole palette. */
export function deriveTheme(base: GridTheme, patch: Partial<GridTheme>): GridTheme {
  return Object.freeze({ ...base, ...patch });
}

/**
 * The eight colour names the number-format mini-language can name, as concrete
 * hex. `[Red]-0.00` has to render red on both palettes, so the dark variants are
 * lifted enough to stay legible on a near-black background.
 */
const FORMAT_COLORS_LIGHT: Readonly<Record<string, string>> = Object.freeze({
  Black: '#000000',
  Blue: '#0000ff',
  Cyan: '#00ffff',
  Green: '#008000',
  Magenta: '#ff00ff',
  Red: '#ff0000',
  White: '#ffffff',
  Yellow: '#ffff00',
});

const FORMAT_COLORS_DARK: Readonly<Record<string, string>> = Object.freeze({
  Black: '#e6e8ec',
  Blue: '#7aa7ff',
  Cyan: '#67e8f9',
  Green: '#5ec98a',
  Magenta: '#f07ae0',
  Red: '#ff6b5e',
  White: '#ffffff',
  Yellow: '#ffe066',
});

export function formatColorToHex(name: string | undefined, theme: GridTheme): string | undefined {
  if (!name) return undefined;
  const table = theme.dark ? FORMAT_COLORS_DARK : FORMAT_COLORS_LIGHT;
  return table[name];
}
