import { describe, expect, it } from 'vitest';
import {
  DARK,
  LIGHT,
  bracketColor,
  nextPreference,
  paletteFor,
  paletteVariables,
  resolveTheme,
} from '../src/renderer/model/theme.js';

describe('resolveTheme', () => {
  it('follows the system when set to system', () => {
    expect(resolveTheme('system', true)).toBe('dark');
    expect(resolveTheme('system', false)).toBe('light');
  });

  it('overrides the system when set explicitly', () => {
    expect(resolveTheme('light', true)).toBe('light');
    expect(resolveTheme('dark', false)).toBe('dark');
  });
});

describe('palettes', () => {
  it('returns the named palette', () => {
    expect(paletteFor('dark')).toBe(DARK);
    expect(paletteFor('light')).toBe(LIGHT);
  });

  it('defines the same keys in both themes, so neither can miss a colour', () => {
    expect(Object.keys(LIGHT).sort()).toEqual(Object.keys(DARK).sort());
  });

  it('gives the two themes different backgrounds', () => {
    expect(LIGHT.gridBg).not.toBe(DARK.gridBg);
  });

  it('supplies bracket colours for nesting', () => {
    expect(LIGHT.bracketDepths.length).toBeGreaterThan(2);
  });
});

describe('paletteVariables', () => {
  it('renders custom properties in kebab case', () => {
    const vars = paletteVariables(LIGHT);
    expect(vars['--mz-grid-line']).toBe(LIGHT.gridLine);
    expect(vars['--mz-text-muted']).toBe(LIGHT.textMuted);
  });

  it('expands the bracket array into indexed properties', () => {
    const vars = paletteVariables(LIGHT);
    expect(vars['--mz-bracket-0']).toBe(LIGHT.bracketDepths[0]);
    expect(vars['--mz-bracket-3']).toBe(LIGHT.bracketDepths[3]);
  });

  it('does not emit the palette name as a colour', () => {
    expect(paletteVariables(LIGHT)['--mz-name']).toBeUndefined();
  });
});

describe('bracketColor', () => {
  it('cycles past the end of the palette', () => {
    expect(bracketColor(LIGHT, 4)).toBe(LIGHT.bracketDepths[0]);
  });

  it('handles a negative depth without throwing', () => {
    expect(bracketColor(LIGHT, -1)).toBe(LIGHT.bracketDepths[3]);
  });
});

describe('nextPreference', () => {
  it('cycles system, light, dark', () => {
    expect(nextPreference('system')).toBe('light');
    expect(nextPreference('light')).toBe('dark');
    expect(nextPreference('dark')).toBe('system');
  });
});
