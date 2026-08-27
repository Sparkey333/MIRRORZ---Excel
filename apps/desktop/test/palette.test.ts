import { describe, expect, it } from 'vitest';
import { moveIndex, pushRecent, searchPalette, type PaletteItem } from '../src/renderer/model/palette.js';

const items: PaletteItem[] = [
  { id: 'edit.undo', kind: 'command', title: 'Undo', detail: 'Edit', keywords: 'revert back' },
  { id: 'view.history', kind: 'command', title: 'Show history', detail: 'View' },
  { id: 'sheet:Summary', kind: 'sheet', title: 'Summary', detail: 'Sheet 1' },
  { id: 'sheet:Sales', kind: 'sheet', title: 'Sales', detail: 'Sheet 2' },
  { id: 'name:TaxRate', kind: 'name', title: 'TaxRate', detail: 'Assumptions!$B$2' },
];

describe('searchPalette', () => {
  it('returns everything for an empty query', () => {
    expect(searchPalette(items, '')).toHaveLength(items.length);
  });

  it('finds a sheet by name', () => {
    expect(searchPalette(items, 'summ')[0]?.item.id).toBe('sheet:Summary');
  });

  it('finds a command by name', () => {
    expect(searchPalette(items, 'undo')[0]?.item.id).toBe('edit.undo');
  });

  it('finds a command by keyword when the title does not match', () => {
    expect(searchPalette(items, 'revert').map((r) => r.item.id)).toContain('edit.undo');
  });

  it('finds a named range by its target', () => {
    expect(searchPalette(items, 'Assumptions').map((r) => r.item.id)).toContain('name:TaxRate');
  });

  it('restricts to commands after a > prefix', () => {
    const results = searchPalette(items, '>s');
    expect(results.every((r) => r.item.kind === 'command')).toBe(true);
  });

  it('restricts to sheets after an @ prefix', () => {
    const results = searchPalette(items, '@s');
    expect(results.map((r) => r.item.id)).toEqual(['sheet:Sales', 'sheet:Summary']);
  });

  it('restricts by an explicit kind option', () => {
    const results = searchPalette(items, '', { kind: 'name' });
    expect(results).toHaveLength(1);
  });

  it('ranks sheets above commands on an equal text match', () => {
    const tied: PaletteItem[] = [
      { id: 'c', kind: 'command', title: 'Balance' },
      { id: 's', kind: 'sheet', title: 'Balance' },
    ];
    expect(searchPalette(tied, 'balance')[0]?.item.id).toBe('s');
  });

  it('floats a recently used item on an equal match', () => {
    const tied: PaletteItem[] = [
      { id: 'a', kind: 'command', title: 'Format' },
      { id: 'b', kind: 'command', title: 'Format' },
    ];
    expect(searchPalette(tied, 'format', { recent: ['b'] })[0]?.item.id).toBe('b');
  });

  it('sinks a disabled command below an enabled one', () => {
    const tied: PaletteItem[] = [
      { id: 'a', kind: 'command', title: 'Redo', enabled: false },
      { id: 'b', kind: 'command', title: 'Redo', enabled: true },
    ];
    expect(searchPalette(tied, 'redo')[0]?.item.id).toBe('b');
  });

  it('honours the limit', () => {
    expect(searchPalette(items, '', { limit: 2 })).toHaveLength(2);
  });

  it('returns nothing when nothing matches', () => {
    expect(searchPalette(items, 'zzzz')).toEqual([]);
  });
});

describe('moveIndex', () => {
  it('moves forwards', () => {
    expect(moveIndex(0, 1, 3)).toBe(1);
  });

  it('wraps past the end', () => {
    expect(moveIndex(2, 1, 3)).toBe(0);
  });

  it('wraps before the start', () => {
    expect(moveIndex(0, -1, 3)).toBe(2);
  });

  it('stays at zero for an empty list', () => {
    expect(moveIndex(0, 1, 0)).toBe(0);
  });
});

describe('pushRecent', () => {
  it('puts the newest first', () => {
    expect(pushRecent(['a', 'b'], 'c')).toEqual(['c', 'a', 'b']);
  });

  it('de-duplicates', () => {
    expect(pushRecent(['a', 'b'], 'b')).toEqual(['b', 'a']);
  });

  it('caps the list', () => {
    expect(pushRecent(['a', 'b', 'c'], 'd', 2)).toEqual(['d', 'a']);
  });
});
