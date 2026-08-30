import { describe, expect, it } from 'vitest';
import {
  FORMAT_PRESETS,
  adjustDecimals,
  presetForCode,
  presetSample,
  presetsByCategory,
  validateFormatCode,
} from '../src/renderer/model/number-formats.js';

describe('format presets', () => {
  it('starts with General', () => {
    expect(FORMAT_PRESETS[0]?.id).toBe('general');
    expect(FORMAT_PRESETS[0]?.code).toBeUndefined();
  });

  it('gives every preset a unique id', () => {
    const ids = FORMAT_PRESETS.map((p) => p.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('renders a sample through the real format engine', () => {
    const thousands = FORMAT_PRESETS.find((p) => p.id === 'thousands')!;
    expect(presetSample(thousands)).toBe('1,234.57');
  });

  it('renders a percent sample as a percentage', () => {
    const percent = FORMAT_PRESETS.find((p) => p.id === 'percent')!;
    expect(presetSample(percent)).toBe('12.34%');
  });

  it('renders a date sample as a date, not a serial', () => {
    const short = FORMAT_PRESETS.find((p) => p.id === 'date-short')!;
    expect(presetSample(short)).toMatch(/^\d{4}-\d{2}-\d{2}$/);
  });

  it('groups presets by category in menu order', () => {
    const groups = presetsByCategory();
    expect(groups.map((g) => g.category)).toEqual([
      'general',
      'number',
      'currency',
      'date',
      'time',
      'text',
    ]);
  });
});

describe('presetForCode', () => {
  it('maps an undefined code to General', () => {
    expect(presetForCode(undefined)?.id).toBe('general');
  });

  it('maps the literal word General to General', () => {
    expect(presetForCode('general')?.id).toBe('general');
  });

  it('finds a preset by its code', () => {
    expect(presetForCode('0.00')?.id).toBe('decimal2');
  });

  it('returns undefined for a code no preset uses', () => {
    expect(presetForCode('0.000###')).toBeUndefined();
  });
});

describe('validateFormatCode', () => {
  it('rejects an empty code rather than silently meaning General', () => {
    const result = validateFormatCode('   ');
    expect(result.valid).toBe(false);
    expect(result.problem).toMatch(/General/);
  });

  it('accepts a plain numeric code and previews it', () => {
    const result = validateFormatCode('#,##0.0');
    expect(result.valid).toBe(true);
    expect(result.preview).toBe('1,234.6');
  });

  it('previews a date code against a date, not a number', () => {
    const result = validateFormatCode('yyyy-mm-dd');
    expect(result.isDate).toBe(true);
    expect(result.preview).toMatch(/^\d{4}-/);
  });

  it('accepts a four-section code', () => {
    expect(validateFormatCode('0;-0;"zero";@').valid).toBe(true);
  });

  it('rejects a five-section code with a readable reason', () => {
    const result = validateFormatCode('0;0;0;0;0');
    expect(result.valid).toBe(false);
    expect(result.problem).toMatch(/four sections/);
  });

  it('does not count a semicolon inside quotes as a section break', () => {
    expect(validateFormatCode('"a;b"0').valid).toBe(true);
  });

  it('does not count a semicolon inside brackets as a section break', () => {
    expect(validateFormatCode('[>0]0;[<0]-0').valid).toBe(true);
  });
});

describe('adjustDecimals', () => {
  it('adds a decimal place to a plain integer format', () => {
    expect(adjustDecimals('0', 1)).toBe('0.0');
  });

  it('adds a decimal place to General by treating it as 0', () => {
    expect(adjustDecimals(undefined, 1)).toBe('0.0');
  });

  it('extends an existing decimal run', () => {
    expect(adjustDecimals('0.00', 1)).toBe('0.000');
  });

  it('removes a decimal place', () => {
    expect(adjustDecimals('0.00', -1)).toBe('0.0');
  });

  it('removes the decimal point entirely at zero places', () => {
    expect(adjustDecimals('0.0', -1)).toBe('0');
  });

  it('will not go below zero decimal places', () => {
    expect(adjustDecimals('0', -1)).toBe('0');
  });

  it('keeps a thousands separator when adding places', () => {
    expect(adjustDecimals('#,##0', 1)).toBe('#,##0.0');
  });

  it('keeps a currency prefix rather than replacing the code', () => {
    expect(adjustDecimals('"$"#,##0.00', 1)).toBe('"$"#,##0.000');
  });

  it('keeps a suffix after the decimal run', () => {
    expect(adjustDecimals('0.00%', 1)).toBe('0.000%');
  });

  it('caps the number of places', () => {
    expect(adjustDecimals('0.00', 99)).toBe(`0.${'0'.repeat(30)}`);
  });
});
