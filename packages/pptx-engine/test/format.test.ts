import { describe, expect, it } from 'vitest';
import { formatScalar, renderField, validateFormat } from '../src/format.ts';

const ctx = { locale: 'en-US', currency: 'USD' };

describe('formatScalar', () => {
  it.each([
    [1234.567, undefined, '1,234.57'],
    [1234, undefined, '1,234'],
    ['00123', undefined, '00123'],
    ['Plant A', undefined, 'Plant A'],
    [0.853, 'percent', '85.3%'],
    [0.85, 'percent:1', '85.0%'],
    [0.8567, 'percent:0', '86%'],
    [1234.5, 'number:0', '1,235'],
    [1234.5, 'number:2', '1,234.50'],
    [1234.5, 'currency', '$1,234.50'],
    [1234.5, 'currency:0', '$1,235'],
    ['0.5', 'percent', '50%'],
    [42, 'text', '42'],
    [new Date(Date.UTC(2026, 10, 1)), undefined, 'Nov 1, 2026'],
    ['2026-11-01', 'date', 'Nov 1, 2026'],
    [null, 'number', ''],
    [true, undefined, 'true'],
  ] as const)('%s as %s → %s', (v, f, expected) => {
    expect(formatScalar(v, f, ctx)).toBe(expected);
  });

  it('respects the locale', () => {
    expect(formatScalar(1234.5, 'number:1', { locale: 'de-DE', currency: 'EUR' })).toBe('1.234,5');
    expect(formatScalar(1234.5, 'currency', { locale: 'de-DE', currency: 'EUR' })).toBe('1.234,50 €');
  });
});

describe('validateFormat', () => {
  it('accepts known formats and empty', () => {
    for (const f of [undefined, '', 'percent', 'Percent:1', 'number:3', 'currency', 'date', 'text']) {
      expect(validateFormat(f)).toBeUndefined();
    }
  });
  it('explains unknown formats', () => {
    expect(validateFormat('0.0%')).toMatch(/not a format DeckForge knows/);
    expect(validateFormat('date:2')).toMatch(/decimals only apply/);
  });
});

describe('renderField', () => {
  const entry = { value: 0.78, previous: 0.71, format: 'percent:1' };
  it('renders previous, change, change_pct and trend', () => {
    expect(renderField(entry, 'value', ctx)).toEqual({ text: '78.0%' });
    expect(renderField(entry, 'previous', ctx)).toEqual({ text: '71.0%' });
    expect(renderField(entry, 'change', ctx)).toEqual({ text: '+7.0%' });
    expect(renderField(entry, 'change_pct', ctx)).toEqual({ text: '+9.9%' });
    expect(renderField(entry, 'trend', ctx)).toEqual({ text: '▲' });
    expect(renderField({ value: 5, previous: 5 }, 'trend', ctx)).toEqual({ text: '▶' });
    expect(renderField({ value: 4, previous: 5 }, 'change', ctx)).toEqual({ text: '-1' });
  });
  it('explains when trend cannot be computed', () => {
    expect(renderField({ value: 'n/a', previous: 5 }, 'trend', ctx)).toHaveProperty('error');
  });
});
