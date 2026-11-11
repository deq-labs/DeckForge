import type { Scalar, TextField, ValueEntry } from './types.ts';

export interface FormatContext {
  locale: string;
  currency: string;
}

interface ParsedFormat {
  kind: 'auto' | 'text' | 'number' | 'percent' | 'currency' | 'date';
  decimals: number | undefined;
}

const FORMAT_RE = /^(text|number|percent|currency|date)(?::(\d{1,2}))?$/;

/** Returns an error message for an invalid format string, or undefined if it's valid. */
export function validateFormat(format: string | undefined): string | undefined {
  if (format === undefined || format.trim() === '') return undefined;
  const m = FORMAT_RE.exec(format.trim().toLowerCase());
  if (!m) {
    return `"${format}" is not a format DeckForge knows. Use one of: text, number, percent, currency, date (optionally with decimals, e.g. percent:1), or leave it empty.`;
  }
  if (m[2] !== undefined && (m[1] === 'text' || m[1] === 'date')) {
    return `"${format}": decimals only apply to number, percent and currency.`;
  }
  return undefined;
}

function parseFormat(format: string | undefined): ParsedFormat {
  if (format === undefined || format.trim() === '') return { kind: 'auto', decimals: undefined };
  const m = FORMAT_RE.exec(format.trim().toLowerCase());
  if (!m) throw new Error(validateFormat(format));
  return { kind: m[1] as ParsedFormat['kind'], decimals: m[2] === undefined ? undefined : Number(m[2]) };
}

function toNumber(v: Scalar): number | undefined {
  if (typeof v === 'number') return Number.isFinite(v) ? v : undefined;
  if (typeof v === 'string' && v.trim() !== '') {
    const n = Number(v.trim());
    return Number.isFinite(n) ? n : undefined;
  }
  return undefined;
}

function fixed(decimals: number | undefined, defaultMax: number): Intl.NumberFormatOptions {
  return decimals === undefined
    ? { minimumFractionDigits: 0, maximumFractionDigits: defaultMax }
    : { minimumFractionDigits: decimals, maximumFractionDigits: decimals };
}

export function formatScalar(v: Scalar | undefined, format: string | undefined, ctx: FormatContext, signed = false): string {
  if (v === null || v === undefined) return '';
  const f = parseFormat(format);
  const signDisplay: Intl.NumberFormatOptions['signDisplay'] = signed ? 'exceptZero' : 'auto';

  if (f.kind === 'text') return v instanceof Date ? formatDate(v, ctx) : String(v);
  if (v instanceof Date) return formatDate(v, ctx);
  if (f.kind === 'date') {
    const d = typeof v === 'string' ? new Date(v) : undefined;
    return d && !Number.isNaN(d.getTime()) ? formatDate(d, ctx) : String(v);
  }

  const n = toNumber(v);
  if (n === undefined) return String(v);

  switch (f.kind) {
    case 'percent':
      return new Intl.NumberFormat(ctx.locale, { style: 'percent', signDisplay, ...fixed(f.decimals, 1) }).format(n);
    case 'currency': {
      const opts: Intl.NumberFormatOptions = { style: 'currency', currency: ctx.currency, signDisplay };
      if (f.decimals !== undefined) Object.assign(opts, fixed(f.decimals, f.decimals));
      return new Intl.NumberFormat(ctx.locale, opts).format(n);
    }
    case 'number':
      return new Intl.NumberFormat(ctx.locale, { signDisplay, ...fixed(f.decimals, 2) }).format(n);
    case 'auto':
      // Strings that look numeric are left exactly as written; only real numbers get grouping.
      if (typeof v === 'string') return v;
      return new Intl.NumberFormat(ctx.locale, { signDisplay, ...fixed(undefined, 2) }).format(n);
  }
}

function formatDate(d: Date, ctx: FormatContext): string {
  // Spreadsheet dates have no time zone; ExcelJS returns them as UTC midnight.
  return new Intl.DateTimeFormat(ctx.locale, { dateStyle: 'medium', timeZone: 'UTC' }).format(d);
}

export function normalizeEntry(v: ValueEntry | Scalar | undefined): ValueEntry | undefined {
  if (v === undefined) return undefined;
  if (v !== null && typeof v === 'object' && !(v instanceof Date)) return v;
  return { value: v };
}

/**
 * Renders one field of a value entry. Returns `{ error }` when the field can't be computed,
 * e.g. `.trend` without a numeric previous value.
 */
export function renderField(entry: ValueEntry, field: TextField, ctx: FormatContext): { text: string } | { error: string } {
  const { value, previous, format } = entry;
  switch (field) {
    case 'value':
      return { text: formatScalar(value, format, ctx) };
    case 'previous':
      return { text: formatScalar(previous ?? null, format, ctx) };
    case 'change':
    case 'change_pct':
    case 'trend': {
      const cur = toNumber(value);
      const prev = toNumber(previous ?? null);
      if (cur === undefined || prev === undefined) {
        return { error: `needs numbers in both "value" and "previous" to show .${field}` };
      }
      if (field === 'trend') return { text: cur > prev ? '▲' : cur < prev ? '▼' : '▶' };
      if (field === 'change') return { text: formatScalar(cur - prev, format, ctx, true) };
      if (prev === 0) return { text: '' };
      return { text: formatScalar((cur - prev) / Math.abs(prev), 'percent:1', ctx, true) };
    }
  }
}
