import type { Cell } from 'exceljs';
import type { Scalar } from '@deckforge/pptx-engine';

export type CellResult = { ok: true; value: Scalar } | { ok: false; reason: string };

function convert(v: unknown): CellResult {
  if (v === null || v === undefined) return { ok: true, value: null };
  if (typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return { ok: true, value: v };
  if (v instanceof Date) return { ok: true, value: v };
  if (typeof v === 'object') {
    const o = v as Record<string, unknown>;
    if ('formula' in o || 'sharedFormula' in o) {
      if (o.result === undefined) {
        return { ok: false, reason: 'has a formula with no saved result. Open the workbook in Excel, save it, and try again' };
      }
      return convert(o.result);
    }
    if ('error' in o) return { ok: false, reason: `shows the error ${String(o.error)}. Fix the cell in Excel` };
    if (Array.isArray(o.richText)) {
      return { ok: true, value: (o.richText as { text: string }[]).map((t) => t.text).join('') };
    }
    if ('text' in o) return convert(o.text);
  }
  return { ok: true, value: String(v) };
}

export function cellValue(cell: Cell): CellResult {
  return convert(cell.value);
}

/** Plain text of a cell, for headers and keys. */
export function cellText(cell: Cell): string {
  const r = cellValue(cell);
  if (!r.ok || r.value === null) return '';
  return r.value instanceof Date ? r.value.toISOString().slice(0, 10) : String(r.value);
}

export function isBlank(v: Scalar): boolean {
  return v === null || (typeof v === 'string' && v.trim() === '');
}

/**
 * Maps an Excel number format to a DeckForge format so cells look the way they do in Excel.
 * Returns undefined for General/text formats or anything unrecognised.
 */
export function formatFromNumFmt(numFmt: string | undefined): string | undefined {
  if (!numFmt) return undefined;
  const positive = numFmt.split(';')[0] ?? '';
  const isCurrency = /[$€£¥]|\[\$/.test(positive);
  // Drop quoted text, escaped characters and [color]/[$locale] blocks before looking at placeholders.
  const f = positive.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, '');
  if (f.trim() === '' || /^general$/i.test(f.trim()) || f.trim() === '@') return undefined;
  const decimals = /\.([0#?]+)/.exec(f)?.[1]?.length ?? 0;
  if (f.includes('%')) return `percent:${decimals}`;
  if (isCurrency) return `currency:${decimals}`;
  if (/[dmyhs]/i.test(f)) return 'date';
  if (/[0#]/.test(f)) return `number:${decimals}`;
  return undefined;
}
