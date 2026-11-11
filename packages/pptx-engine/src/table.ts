/** Reading and filling DrawingML tables (`a:tbl`). See design-log 0003, "Table fill rules". */
import { NS, child, children, create, is, path, removeNode, textOf } from './xml.ts';
import type { Element } from './xml.ts';
import { writeRunText } from './tags.ts';

export interface TableInfo {
  columns: string[];
  sampleRows: string[][];
  /** Why the table can't be filled, phrased to follow "Table 'x' on slide n ...". */
  problem?: string;
}

const rowsOf = (tbl: Element): Element[] => children(tbl, NS.a, 'tr');
const cellsOf = (tr: Element): Element[] => children(tr, NS.a, 'tc');

export function normalizeHeader(s: string): string {
  return s.replace(/\s+/g, ' ').trim().toLowerCase();
}

function paragraphText(p: Element): string {
  let s = '';
  for (const el of children(p)) {
    if (is(el, NS.a, 'r') || is(el, NS.a, 'fld')) s += textOf(child(el, NS.a, 't') ?? el);
    else if (is(el, NS.a, 'br')) s += '\n';
  }
  return s;
}

export function cellText(tc: Element): string {
  const tx = child(tc, NS.a, 'txBody');
  if (!tx) return '';
  return children(tx, NS.a, 'p').map(paragraphText).join('\n');
}

function isMerged(tc: Element): boolean {
  const truthy = (v: string | null): boolean => v === '1' || v === 'true';
  return (
    Number(tc.getAttribute('gridSpan') ?? '1') > 1 ||
    Number(tc.getAttribute('rowSpan') ?? '1') > 1 ||
    truthy(tc.getAttribute('hMerge')) ||
    truthy(tc.getAttribute('vMerge'))
  );
}

export function readTable(tbl: Element): TableInfo {
  const rows = rowsOf(tbl);
  const header = rows[0];
  if (!header) return { columns: [], sampleRows: [], problem: 'has no rows' };
  const columns = cellsOf(header).map((tc) => cellText(tc).replace(/\s+/g, ' ').trim());
  const body = rows.slice(1);
  const sampleRows = body.map((tr) => cellsOf(tr).map(cellText));
  const info: TableInfo = { columns, sampleRows };
  if (body.length === 0) {
    info.problem = 'needs at least one row under the header row, so DeckForge can copy its formatting. Add one row to the table in PowerPoint';
  } else if (body.some((tr) => cellsOf(tr).some(isMerged))) {
    info.problem = 'has merged cells below the header row, so DeckForge cannot add or remove rows. Unmerge those cells in PowerPoint';
  }
  return info;
}

/**
 * Fills the table's body rows. `columnMap[i]` is the data column for template column i,
 * or -1 to leave that column blank. Returns nothing; validation happens before this call.
 */
export function fillTable(frame: Element, tbl: Element, columnMap: number[], rows: string[][]): void {
  const body = rowsOf(tbl).slice(1);
  const styleRow = body[body.length - 1]!;
  let lastPlaced: Element | undefined;

  rows.forEach((row, i) => {
    let tr = body[i];
    if (!tr) {
      tr = styleRow.cloneNode(true) as Element;
      tbl.insertBefore(tr, lastPlaced!.nextSibling);
    }
    cellsOf(tr).forEach((tc, col) => {
      const src = columnMap[col] ?? -1;
      setCellText(tc, src < 0 ? '' : (row[src] ?? ''));
    });
    lastPlaced = tr;
  });
  for (const tr of body.slice(rows.length)) removeNode(tr);

  // Keep the frame's height in step with its rows so the selection box matches the table.
  const ext = path(frame, NS.p, 'xfrm');
  const extEl = ext && child(ext, NS.a, 'ext');
  if (extEl) {
    const total = rowsOf(tbl).reduce((sum, tr) => sum + Number(tr.getAttribute('h') ?? '0'), 0);
    if (total > 0) extEl.setAttribute('cy', String(total));
  }
}

/** Replaces a cell's text with `text`, keeping the formatting of its first run. */
export function setCellText(tc: Element, text: string): void {
  let tx = child(tc, NS.a, 'txBody');
  if (!tx) {
    tx = create(tc, NS.a, 'txBody');
    tx.appendChild(create(tc, NS.a, 'bodyPr'));
    tx.appendChild(create(tc, NS.a, 'lstStyle'));
    tc.insertBefore(tx, tc.firstChild);
  }
  const paras = children(tx, NS.a, 'p');
  let p = paras[0];
  if (!p) {
    p = create(tx, NS.a, 'p');
    tx.appendChild(p);
  }
  for (const extra of paras.slice(1)) removeNode(extra);

  const firstRun = child(p, NS.a, 'r');
  const runPr = firstRun ? child(firstRun, NS.a, 'rPr') : undefined;
  const endPr = child(p, NS.a, 'endParaRPr');
  for (const el of children(p)) {
    if (!is(el, NS.a, 'pPr') && !is(el, NS.a, 'endParaRPr')) removeNode(el);
  }
  if (text === '') return;

  const r = create(p, NS.a, 'r');
  if (runPr) {
    r.appendChild(runPr.cloneNode(true));
  } else if (endPr) {
    // endParaRPr has the same schema type as rPr; copy it so a blank template cell keeps its look.
    const rPr = create(p, NS.a, 'rPr');
    for (const attr of Array.from(endPr.attributes)) rPr.setAttribute(attr.name, attr.value);
    for (const c of children(endPr)) rPr.appendChild(c.cloneNode(true));
    r.appendChild(rPr);
  }
  const t = create(p, NS.a, 't');
  r.appendChild(t);
  p.insertBefore(r, endPr ?? null);
  writeRunText(r, t, text);
}
