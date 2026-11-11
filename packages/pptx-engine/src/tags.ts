/**
 * `{{tag}}` discovery and replacement inside DrawingML paragraphs (`a:p`).
 *
 * PowerPoint often splits what the user typed as one tag across several runs (spell-check
 * marks, partial formatting, autocorrect). We treat each unbroken sequence of `a:r` runs as
 * one string, find tags in it, and write the result back so the tag's text lands in the run
 * where `{{` started (keeping that run's formatting) and is removed from the runs after it.
 * Line breaks (`a:br`) and fields (`a:fld`) end a sequence: a tag cannot span them.
 */
import { NS, child, children, create, is, removeNode, setText, textOf } from './xml.ts';
import type { Element } from './xml.ts';
import { TEXT_FIELDS } from './types.ts';
import type { TextField } from './types.ts';

export const SLOT_NAME_RE = /^[A-Za-z_][A-Za-z0-9_-]*$/;
const TAG_RE = /\{\{([^{}\r\n]*)\}\}/g;

export type ParsedTag =
  | { ok: true; name: string; field: TextField }
  | { ok: false; raw: string; reason: string };

export function parseTag(inner: string): ParsedTag {
  const raw = inner.trim();
  if (raw === '') return { ok: false, raw, reason: 'the tag is empty' };
  const dot = raw.indexOf('.');
  const name = dot < 0 ? raw : raw.slice(0, dot);
  const field = dot < 0 ? 'value' : raw.slice(dot + 1).toLowerCase();
  if (!SLOT_NAME_RE.test(name)) {
    return {
      ok: false,
      raw,
      reason: 'tag names may only use letters, digits, _ and -, must start with a letter or _, and cannot contain spaces',
    };
  }
  if (!(TEXT_FIELDS as readonly string[]).includes(field)) {
    return { ok: false, raw, reason: `".${field}" is not a known suffix; use .previous, .change, .change_pct or .trend` };
  }
  return { ok: true, name, field: field as TextField };
}

interface RunText {
  run: Element;
  t: Element;
  text: string;
}

function segments(p: Element): RunText[][] {
  const segs: RunText[][] = [];
  let cur: RunText[] = [];
  for (const el of children(p)) {
    if (is(el, NS.a, 'r')) {
      const t = child(el, NS.a, 't');
      if (t) cur.push({ run: el, t, text: textOf(t) });
      continue;
    }
    if (is(el, NS.a, 'pPr') || is(el, NS.a, 'endParaRPr')) continue;
    if (cur.length) segs.push(cur);
    cur = [];
  }
  if (cur.length) segs.push(cur);
  return segs;
}

/** Inner text of every `{{...}}` in the paragraph, in reading order. */
export function findTags(p: Element): string[] {
  const out: string[] = [];
  for (const seg of segments(p)) {
    const full = seg.map((r) => r.text).join('');
    for (const m of full.matchAll(TAG_RE)) out.push(m[1] ?? '');
  }
  return out;
}

/**
 * Replaces tags in a paragraph. `resolve` gets the tag's inner text and returns the
 * replacement, or `undefined` to leave the tag untouched.
 */
export function replaceTags(p: Element, resolve: (inner: string) => string | undefined): void {
  for (const seg of segments(p)) {
    const texts = seg.map((r) => r.text);
    const full = texts.join('');
    const matches = [...full.matchAll(TAG_RE)];
    if (matches.length === 0) continue;

    const starts: number[] = [];
    let acc = 0;
    for (const t of texts) {
      starts.push(acc);
      acc += t.length;
    }
    const runAt = (pos: number): number => {
      let i = 0;
      while (i + 1 < starts.length && starts[i + 1]! <= pos) i++;
      return i;
    };

    const touched = new Set<number>();
    // Right to left, so offsets to the left of each match stay valid.
    for (const m of matches.reverse()) {
      const repl = resolve(m[1] ?? '');
      if (repl === undefined) continue;
      const s = m.index;
      const e = s + m[0].length;
      const i = runAt(s);
      const j = runAt(e - 1);
      const si = s - starts[i]!;
      const ej = e - starts[j]!;
      if (i === j) {
        texts[i] = texts[i]!.slice(0, si) + repl + texts[i]!.slice(ej);
      } else {
        texts[i] = texts[i]!.slice(0, si) + repl;
        for (let k = i + 1; k < j; k++) texts[k] = '';
        texts[j] = texts[j]!.slice(ej);
      }
      for (let k = i; k <= j; k++) touched.add(k);
    }

    for (const k of touched) {
      const { run, t } = seg[k]!;
      const text = texts[k]!;
      if (text === '') removeNode(run);
      else writeRunText(run, t, text);
    }
  }
}

/** Sets a run's text; newlines become `a:br` line breaks with the run's formatting. */
export function writeRunText(run: Element, t: Element, text: string): void {
  const lines = text.split(/\r\n|\r|\n/);
  setText(t, lines[0] ?? '');
  const rPr = child(run, NS.a, 'rPr');
  let after: Element = run;
  for (const line of lines.slice(1)) {
    const br = create(run, NS.a, 'br');
    if (rPr) br.appendChild(rPr.cloneNode(true));
    insertAfter(br, after);
    const next = run.cloneNode(true) as Element;
    setText(child(next, NS.a, 't')!, line);
    insertAfter(next, br);
    after = next;
  }
}

function insertAfter(node: Element, ref: Element): void {
  ref.parentNode!.insertBefore(node, ref.nextSibling);
}
