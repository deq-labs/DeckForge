/**
 * Check / fill / batch orchestration shared by the CLI and the desktop app.
 * Pure: bytes in, decks and plain-language issues out. Callers do the file I/O.
 */
import { FillError, Template } from '@deckforge/pptx-engine';
import type { FillOptions, Issue, Manifest } from '@deckforge/pptx-engine';
import { readWorkbook } from '@deckforge/data-xlsx';
import type { BatchDeck } from '@deckforge/data-xlsx';

/** An issue with an optional "how to fix it" line for the Lite workbook layout. */
export interface Problem extends Issue {
  hint?: string;
  /** The batch the problem belongs to, when the workbook has batches. */
  batch?: string;
}

export interface CheckResult {
  manifest: Manifest;
  problems: Problem[];
  /** Batch values (empty when the workbook has no batch column). */
  batches: string[];
  /** Decks a run would produce; 0 when there are errors. */
  deckCount: number;
  ok: boolean;
}

export interface OutputDeck {
  /** Batch value, or null for a single deck. */
  batch: string | null;
  fileName: string;
  bytes: Uint8Array;
}

export interface RunResult extends CheckResult {
  decks: OutputDeck[];
}

export interface RunOptions extends FillOptions {
  /** Template file name without extension, used in output names. */
  templateName: string;
  /** Output name pattern; `{template}` and `{batch}` are replaced. Default `{template}-{batch}.pptx`, or `{template}.filled.pptx` without batches. */
  namePattern?: string;
}

export async function loadTemplate(bytes: Uint8Array): Promise<Template> {
  return Template.load(bytes);
}

/** Fix-it hints for the Lite workbook layout, keyed by the engine's issue codes. */
export function hintFor(issue: Issue, ctx: { manifest?: Manifest; blankKeys?: string[] }): string | undefined {
  const slot = issue.slot;
  if (!slot) return undefined;
  const key = slot.toLowerCase();
  switch (issue.code) {
    case 'missing-value': {
      if (ctx.blankKeys?.includes(key)) return `Fill in the value for "${slot}" in the kpis sheet (its value cell is empty).`;
      const repeat = ctx.manifest?.repeats.find((r) => r.tags.some((t) => t.name.toLowerCase() === key));
      if (repeat) return `Add a column "${slot}" to the "${repeat.name}" sheet.`;
      return `Add a row with key "${slot}" to the kpis sheet (or the meta sheet).`;
    }
    case 'missing-rows':
      return `Add a sheet named "${slot}" with one row per slide copy; its column headers are the tag names.`;
    case 'missing-table':
      return `Add a sheet named "${slot}" with the table's header row in row 1.`;
    case 'missing-chart':
      return `Add a sheet named "${slot}": categories in column A, one series per column from B.`;
    default:
      return undefined;
  }
}

/** Characters Windows (and so most file shares) won't accept in a file name. */
export function safeFileName(s: string): string {
  // eslint-disable-next-line no-control-regex
  const cleaned = s.replace(/[<>:"/\\|?*\u0000-\u001f]/g, '-').replace(/[. ]+$/, '').trim();
  return cleaned || 'deck';
}

function outputName(opts: RunOptions, batch: string | null): string {
  const pattern = opts.namePattern ?? (batch === null ? '{template}.filled.pptx' : '{template}-{batch}.pptx');
  let name = safeFileName(pattern.replaceAll('{template}', opts.templateName).replaceAll('{batch}', batch ?? 'all'));
  if (!name.toLowerCase().endsWith('.pptx')) name += '.pptx';
  return name;
}

function tagBatch(issues: Issue[], batch: string | null, exclude: Issue[]): Problem[] {
  return issues
    .filter((i) => !exclude.some((m) => m.message === i.message))
    .map((i) => (batch === null ? { ...i } : { ...i, batch, message: `[${batch}] ${i.message}` }));
}

function dedupe(problems: Problem[]): Problem[] {
  const seen = new Set<string>();
  return problems.filter((p) => (seen.has(p.message) ? false : (seen.add(p.message), true)));
}

/**
 * Checks a template (and optionally a workbook) without producing files, or produces the decks.
 * With `produce: false` this is the template checker; with a workbook it is a full dry run.
 */
async function execute(template: Template, data: Uint8Array | undefined, options: RunOptions, produce: boolean): Promise<RunResult> {
  const manifest = template.slots();
  const problems: Problem[] = [...manifest.issues];
  let blankKeys: string[] = [];
  let decks: BatchDeck[] = [];
  const out: OutputDeck[] = [];
  let failed = false;

  if (data) {
    const read = await readWorkbook(data, manifest);
    blankKeys = read.blankKeys;
    problems.push(...read.issues);
    decks = read.decks;
    if (!read.issues.some((i) => i.severity === 'error')) {
      const names = new Set<string>();
      for (const { batch, data: deckData } of decks) {
        const fileName = outputName(options, decks.length > 1 || batch !== null ? batch : null);
        if (names.has(fileName.toLowerCase())) {
          problems.push({
            severity: 'error',
            code: 'data-layout',
            message: `Two batches would both be saved as "${fileName}". Make the batch values differ by more than punctuation.`,
          });
          failed = true;
          continue;
        }
        names.add(fileName.toLowerCase());
        try {
          const deck = await template.fill(deckData, options);
          problems.push(...tagBatch(deck.issues, batch, manifest.issues));
          if (produce) out.push({ batch, fileName, bytes: deck.toBuffer() });
        } catch (e) {
          if (!(e instanceof FillError)) throw e;
          problems.push(...tagBatch(e.issues, batch, manifest.issues));
          failed = true;
        }
      }
    }
  }

  const all = dedupe(problems).map((p) => {
    const hint = hintFor(p, { manifest, blankKeys });
    return hint ? { ...p, hint } : p;
  });
  const ok = !failed && !all.some((p) => p.severity === 'error');
  return {
    manifest,
    problems: [...all.filter((p) => p.severity === 'error'), ...all.filter((p) => p.severity === 'warning')],
    batches: decks.map((d) => d.batch).filter((b): b is string => b !== null),
    deckCount: ok ? decks.length : 0,
    ok,
    // All or nothing: never hand back a partial set of decks.
    decks: ok ? out : [],
  };
}

export async function check(template: Template, data: Uint8Array | undefined, options: RunOptions): Promise<CheckResult> {
  const { decks: _decks, ...result } = await execute(template, data, options, false);
  return result;
}

export async function run(template: Template, data: Uint8Array, options: RunOptions): Promise<RunResult> {
  return execute(template, data, options, true);
}
