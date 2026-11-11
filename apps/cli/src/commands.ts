import { basename, dirname, extname, join } from 'node:path';
import type { FillOptions, Manifest, OnMissing, Template } from '@deckforge/pptx-engine';
import { createWorkbook } from '@deckforge/data-xlsx';
import { check as checkJob, loadTemplate as loadTpl, run } from '@deckforge/jobs';
import type { Problem, RunOptions } from '@deckforge/jobs';
import { UserError, readInput, samePath, writeOutput } from './files.ts';
import { plural, printProblems } from './report.ts';
import type { Io } from './report.ts';

async function loadTemplate(path: string): Promise<Template> {
  const bytes = readInput(path, 'template');
  try {
    return await loadTpl(bytes);
  } catch (e) {
    throw new UserError(`${path}: ${(e as Error).message}`);
  }
}

function stem(path: string): string {
  return basename(path, extname(path));
}

function withSuffix(path: string, suffix: string, ext: string): string {
  return join(dirname(path), `${stem(path)}${suffix}${ext}`);
}

// ---------------------------------------------------------------- slots

export interface SlotsOptions {
  json?: boolean;
}

export async function slots(templatePath: string, opts: SlotsOptions, io: Io): Promise<number> {
  const m = (await loadTemplate(templatePath)).slots();
  if (opts.json) {
    io.out(JSON.stringify(m, null, 2));
    return 0;
  }
  io.out(describe(m, templatePath));
  printProblems(m.issues, io);
  return m.issues.some((i) => i.severity === 'error') ? 1 : 0;
}

function describe(m: Manifest, path: string): string {
  const lines: string[] = [`${basename(path)}: ${plural(m.slideCount, 'slide')}`];
  lines.push('', `Text tags (${m.text.length})  -> kpis sheet`);
  for (const t of m.text) {
    const extra = t.fields.filter((f) => f !== 'value');
    const where = [...(t.slides.length ? [`slide ${t.slides.join(', ')}`] : []), ...t.layouts.map((l) => `layout '${l}'`)].join('; ');
    lines.push(`  {{${t.name}}}${extra.length ? `  also ${extra.map((f) => `.${f}`).join(' ')}` : ''}  (${where})`);
  }
  lines.push('', `Tables (${m.tables.length})  -> one sheet each`);
  for (const t of m.tables) lines.push(`  ${t.name}  (slide ${t.slide}): ${t.columns.join(' | ')}`);
  lines.push('', `Charts (${m.charts.length})  -> one sheet each`);
  for (const c of m.charts) {
    const kind = c.chartType ?? 'not fillable';
    lines.push(`  ${c.name}  (slide ${c.slide}, ${kind}): ${c.sample.series.length} series — ${c.sample.series.map((s) => s.name).join(', ')}`);
  }
  if (m.repeats.length) {
    lines.push('', `Repeated slides (${m.repeats.length})  -> one sheet each, one row per copy`);
    for (const r of m.repeats) lines.push(`  ${r.name}  (slide ${r.slides.join(', ')}): columns ${r.tags.map((t) => t.name).join(', ')}`);
  }
  const conditional = m.rules.filter((r) => r.kind !== 'repeat');
  if (conditional.length) {
    lines.push('', 'Slide rules');
    for (const r of conditional) lines.push(`  slide ${r.slide}: #${r.kind} ${r.target}`);
  }
  return lines.join('\n');
}

// ---------------------------------------------------------------- init-data

export interface InitDataOptions {
  output?: string;
  sample?: boolean;
  force?: boolean;
}

export async function initData(templatePath: string, opts: InitDataOptions, io: Io): Promise<number> {
  const m = (await loadTemplate(templatePath)).slots();
  const out = opts.output ?? withSuffix(templatePath, '.data', '.xlsx');
  const bytes = await createWorkbook(m, { sample: opts.sample ?? true });
  writeOutput(out, bytes, { overwrite: opts.force ?? false, what: 'workbook' });
  printProblems(m.issues, io);
  io.out(`Created ${out}`);
  const parts = [
    `kpis: ${plural(m.text.length, 'row')} to fill in`,
    plural(m.tables.length, 'table sheet'),
    plural(m.charts.filter((c) => c.chartType).length, 'chart sheet'),
    ...(m.repeats.length ? [plural(m.repeats.length, 'repeat sheet')] : []),
  ];
  io.out(`  ${parts.join('; ')}.`);
  io.out(`Next: fill in the values, then run  deckforge fill "${templatePath}" "${out}"`);
  return 0;
}

// ---------------------------------------------------------------- fill / batch / check

export interface FillCliOptions {
  output?: string;
  locale?: string;
  currency?: string;
  missing?: string;
}

export interface BatchCliOptions extends FillCliOptions {
  outDir?: string;
  /** File name pattern; {template} and {batch} are replaced. */
  name?: string;
}

const MISSING: OnMissing[] = ['error', 'keep', 'blank'];

function fillOptions(opts: FillCliOptions): FillOptions {
  const missing = (opts.missing ?? 'error') as OnMissing;
  if (!MISSING.includes(missing)) throw new UserError(`--missing must be one of: ${MISSING.join(', ')}.`);
  return {
    ...(opts.locale ? { locale: opts.locale } : {}),
    ...(opts.currency ? { currency: opts.currency.toUpperCase() } : {}),
    onMissing: missing,
  };
}

function wrapRead<T>(dataPath: string, f: () => Promise<T>): Promise<T> {
  return f().catch((e: unknown) => {
    if (e instanceof Error && /Excel workbook/.test(e.message)) throw new UserError(`${dataPath}: ${e.message}`);
    throw e;
  });
}

const keepHint = (opts: FillOptions): string => (opts.onMissing === 'error' ? ' (or use --missing keep to leave unfilled tags visible)' : '');

export async function fill(templatePath: string, dataPath: string, opts: FillCliOptions, io: Io): Promise<number> {
  const options = fillOptions(opts);
  const out = opts.output ?? withSuffix(templatePath, '.filled', '.pptx');
  if (samePath(out, templatePath)) throw new UserError('The output would overwrite the template. Choose another name with -o.');
  if (samePath(out, dataPath)) throw new UserError('The output would overwrite the data workbook. Choose another name with -o.');

  const tpl = await loadTemplate(templatePath);
  const data = readInput(dataPath, 'data workbook');
  const runOpts: RunOptions = { ...options, templateName: stem(templatePath) };
  const pre = await wrapRead(dataPath, () => checkJob(tpl, data, runOpts));
  if (pre.batches.length > 1) {
    throw new UserError(
      `The workbook has a batch column with ${pre.batches.length} values (${pre.batches.join(', ')}). ` +
        `Use  deckforge batch "${templatePath}" "${dataPath}"  to write one deck per value.`,
    );
  }
  const result = await run(tpl, data, runOpts);
  printProblems(result.problems, io);
  if (!result.ok) {
    io.err(`\nNo deck was written. Fix the problems above and run again${keepHint(options)}.`);
    return 1;
  }
  writeOutput(out, result.decks[0]!.bytes, { overwrite: true, what: 'deck' });
  io.out(`Wrote ${out}`);
  return 0;
}

export async function batch(templatePath: string, dataPath: string, opts: BatchCliOptions, io: Io): Promise<number> {
  const options = fillOptions(opts);
  const tpl = await loadTemplate(templatePath);
  const data = readInput(dataPath, 'data workbook');
  const result = await wrapRead(dataPath, () =>
    run(tpl, data, { ...options, templateName: stem(templatePath), namePattern: opts.name ?? '{template}-{batch}.pptx' }),
  );
  if (result.ok && result.batches.length === 0) {
    io.err('warning: the workbook has no batch column, so this makes a single deck. Add a column named "batch" to any sheet to split it.');
  }
  printProblems(result.problems, io);
  if (!result.ok) {
    io.err(`\nNo decks were written. Fix the problems above and run again${keepHint(options)}.`);
    return 1;
  }
  const outDir = opts.outDir ?? join(dirname(templatePath), `${stem(templatePath)}-decks`);
  for (const d of result.decks) writeOutput(join(outDir, d.fileName), d.bytes, { overwrite: true, what: 'deck' });
  io.out(`Wrote ${plural(result.decks.length, 'deck')} to ${outDir}`);
  for (const d of result.decks) io.out(`  ${d.fileName}`);
  return 0;
}

export async function check(templatePath: string, dataPath: string | undefined, opts: FillCliOptions, io: Io): Promise<number> {
  const options = fillOptions(opts);
  const tpl = await loadTemplate(templatePath);
  const data = dataPath ? readInput(dataPath, 'data workbook') : undefined;
  const result = await wrapRead(dataPath ?? '', () => checkJob(tpl, data, { ...options, templateName: stem(templatePath) }));
  printProblems(result.problems, io);
  const errors = result.problems.filter((p: Problem) => p.severity === 'error').length;
  const summary = `${plural(errors, 'error')}, ${plural(result.problems.length - errors, 'warning')}`;
  const what = `${basename(templatePath)}${dataPath ? ` + ${basename(dataPath)}` : ''}`;
  if (!result.ok) {
    io.out(`\n${what}: ${summary}. Fix the errors before generating.`);
    return 1;
  }
  const decks = dataPath ? `; ${plural(result.deckCount, 'deck')} would be made` : '';
  io.out(`\n${what}: ready (${summary}${decks}).`);
  return 0;
}
