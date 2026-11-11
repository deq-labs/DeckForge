/** File-level job handling for the desktop app, on top of the pure @deckforge/jobs package. */
import { mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { basename, extname, join } from 'node:path';
import type { Manifest } from '@deckforge/pptx-engine';
import { createWorkbook } from '@deckforge/data-xlsx';
import { check, loadTemplate, run } from '@deckforge/jobs';
import type { CheckResult, Problem } from '@deckforge/jobs';
import type { CheckView, GenerateView, JobSpec, ProblemView, SlotView, TemplateInfo } from '../api.ts';

export const WATERMARK = 'Made with DeckForge';

async function readUserFile(path: string, what: string): Promise<Uint8Array> {
  try {
    return new Uint8Array(await readFile(path));
  } catch (e) {
    const code = (e as NodeJS.ErrnoException).code;
    if (code === 'ENOENT') throw new Error(`The ${what} "${path}" can't be found. It may have been moved or renamed.`, { cause: e });
    if (code === 'EBUSY' || code === 'EPERM') throw new Error(`The ${what} "${basename(path)}" is locked. If it's open in another program, close it and try again.`, { cause: e });
    throw e;
  }
}

function stem(path: string): string {
  return basename(path, extname(path));
}

function toView(p: Problem): ProblemView {
  return {
    severity: p.severity,
    code: p.code,
    message: p.message,
    ...(p.hint ? { hint: p.hint } : {}),
    ...(p.slot ? { slot: p.slot } : {}),
    ...(p.batch ? { batch: p.batch } : {}),
  };
}

/** Human summary of what a template can be filled with (pure; tested). */
export function summarizeSlots(m: Manifest): SlotView[] {
  const unsupported = new Set(m.issues.filter((i) => i.code === 'unsupported' && i.slot).map((i) => i.slot!.toLowerCase()));
  const slots: SlotView[] = [];
  for (const t of m.text) {
    const where = [...(t.slides.length ? [`slide ${t.slides.join(', ')}`] : []), ...t.layouts.map((l) => `layout '${l}'`)].join('; ');
    const extra = t.fields.filter((f) => f !== 'value');
    slots.push({ kind: 'text', name: t.name, where, detail: extra.length ? `also ${extra.map((f) => `.${f}`).join(' ')}` : 'kpis sheet', fillable: true });
  }
  for (const t of m.tables) {
    slots.push({ kind: 'table', name: t.name, where: `slide ${t.slide}`, detail: t.columns.join(' | '), fillable: t.fillable && !unsupported.has(t.name.toLowerCase()) });
  }
  for (const c of m.charts) {
    const n = c.sample.series.length;
    slots.push({ kind: 'chart', name: c.name, where: `slide ${c.slide}`, detail: c.chartType ? `${c.chartType}, ${n} series` : 'not fillable', fillable: c.chartType !== undefined });
  }
  for (const r of m.repeats) {
    slots.push({ kind: 'repeat', name: r.name, where: `slide ${r.slides.join(', ')}`, detail: `one copy per row; columns ${r.tags.map((t) => t.name).join(', ')}`, fillable: true });
  }
  return slots;
}

export async function inspectTemplate(path: string): Promise<TemplateInfo> {
  let tpl;
  try {
    tpl = await loadTemplate(await readUserFile(path, 'template'));
  } catch (e) {
    throw new Error(`${basename(path)}: ${(e as Error).message}`, { cause: e });
  }
  const m = tpl.slots();
  return {
    path,
    fileName: basename(path),
    slideCount: m.slideCount,
    slots: summarizeSlots(m),
    rules: m.rules.filter((r) => r.kind !== 'repeat').map((r) => `Slide ${r.slide}: #${r.kind} ${r.target}`),
    problems: m.issues.map((i) => toView(i)),
  };
}

function runOptions(spec: JobSpec, watermark: string | undefined): Parameters<typeof run>[2] {
  return {
    templateName: stem(spec.templatePath),
    locale: spec.options.locale,
    currency: spec.options.currency,
    onMissing: spec.options.missing,
    ...(spec.namePattern ? { namePattern: spec.namePattern } : {}),
    ...(watermark ? { watermark } : {}),
  };
}

function checkView(r: CheckResult): CheckView {
  return { ok: r.ok, problems: r.problems.map(toView), batches: r.batches, deckCount: r.deckCount };
}

async function load(spec: JobSpec): Promise<{ tpl: Awaited<ReturnType<typeof loadTemplate>>; data: Uint8Array }> {
  const [tplBytes, data] = await Promise.all([readUserFile(spec.templatePath, 'template'), readUserFile(spec.dataPath, 'data workbook')]);
  return { tpl: await loadTemplate(tplBytes), data };
}

export async function checkJob(spec: JobSpec): Promise<CheckView> {
  const { tpl, data } = await load(spec);
  return checkView(await check(tpl, data, runOptions(spec, undefined)));
}

export async function generateJob(spec: JobSpec, watermark: string | undefined): Promise<GenerateView> {
  const { tpl, data } = await load(spec);
  const result = await run(tpl, data, runOptions(spec, watermark));
  const view: GenerateView = { ...checkView(result), outDir: spec.outDir, files: [], watermarked: Boolean(watermark) };
  if (!result.ok) return view;
  await mkdir(spec.outDir, { recursive: true });
  for (const d of result.decks) {
    const target = join(spec.outDir, d.fileName);
    const tmp = `${target}.${process.pid}.tmp`;
    try {
      await writeFile(tmp, d.bytes);
      await rename(tmp, target);
    } catch (e) {
      const code = (e as NodeJS.ErrnoException).code;
      if (code === 'EBUSY' || code === 'EPERM' || code === 'EACCES') {
        throw new Error(`Can't save "${d.fileName}". If it's open in PowerPoint, close it and generate again.`, { cause: e });
      }
      throw e;
    }
    view.files.push(target);
  }
  return view;
}

export async function writeDataFile(templatePath: string, outPath: string): Promise<void> {
  const tpl = await loadTemplate(await readUserFile(templatePath, 'template'));
  await writeFile(outPath, await createWorkbook(tpl.slots()));
}
