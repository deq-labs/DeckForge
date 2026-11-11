/**
 * Smoke test for output validity. For every template in fixtures/templates/:
 *   - fill it with its own sample data, with more rows/points than the template, and with none;
 *   - write the outputs to out/smoke/;
 *   - check each output against the Open XML schemas (tools/ooxml-validate, needs the .NET SDK);
 *   - convert each output to PDF with LibreOffice (needs soffice).
 *
 * Run: pnpm smoke [--require-all]
 * With --require-all a missing tool is a failure (CI); otherwise that check is skipped.
 */
import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Template, sampleData } from '@deckforge/pptx-engine';
import type { DeckData } from '@deckforge/pptx-engine';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const templatesDir = join(root, 'fixtures/templates');
const outDir = join(root, 'out/smoke');
const requireAll = process.argv.includes('--require-all');

function grow(d: DeckData): DeckData {
  const tables = Object.fromEntries(Object.entries(d.tables ?? {}).map(([k, t]) => [k, { ...t, rows: [...t.rows, ...t.rows, ...t.rows.slice(0, 1)] }]));
  const charts = Object.fromEntries(
    Object.entries(d.charts ?? {}).map(([k, c]) => [
      k,
      { categories: [...c.categories, 'Extra'], series: c.series.map((s) => ({ ...s, values: [...s.values, 1] })) },
    ]),
  );
  return { ...d, tables, charts };
}

function empty(d: DeckData): DeckData {
  const tables = Object.fromEntries(Object.entries(d.tables ?? {}).map(([k, t]) => [k, { ...t, rows: [] }]));
  const charts = Object.fromEntries(
    Object.entries(d.charts ?? {}).map(([k, c]) => [k, { categories: [], series: c.series.map((s) => ({ ...s, values: [] })) }]),
  );
  return { ...d, tables, charts };
}

function findSoffice(): string | undefined {
  const candidates = [
    'C:/Program Files/LibreOffice/program/soffice.exe',
    'C:/Program Files (x86)/LibreOffice/program/soffice.exe',
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
  ];
  for (const c of candidates) if (existsSync(c)) return c;
  const probe = spawnSync(process.platform === 'win32' ? 'where' : 'which', ['soffice'], { encoding: 'utf8' });
  return probe.status === 0 ? probe.stdout.split(/\r?\n/)[0]?.trim() : undefined;
}

function hasDotnet(): boolean {
  return spawnSync('dotnet', ['--version'], { encoding: 'utf8' }).status === 0;
}

let failures = 0;
const fail = (msg: string): void => {
  failures++;
  console.error(`FAIL ${msg}`);
};

rmSync(outDir, { recursive: true, force: true });
mkdirSync(outDir, { recursive: true });

const templates = readdirSync(templatesDir).filter((f) => f.endsWith('.pptx'));
const outputs: string[] = [];
for (const file of templates) {
  const tpl = await Template.load(new Uint8Array(readFileSync(join(templatesDir, file))));
  const base = sampleData(tpl.slots());
  for (const [label, data] of [['sample', base], ['grown', grow(base)], ['empty', empty(base)]] as const) {
    const out = join(outDir, `${file.replace(/\.pptx$/, '')}.${label}.pptx`);
    try {
      const deck = await tpl.fill(data);
      writeFileSync(out, deck.toBuffer());
      outputs.push(out);
      console.log(`ok   filled ${file} (${label})`);
    } catch (e) {
      fail(`fill ${file} (${label}): ${(e as Error).message}`);
    }
  }
}

if (hasDotnet()) {
  const project = join(root, 'tools/ooxml-validate');
  const args = ['run', '--project', project, '-v', 'q', '--', ...outputs, '--baseline', ...templates.map((t) => join(templatesDir, t))];
  const r = spawnSync('dotnet', args, { encoding: 'utf8' });
  const blocked = /Application Control policy has blocked/.test(`${r.stdout}${r.stderr}`);
  if (blocked && !requireAll) {
    // Windows Smart App Control / WDAC can refuse locally built, unsigned assemblies. CI still validates.
    console.log('skip Open XML validation (Windows Application Control blocked the locally built validator)');
  } else {
    process.stdout.write(r.stdout);
    process.stderr.write(r.stderr);
    if (r.status !== 0) fail('Open XML schema validation');
  }
} else if (requireAll) {
  fail('dotnet not found; cannot run the Open XML validator');
} else {
  console.log('skip Open XML validation (dotnet not found)');
}

const soffice = findSoffice();
if (soffice) {
  const pdfDir = join(outDir, 'pdf');
  mkdirSync(pdfDir, { recursive: true });
  for (const out of outputs) {
    try {
      execFileSync(soffice, ['--headless', '--norestore', '--convert-to', 'pdf', '--outdir', pdfDir, out], { stdio: 'pipe', timeout: 120_000 });
      const pdf = join(pdfDir, out.split(/[\\/]/).pop()!.replace(/\.pptx$/, '.pdf'));
      if (!existsSync(pdf) || readFileSync(pdf).byteLength < 1000) fail(`LibreOffice produced no PDF for ${out}`);
      else console.log(`ok   LibreOffice opened ${out}`);
    } catch (e) {
      fail(`LibreOffice could not convert ${out}: ${(e as Error).message}`);
    }
  }
} else if (requireAll) {
  fail('LibreOffice (soffice) not found');
} else {
  console.log('skip LibreOffice round trip (soffice not found)');
}

if (failures > 0) {
  console.error(`\n${failures} smoke check(s) failed`);
  process.exit(1);
}
console.log(`\nsmoke: ${outputs.length} decks ok`);
