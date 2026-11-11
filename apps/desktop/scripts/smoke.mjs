// End-to-end UI smoke test: launches the built app offscreen against a fixture, drives the real
// IPC API (inspect, check, generate), and saves a screenshot of the populated window.
// Usage: pnpm --filter @deckforge/desktop smoke [outDir]
import { spawnSync } from 'node:child_process';
import { copyFileSync, mkdirSync, readFileSync, rmSync } from 'node:fs';
import { createRequire } from 'node:module';
import { resolve } from 'node:path';

const out = resolve(process.argv[2] ?? '../../out/desktop-smoke');
rmSync(out, { recursive: true, force: true });
mkdirSync(out, { recursive: true });
const template = resolve(out, 'ops-review.pptx');
copyFileSync(resolve('../../fixtures/templates/ppt-ops-review.pptx'), template);

// A complete workbook: the template's own sample data plus numbers for every KPI.
const { Template, sampleData } = await import('@deckforge/pptx-engine');
const ExcelJS = (await import('exceljs')).default;
const { createWorkbook } = await import('@deckforge/data-xlsx');
const tpl = await Template.load(new Uint8Array(readFileSync(template)));
const wb = new ExcelJS.Workbook();
await wb.xlsx.load(await createWorkbook(tpl.slots()));
const values = sampleData(tpl.slots()).values;
wb.getWorksheet('kpis').eachRow((row, n) => {
  if (n === 1) return;
  const v = values[String(row.getCell(1).value)];
  row.getCell(2).value = typeof v === 'object' && v ? v.value : v;
  if (typeof v === 'object' && v && 'previous' in v) row.getCell(3).value = v.previous;
});
const data = resolve(out, 'data.xlsx');
await wb.xlsx.writeFile(data);

// Licensed run: a throwaway key pair, its public key passed via the dev-only override.
const { generateKeyPair, signLicense } = await import('@deckforge/licensing');
const kp = await generateKeyPair();
const licenseKey = await signLicense({ v: 1, id: 'SMOKE-1', product: 'deckforge', edition: 'lite', licensee: 'Smoke Test', seats: 1, issuedAt: new Date().toISOString(), updatesUntil: '2099-12-31' }, kp.privateKey);

const electron = createRequire(import.meta.url)('electron');
const cfg = { template, data, outDir: resolve(out, 'decks'), screenshot: resolve(out, 'window.png'), report: resolve(out, 'report.json'), licenseKey };
const r = spawnSync(electron, ['.'], { env: { ...process.env, DECKFORGE_SMOKE: JSON.stringify(cfg), DECKFORGE_LICENSE_PUBLIC_KEY: kp.publicKey }, timeout: 120_000, stdio: 'inherit' });
const report = JSON.parse(readFileSync(cfg.report, 'utf8'));
console.log(report);
const ok = r.status === 0 && report.hasApi && report.checkOk && report.files?.length === 1 && report.activatedMode === 'licensed' && report.watermarked === false && report.license === 'licensed';
console.log(ok ? `desktop smoke ok; screenshot: ${cfg.screenshot}` : 'desktop smoke FAILED');
process.exit(ok ? 0 : 1);
