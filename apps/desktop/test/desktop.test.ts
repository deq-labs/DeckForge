import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { Template } from '@deckforge/pptx-engine';
import type { CheckView, Recipe, SlotView } from '../electron/api.ts';
import { RecipeStore } from '../electron/services/recipes.ts';
import { candidatePaths, chooseTool, detectTools } from '../electron/services/preview.ts';
import { inspectTemplate, summarizeSlots } from '../electron/services/jobs.ts';
import { defaultOutDir, relativeTime, slotStatus, splitPath, stem } from '../src/logic.ts';

const FIXTURES = resolve(import.meta.dirname, '../../../fixtures/templates');

let dir: string;
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'deckforge-desktop-'));
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

const recipe = (id: string, name = id): Recipe => ({
  id,
  name,
  spec: { templatePath: 'C:\\t.pptx', dataPath: 'C:\\d.xlsx', outDir: 'C:\\out', options: { locale: 'en-US', currency: 'USD', missing: 'error' } },
});

describe('RecipeStore', () => {
  it('saves (newest first), replaces by id, and removes', async () => {
    const store = new RecipeStore(dir);
    expect(await store.list()).toEqual([]);
    await store.save(recipe('a'));
    await store.save(recipe('b'));
    expect((await store.save(recipe('a', 'A2'))).map((r) => r.name)).toEqual(['A2', 'b']);
    expect((await store.remove('b')).map((r) => r.id)).toEqual(['a']);
    expect((await new RecipeStore(dir).list()).map((r) => r.id)).toEqual(['a']);
  });

  it('survives a damaged file by setting it aside', async () => {
    writeFileSync(join(dir, 'recipes.json'), '{not json');
    const store = new RecipeStore(dir);
    expect(await store.list()).toEqual([]);
    await store.save(recipe('x'));
    expect(JSON.parse(readFileSync(join(dir, 'recipes.json'), 'utf8')).recipes).toHaveLength(1);
  });

  it('rejects incomplete recipes', async () => {
    await expect(new RecipeStore(dir).save({ id: 'x', name: 'x' } as Recipe)).rejects.toThrow(/incomplete/);
  });
});

describe('preview tool detection', () => {
  it('prefers PowerPoint, falls back to LibreOffice, else none', () => {
    expect(chooseTool({ powerpoint: 'p', libreoffice: 'l' })).toBe('powerpoint');
    expect(chooseTool({ libreoffice: 'l' })).toBe('libreoffice');
    expect(chooseTool({})).toBeNull();
  });

  it('looks in the usual places per platform', () => {
    const win = candidatePaths('win32', { ProgramFiles: 'C:\\PF' });
    expect(win.powerpoint[0]).toBe('C:\\PF\\Microsoft Office\\root\\Office16\\POWERPNT.EXE');
    expect(win.libreoffice[0]).toBe('C:\\PF\\LibreOffice\\program\\soffice.exe');
    expect(candidatePaths('darwin', {}).libreoffice).toEqual(['/Applications/LibreOffice.app/Contents/MacOS/soffice']);
    const found = detectTools('linux', {}, (p) => p === '/usr/bin/soffice');
    expect(found).toEqual({ libreoffice: '/usr/bin/soffice' });
  });
});

describe('template summary', () => {
  it('lists every slot with where it is and what it needs', async () => {
    const m = (await Template.load(new Uint8Array(readFileSync(join(FIXTURES, 'l2-features.pptx'))))).slots();
    const slots = summarizeSlots(m);
    expect(slots.find((s) => s.name === 'company')).toMatchObject({ kind: 'text', where: "layout 'BRANDED'" });
    expect(slots.find((s) => s.name === 'volume_yield')).toMatchObject({ kind: 'chart', detail: 'combo, 2 series', fillable: true });
    expect(slots.find((s) => s.name === 'sites')).toMatchObject({ kind: 'repeat', detail: 'one copy per row; columns site, oee, manager' });
  });

  it('explains a file that is missing or not a deck', async () => {
    await expect(inspectTemplate(join(dir, 'nope.pptx'))).rejects.toThrow(/can't be found/);
    writeFileSync(join(dir, 'bad.pptx'), 'hello');
    await expect(inspectTemplate(join(dir, 'bad.pptx'))).rejects.toThrow(/bad\.pptx: This file is not a PowerPoint/);
  });
});

describe('renderer logic', () => {
  it('handles Windows and POSIX paths', () => {
    expect(splitPath('C:\\a\\b\\deck.pptx')).toEqual({ dir: 'C:\\a\\b', file: 'deck.pptx', sep: '\\' });
    expect(stem('/x/y/ops review.pptx')).toBe('ops review');
    expect(defaultOutDir('C:\\a\\ops.pptx')).toBe('C:\\a\\ops-decks');
    expect(defaultOutDir('/a/ops.pptx')).toBe('/a/ops-decks');
  });

  it('marks slot status from the last check', () => {
    const slot: SlotView = { kind: 'table', name: 'Downtime', where: 'slide 2', detail: '', fillable: true };
    const check = (slotName?: string): CheckView => ({
      ok: !slotName,
      batches: [],
      deckCount: 1,
      problems: slotName ? [{ severity: 'error', code: 'missing-table', message: 'x', slot: slotName }] : [],
    });
    expect(slotStatus(slot, null)).toBe('pending');
    expect(slotStatus(slot, check())).toBe('ok');
    expect(slotStatus(slot, check('downtime'))).toBe('problem');
    expect(slotStatus({ ...slot, fillable: false }, check())).toBe('unsupported');
  });

  it('describes when a recipe last ran', () => {
    const now = Date.parse('2026-11-11T12:00:00Z');
    expect(relativeTime(undefined, now)).toBe('never run');
    expect(relativeTime('2026-11-11T11:59:50Z', now)).toBe('just now');
    expect(relativeTime('2026-11-11T09:00:00Z', now)).toBe('3 hours ago');
    expect(relativeTime('2026-10-12T12:00:00Z', now)).toBe('30 days ago');
  });
});
