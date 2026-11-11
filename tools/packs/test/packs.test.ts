import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { check, loadTemplate } from '@deckforge/jobs';
import { PACKS_DIR, PACK_NAMES } from '../src/build.ts';

const expected: Record<string, { decks: number; batches: string[] }> = {
  'ops-review': { decks: 1, batches: [] },
  'kpi-scorecard': { decks: 1, batches: [] },
  'client-monthly': { decks: 3, batches: ['Harbor Dental', 'Summit Outdoor', 'Vela Coffee'] },
};

describe('template packs', () => {
  it('ships the three packs from the spec', () => {
    expect(PACK_NAMES).toEqual(Object.keys(expected));
  });

  for (const name of PACK_NAMES) {
    it(`${name}: the sample workbook fills the template with no errors or warnings`, async () => {
      const tpl = await loadTemplate(new Uint8Array(readFileSync(join(PACKS_DIR, name, 'template.pptx'))));
      const result = await check(tpl, new Uint8Array(readFileSync(join(PACKS_DIR, name, 'sample-data.xlsx'))), { templateName: name });
      expect(result.problems).toEqual([]);
      expect(result.ok).toBe(true);
      expect(result.deckCount).toBe(expected[name]!.decks);
      expect(result.batches).toEqual(expected[name]!.batches);
      // Every slot in the template is fed by the sample (nothing left as a {{tag}}).
      expect(result.manifest.issues).toEqual([]);
    });
  }
});
