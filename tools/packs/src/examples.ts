/**
 * Writes the example decks for each pack (no trial footer) to out/pack-examples/, for preview
 * images and the landing page. Render them with PowerPoint:
 *   powershell -File tools/powerpoint/verify-open.ps1 -Png 'out/pack-examples/*.pptx'
 */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { loadTemplate, run } from '@deckforge/jobs';
import { PACKS_DIR, PACK_NAMES } from './build.ts';

const out = resolve(PACKS_DIR, '../out/pack-examples');
mkdirSync(out, { recursive: true });
for (const name of PACK_NAMES) {
  const tpl = await loadTemplate(new Uint8Array(readFileSync(join(PACKS_DIR, name, 'template.pptx'))));
  const result = await run(tpl, new Uint8Array(readFileSync(join(PACKS_DIR, name, 'sample-data.xlsx'))), { templateName: name, namePattern: '{template}-{batch}.pptx' });
  if (!result.ok) throw new Error(`${name}: ${result.problems.map((p) => p.message).join('; ')}`);
  for (const d of result.decks) {
    const file = d.batch ? d.fileName : `${name}.pptx`;
    writeFileSync(join(out, file), d.bytes);
    console.log(`wrote out/pack-examples/${file}`);
  }
}
