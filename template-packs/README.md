# DeckForge template packs

Ready-made templates, each with a sample workbook that fills it completely. Open the sample in
Excel, replace the numbers with yours, and generate. Or copy the tagging ideas into your own deck.

| Pack | For | Shows off |
| --- | --- | --- |
| [ops-review](ops-review/) | Monthly operations review: OEE, output, scrap, downtime | KPI tiles with trend arrows, a growing table, column + line charts |
| [kpi-scorecard](kpi-scorecard/) | Quarterly quality and delivery scorecard | Mixed-format table, pie chart, combo chart (columns + line on a second axis), a slide that only appears when there were incidents |
| [client-monthly](client-monthly/) | Agencies and consultants: one report per client | **Batch decks**: one workbook makes a deck for every client |

Each folder has `template.pptx`, `sample-data.xlsx` and `preview/` (slides rendered by PowerPoint).

## For maintainers

Templates are built in real PowerPoint by `tools/powerpoint/build-templates.ps1` (into
`fixtures/templates/`); `pnpm packs` copies them here and writes the sample workbooks from
`tools/packs/src/build.ts`. `pnpm --filter @deckforge/packs test` checks that each sample fills its
template with no errors or warnings. Previews:

```sh
node tools/packs/src/examples.ts
powershell -File tools/powerpoint/verify-open.ps1 -Png 'out/pack-examples/*.pptx'
# then copy out/pack-examples/<deck>.slides/*.png into template-packs/<pack>/preview/
```
