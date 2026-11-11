# DeckForge

Fill your own PowerPoint template with data from Excel and get native, editable decks —
on your own machine, with no data leaving your network.

> Status: early development (Lite track, L1). Private for now. See [`docs/spec.md`](docs/spec.md).

## Try it (from source)

Requires Node.js 24+ and pnpm.

```sh
pnpm install
pnpm deckforge slots fixtures/templates/ops-review.pptx         # what can be filled
pnpm deckforge init-data fixtures/templates/ops-review.pptx -o out/ops.xlsx
# ...fill in the kpis sheet in Excel...
pnpm deckforge fill fixtures/templates/ops-review.pptx out/ops.xlsx -o out/ops.pptx
```

## Tagging a template

- Text: type `{{name}}` anywhere in slide text. `{{name.previous}}`, `.change`, `.change_pct`
  and `.trend` use the `previous` column.
- Tables and charts: name the shape in PowerPoint's **Selection Pane** (Home → Arrange →
  Selection Pane), e.g. `downtime`, and add a sheet with the same name to the workbook.
  Names use letters, digits, `_` and `-` (no spaces).

- Slide rules: in a slide's speaker notes, `#repeat sites` makes one copy per row of the
  `sites` sheet; `#show-if incidents` / `#hide-if key` keep or drop the slide.
- Tags also work in speaker notes, chart titles, and normal text boxes on slide layouts/masters.
- Batch decks: add a `batch` column to any sheet, then `deckforge batch template.pptx data.xlsx`.
- `deckforge check template.pptx [data.xlsx]` lists problems without writing anything.

Full rules: [`docs/design-log/0003-lite-tagging-and-data-contract.md`](docs/design-log/0003-lite-tagging-and-data-contract.md).

## Repository layout

| Path | What |
| --- | --- |
| `packages/pptx-engine` | Pure TS engine: discover slots, fill text/tables/charts, write OOXML |
| `packages/data-xlsx` | Reads/creates the Lite workbook layout (batches, repeat sheets) |
| `packages/jobs` | Check / fill / batch orchestration + fix-it hints, shared by CLI and desktop |
| `apps/cli` | `deckforge slots | init-data | fill | batch | check` |
| `tools/fixtures` | Synthetic fixture generator (`pnpm fixtures`) and output smoke test (`pnpm smoke`) |
| `tools/ooxml-validate` | Open XML SDK schema validator (.NET) used by the smoke test |
| `fixtures/templates` | Test templates |
| `docs/design-log` | Design decisions |

## Development

`pnpm check` runs typecheck, lint and tests. `pnpm smoke` fills every fixture and validates the
output (Open XML SDK if `dotnet` is installed, LibreOffice if `soffice` is installed; CI runs both).
