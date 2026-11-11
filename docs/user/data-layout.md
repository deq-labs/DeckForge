# The data workbook

One Excel workbook (.xlsx) feeds one template. **Create data file…** in the app (or
`deckforge init-data`) makes it for you; this page explains what's in it.

## `kpis` sheet — one row per tag

| key | value | previous | format | notes |
| --- | --- | --- | --- | --- |
| oee | 0.784 | 0.751 | | slide 2; uses .trend |
| site | Riverside Plant | | | slide 1 |

- **key**: the tag name without braces.
- **value**: what to show. An empty value counts as missing (so a half-filled sheet is caught).
- **previous**: optional, for `.previous`, `.change`, `.change_pct` and `.trend`.
- **format**: optional. Leave it empty to use the cell's **Excel number format** (format the cell
  as 0.0% and the slide shows 78.4%). Or write one of: `text`, `number`, `percent`, `currency`,
  `date`, with decimals like `percent:1` or `currency:0`.
- **notes**: ignored by DeckForge — yours.

Percentages are fractions, as in Excel: 0.784 shows as 78.4%.

## `meta` sheet (optional)

`key` and `value` columns for plain text (titles, periods, comments). Same as `kpis`, without
`previous` and `format`. A key may be in `kpis` or `meta`, not both.

## One sheet per table

Named like the table. Row 1 = the table's header texts; rows below = the table body. Blank rows
are skipped. Number formats from Excel carry over cell by cell.

## One sheet per chart

Named like the chart. Cell A1 = category label; column A = categories; columns B onward =
one series each, series names in row 1. Empty cells are gaps. Values must be numbers.

## One sheet per repeated slide

For a slide with `#repeat sites` in its notes: a `sites` sheet, one row per copy, one column per
tag ("Plant Manager" fills `{{plant_manager}}`). A column named "oee previous" feeds
`{{oee.previous}}` and the trend fields.

## Batch decks: the `batch` column

Add a column named **batch** to any sheet. Each different value makes its own deck with:

- the rows that have that value, plus
- the rows whose batch cell is empty (shared by every deck).

In `kpis`, a batch row overrides a shared row with the same key. Example: `month` once with an
empty batch; `client` and the numbers once per client.

## Other rules

- Sheet names match table/chart names without regard to case.
- Sheets whose names start with `_` are ignored — use them for notes or raw exports.
- Formulas are fine; DeckForge uses their last calculated result (save the file in Excel first).
- `kpis` and `meta` are reserved sheet names.
