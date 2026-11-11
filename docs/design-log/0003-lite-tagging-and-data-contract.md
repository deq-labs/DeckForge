# 0003 — Lite tagging syntax and workbook contract

- Date: 2026-11-11 · Edition: lite (engine parts: shared) · Status: accepted for L1; L2 items marked *reserved*

## Context
Self-setup in under 30 minutes only works if a buyer can tag a template in PowerPoint and
fill a workbook without reading code. The spec left the slot naming, the `kpis`/`meta`
relationship and chart/table matching rules open.

## Decision

### Slot names
A slot name is `[A-Za-z_][A-Za-z0-9_-]*`, at most 31 characters (Excel's sheet-name limit).
Names are matched **case-insensitively**. Two slots whose names differ only by case are an error.

### Text tags — `{{name}}`
- Written anywhere in slide text, including table cells and grouped shapes. Spaces inside the
  braces are allowed: `{{ oee }}`.
- A tag may be split across runs by PowerPoint (spell-check, partial formatting); the engine
  rejoins it. The replacement keeps the formatting of the run where `{{` starts.
- Suffixes on a tag pick a derived value: `{{oee.previous}}`, *reserved:* `{{oee.trend}}` (▲ ▼ ▶),
  `{{oee.change}}`, `{{oee.change_pct}}`.
- Not scanned in L1: slide masters/layouts, speaker notes, chart titles, SmartArt.

### Tables and charts — named shapes
A table or chart is a slot when its **Selection Pane name** (`p:cNvPr/@name`) is a valid slot
name. PowerPoint's default names ("Table 3", "Chart 5") contain a space, so untouched shapes are
never treated as slots by accident. A named shape that is neither a table nor a chart is ignored.

### Workbook layout (Lite)
| Sheet | Shape | Feeds |
| --- | --- | --- |
| `kpis` | columns `key, value, previous, format` — one row per text tag | `{{key}}`, `{{key.previous}}` |
| `meta` (optional) | columns `key, value` — free text | `{{key}}` |
| one per table slot, named like the slot | row 1 = headers, rows 2+ = body | the table |
| one per chart slot, named like the slot | row 1 = series names (A1 is the category label), column A = categories, other columns = series | the chart |

- `kpis` and `meta` share one namespace; the same key in both is an error.
- Header matching for tables is by header text (trimmed, case-insensitive), so column order in
  the sheet doesn't matter and extra sheet columns are ignored. A template column with no sheet
  column is an error.
- *Reserved (L2):* an optional `batch` column on any sheet splits rows into one deck per value.

### Formats (`format` column)
`text`, `number[:decimals]`, `percent[:decimals]`, `currency[:decimals]`, `date`. Empty means
numbers print with up to 2 decimals and grouping, and everything else prints as-is. `percent`
expects a fraction (0.853 → 85.3%), matching how Excel stores percentages. Locale comes from
the run options (default `en-US`); currency defaults to USD (configurable per run).

### Table fill rules
- Row 1 of the template table is the header and is kept.
- Body rows: the last template body row is the style row. Extra data rows clone it; surplus
  template rows are removed. The frame height is recomputed from the row heights.
- Merged cells in body rows are unsupported in L1 (checker error).

### Chart fill rules (L1: bar/column, line)
- Series count in the sheet must be ≤ the template's; extra template series are removed.
  More series than the template is an error in L1 (cloning series styles is L2).
- Categories are written as strings (`c:strRef`); values as numbers with the template series'
  number format kept. Empty cells become gaps.
- The series caches and range formulas are rewritten, and the embedded workbook is **replaced**
  by a freshly generated one holding exactly the plotted data, so "Edit Data" in PowerPoint
  shows the new numbers.

### Missing data
Engine option `onMissing`: `error` (default for the CLI), `keep` (leave `{{tag}}` visible), or `blank`.

## Alternatives rejected
- A prefix convention for shapes (`table:downtime`): uglier and easy to typo; the space rule
  already separates default names from slot names.
- `meta` as a single row with one column per key: inconsistent with `kpis` and breaks with `batch`.
- Editing the template's embedded workbook in place: ExcelJS round-trips PowerPoint's embedded
  table definitions poorly; a fresh workbook is smaller and predictable. The embedded sheet's
  styling is lost, which no one sees outside "Edit Data".
- Positional table columns: fragile when someone reorders an export.
