# Troubleshooting

DeckForge tells you what's wrong and where. Most messages also say how to fix it; this page adds
detail for the common ones. Nothing is written until every problem is fixed, so a failed run
never leaves half a set of decks.

## "No data for {{tag}} (slide N)"

The template has the tag but the workbook doesn't.

- Add a row with that **key** to the `kpis` sheet — or, if the row is there, fill in its **value**
  (empty counts as missing).
- On a repeated slide, add a column with that name to the slide's sheet.
- To generate anyway, set **Missing data** to "Leave {{tags}} visible" or "Leave blank"
  (`--missing keep|blank` on the command line).

## "Table 'x' on slide N has the column "Minutes", but the data has no column with that header"

The sheet's row 1 must contain each of the table's header texts (any order; capitals and extra
spaces don't matter). Rename the sheet column — or the table header in PowerPoint.

## "Chart 'x' … has 2 series in the template but the data has 3"

The sheet has more series columns than the chart. Either remove a column (series are the
columns from B up to the first empty header), or add a series in PowerPoint
(*Chart Design → Edit Data*) so the chart has a style for it.

## "… is a scatter chart" / "contains a … chart part"

That chart type can't be filled. Change it in PowerPoint (*Chart Design → Change Chart Type*) to
bar/column, line, area, pie or doughnut, or leave it as a static picture.

## "has merged cells below the header row"

Unmerge the body cells of that table in PowerPoint. Merged header cells are fine.

## "Sheet 'x' cell C5 should be a number but contains "n/a""

Chart values must be numbers. Leave the cell empty for a gap.

## "has a formula with no saved result"

The workbook was made by another program and never calculated. Open it in Excel and save.

## ""x" appears twice"

Each key may appear once per deck (once per batch value when using batches). Delete the extra row.

## "looks like a tag but will not be filled"

Something like `{{bad name}}` or `{{ }}` is in the template. Tag names can't contain spaces;
use `_` instead.

## "Sheet 'x' does not match any table, chart or repeated slide"

Just a note: the sheet is ignored. Rename it to match a table/chart, or start its name with `_`
to silence the note.

## "Table 'x' … now runs past the bottom of the slide"

More rows than fit. Make the template's row height or font smaller, or send fewer rows.

## The deck opens, but numbers look wrong

- Percent showing 0.78 instead of 78%: format the Excel cell as a percentage, or put `percent`
  in the `format` column.
- Dates showing as numbers: format the Excel cell as a date.
- Decimal commas/points: choose your style under **Numbers and dates**.

## License

- "This license key is not valid": copy the whole key from the email (it starts with `DFL1.`), or
  use **Lost your key?** to have it sent again.
- "Your license includes versions released until …": this version came out after your free
  updates ended. Keep using the version you have, or renew updates.
- "already active on N computers": remove it from an old computer (License → Remove from this
  computer) or contact support.

## Still stuck?

Run **Check** and send the full message list, plus the template, to support. Template setup help
is also available as a service.
