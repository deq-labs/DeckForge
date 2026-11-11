# Tagging guide

You prepare a template once, in PowerPoint, using your own design. DeckForge only replaces what
you tag.

## Text: `{{tags}}`

Type a tag anywhere in slide text: titles, text boxes, table cells, grouped shapes, speaker notes,
chart titles. `{{ oee }}` with spaces works too.

- **Names** use letters, digits, `_` and `-`, start with a letter, and have no spaces:
  `oee`, `scrap_rate`, `site-name`. Case doesn't matter (`{{OEE}}` = `{{oee}}`).
- **Formatting:** the value takes the font, size and color of the tag's first characters (`{{`).
  Partially formatted or spell-checked tags are fine — DeckForge rejoins them.
- **Comparisons** use the `previous` column of the `kpis` sheet:

| Tag | Shows | Example |
| --- | --- | --- |
| `{{oee}}` | the value | 78.4% |
| `{{oee.previous}}` | the previous value | 75.1% |
| `{{oee.change}}` | value − previous, with sign | +3.3% |
| `{{oee.change_pct}}` | relative change | +4.4% |
| `{{oee.trend}}` | ▲ up, ▼ down, ▶ same | ▲ |

- **Multi-line text:** line breaks in the Excel cell (Alt+Enter) become line breaks on the slide.
- **On every slide:** put a tag in a normal text box on the slide master or a layout (e.g. a
  `{{company}}` footer) and it's filled everywhere. (Tags inside layout *placeholders* are only
  prompt text in PowerPoint and never appear on slides — DeckForge warns about those.)

## Tables

1. Make the table in PowerPoint with **one header row** and **at least one body row** styled the
   way you want (that row is copied for extra data).
2. Name it in the Selection Pane, e.g. `downtime`.
3. In the workbook, the `downtime` sheet's header row matches the table's headers (order doesn't
   matter; extra columns are ignored). Rows below fill the table, which grows or shrinks to fit.

Not supported: merged cells below the header row (DeckForge can't add rows to those).

## Charts

Bar/column, line, area, pie, doughnut (also 3-D versions), and combinations such as columns
with a line on a second axis.

1. Make the chart in PowerPoint with the series you want (series colors and styles are kept).
2. Name it in the Selection Pane, e.g. `oee_trend`.
3. In the workbook, the `oee_trend` sheet has categories in column A and one series per column
   from B, with series names in row 1. Any number of categories works.

The data can have fewer series than the template chart (extra ones are removed), not more.
Not supported: scatter, bubble, stock, radar and surface charts; SmartArt.

## Slide rules (in the speaker notes)

| Notes line | Effect |
| --- | --- |
| `#repeat sites` | One copy of the slide per row of the `sites` sheet. Tags on that slide are filled from the row's columns (header "Plant Manager" → `{{plant_manager}}`). |
| `#show-if incidents` | Keep the slide only if `incidents` has a value that isn't empty, 0, "no", "false" or "none". |
| `#hide-if skip_appendix` | Remove the slide if `skip_appendix` has such a value. |

These lines are removed from the finished deck's notes; your other notes stay.

## Check your template

Choose the template in the app (or run `deckforge check template.pptx`). It lists every tag,
table and chart it found and warns about anything it can't fill, with what to change.
