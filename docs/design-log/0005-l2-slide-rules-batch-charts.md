# 0005 — L2: slide rules, batch decks, more chart types

- Date: 2026-11-11 · Edition: lite (engine parts: shared) · Status: accepted

## Slide rules — speaker-notes directives
A line in a slide's speaker notes that starts with `#` and a known word is an instruction:

| Line | Effect |
| --- | --- |
| `#repeat sites` | One copy of the slide per row of the `sites` sheet. Tags on the slide are filled from that row's columns first, then from `kpis`/`meta`. Zero rows removes the slide. |
| `#show-if incidents` | Keep the slide only if `incidents` has a value that isn't empty, `0`, `no`, `false` or `none`. |
| `#hide-if no_exports` | Remove the slide if `no_exports` has such a value. |

- Rules are evaluated per copy, so `#show-if` on a repeated slide can test a column of the row.
- Directive lines are removed from the output's notes; the rest of the notes are kept and their
  tags filled. A line that looks like a directive but doesn't parse (`#repeat` with no sheet,
  `#showif x`) is a warning, so typos aren't silent. Other `#` lines (`#1 priority`) are notes.
- A repeat sheet's columns map to tags by header: lower-cased, spaces and punctuation become
  `_` ("Plant Manager" → `{{plant_manager}}`). A column named `<tag>_previous` feeds
  `{{tag.previous}}` and the trend fields.
- Template slide numbers are used in all messages, even after copies shift the output numbering.

## Slide duplication (engine)
Copying a slide copies its slide part, its notes slide, and each chart with that chart's
embedded workbook and style parts; layouts, images and media stay shared. The copy gets a new
`p:sldId` (and is added to the same PowerPoint section, if the deck has sections). After all
slide operations, parts no longer reachable from the package root are removed, so deleted
slides leave no orphans.

## More text locations
Tags are also filled in **speaker notes**, **chart and axis titles**, and in non-placeholder
shapes on **slide layouts and masters** (e.g. a `{{company}}` footer on the master). Tags inside
layout *placeholders* are prompt text PowerPoint never shows, so they produce a warning.
SmartArt is not supported; a tag inside SmartArt produces a warning.

## Charts
Supported: bar/column, line, area, pie, doughnut (2-D and 3-D variants of bar/line/area/pie),
and combos of those (e.g. columns + a line on a secondary axis). Series are matched by their
order in the chart; a combo needs at least enough series to keep every chart group non-empty.
Scatter, bubble, stock, radar and surface remain unsupported (numeric x values / special shapes).

## Batch decks
Any sheet may have a `batch` column. Each distinct non-blank value produces one deck containing
the rows with that value **plus** rows whose `batch` is blank (shared rows). For `kpis`/`meta`,
a batch-specific row overrides a shared row with the same key. Output files are named
`<template>-<batch>.pptx` by default, with characters Windows forbids replaced by `-`.
`deckforge fill` refuses a workbook with batches and points to `deckforge batch`.

## Checker
`deckforge check <template> [data]` lists template problems (unsupported shapes, bad tags,
misplaced tags, duplicate names) and, with a workbook, does a full dry run of every batch
without writing files. After filling, a table that now runs past the bottom of the slide is a
warning.
