# Getting started with DeckForge

DeckForge fills **your own PowerPoint template** with numbers from an **Excel workbook** and
saves a normal, editable deck — or one deck per site or client. Everything runs on your computer;
nothing is uploaded.

Plan on 20–30 minutes the first time. After that, a monthly deck is one click.

## 1. Install (2 minutes)

Download DeckForge for Windows or Mac and run the installer. It starts in **trial mode**: everything
works, and each slide gets a small "Made with DeckForge" note in the corner until you enter a
license key (top-right corner → **Trial** → paste your key).

## 2. Try a template pack (5 minutes)

1. Open DeckForge and click **Choose template…**. Pick `template.pptx` from the *ops-review* pack.
   DeckForge lists everything it can fill: `{{tags}}`, tables and charts.
2. Click **Choose workbook…** and pick the pack's `sample-data.xlsx`.
3. Click **Generate decks**, then **Open**. That's a filled deck — charts included, and still
   editable in PowerPoint (right-click a chart → *Edit Data* shows the new numbers).

## 3. Tag your own deck (10–15 minutes)

Open your usual monthly deck in PowerPoint and save a copy as your template.

- **Numbers and text:** replace a value with a tag, e.g. type `{{oee}}` where it says `78.4%`.
  Keep the formatting you like — the value takes the formatting of the tag.
- **Tables and charts:** select the table or chart, open **Home → Arrange → Selection Pane**, and
  rename it to a short name without spaces, e.g. `downtime`. Its sheet in the workbook gets the same name.

The [tagging guide](tagging-guide.md) has every option (trend arrows, repeating slides, hiding slides).

## 4. Make the data file (5 minutes)

In DeckForge, choose your template and click **Create data file…**. Excel opens a workbook laid
out for your template: a `kpis` sheet with one row per tag, and one sheet per table and chart,
already holding the numbers from your template. Replace them with this month's numbers and save.
The [data layout guide](data-layout.md) explains every sheet.

## 5. Check, generate, save as a recipe

- **Check** tells you, in plain words, anything that doesn't match ("No data for {{oee}} (slide 2)
  — Add a row with key "oee" to the kpis sheet").
- **Generate decks** writes the deck(s) into the output folder.
- **Save as recipe** remembers the template, workbook and options. Next month: update the
  workbook, click **Run** on the recipe. Done.

Stuck? See [troubleshooting](troubleshooting.md).

## Command line (optional)

The same engine is available as a command for scheduled jobs:

```sh
deckforge init-data template.pptx            # create the workbook
deckforge check template.pptx data.xlsx      # dry run, explains problems
deckforge fill template.pptx data.xlsx -o october.pptx
deckforge batch template.pptx data.xlsx      # one deck per value of the batch column
deckforge license activate <key>
```
