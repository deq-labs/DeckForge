# DeckForge — Self-Hosted Report Deck Automation: Build Plan

Nov 11, 2026 · @Joe

## Summary

DeckForge (working name) fills a customer's own PowerPoint template with their data and produces native, editable decks: one deck, or one per plant, customer or region. It runs on the customer's own machine or server, so no data leaves their network.

It is sold in three editions, built and launched in this order:

| Edition | Who sets it up | Data in | Runs on | Sold as |
| --- | --- | --- | --- | --- |
| **Lite (build first)** | The buyer, alone, in under 30 minutes | Excel / CSV in a documented layout | Desktop app (Windows, Mac) + CLI | One-time license, self-serve checkout |
| **Pro (later)** | Buyer's IT, optionally with you | Databases, REST, industrial historians | Docker on their server | Annual or one-time license per server |
| **Services (alongside Lite)** | You | Whatever they have | Their setup | Fixed price per template, or hourly |

**Recommendation:** ship Lite, sell it from your own site, and take report-automation jobs on Upwork at the same time. Start Pro only after Lite has paying users asking for database connections or scheduling.

**What it is not:** not a BI tool, not an AI slide generator, and not a CMS. Templates are designed in PowerPoint itself.

## Target markets

Lite sells to anyone who rebuilds the same deck from a spreadsheet every week or month. Its template packs lead with manufacturing, where your background sells the product. Pro sells into regulated or isolated sites later.

| Edition | Buyer | Recurring decks | Why they buy |
| --- | --- | --- | --- |
| Lite | Operations / CI managers at small and mid-size manufacturers | Monthly ops review, OEE and downtime, quality, customer scorecards | Hours saved; native charts; no IT project needed |
| Lite | Consultants and small agencies | Client monthly reports, one deck per client | Batch: one Excel file → 20 client decks |
| Lite | Finance and ops at small companies, nonprofits, schools | Board packs, KPI updates | Cheaper and simpler than enterprise add-ins; works on Mac |
| Pro | Mid-size manufacturers with MES/historians | Same decks, straight from SQL / PI / Ignition, on a schedule | Data stays on site; no manual export |
| Pro | Defense contractors, utilities, credit unions | Program reviews, reliability reports, board packs | Air-gapped or regulated; cloud tools are off the table |

**Competitors for Lite:** PowerPoint add-ins such as PresentationPoint DataPoint (one-time, Windows), Deckary, think-cell and Macabacus (yearly, aimed at finance). Lite's angle: batch decks from one file, runs without PowerPoint open, Mac support, and ready-made operations template packs.

## Launch validation

For Lite, validation and early sales are the same activity. Put a landing page and a demo video up before Lite is finished, and measure real interest.

1. **Landing page + 2-minute demo** (Excel in → branded batch decks out), with an email signup and a discounted pre-order.
2. **Upwork / services:** take 2–3 report-automation jobs with the CLI. Each one tests a real template and becomes a template pack or a case study.
3. **Posts in places buyers already read:** r/manufacturing, r/PowerPoint, r/excel, LinkedIn operations groups, CI/Lean forums. Show before/after, not a sales pitch.
4. **Talk to 5–10 signups:** what decks, how often, what breaks.

**Go / no-go for finishing Lite (my starting targets, adjust as you learn):**

- [ ] 100+ email signups or 5+ paid pre-orders within 6 weeks of the page going up
- [ ] 2+ services jobs completed with the engine
- [ ] At least 3 people independently describe the same recurring deck

If the targets miss, keep selling services with the CLI and change the template-pack niche before building more of the app.

## Product scope by edition

Lite is the shared engine plus a desktop app. Pro adds a server, connectors and scheduling on top of the same engine. Agents build only the Lite column until Lite launches.

| Capability | Lite | Pro | Services |
| --- | --- | --- | --- |
| Template tags `{{name}}`, named tables and charts | Yes | Yes | You tag their template |
| Text, KPIs, trend arrows | Yes | Yes |  |
| Tables (fill, add/trim rows) | Yes | Yes |  |
| Native charts + embedded workbook | Bar, line, pie, combo | All Lite + scatter, waterfall |  |
| Repeat a slide per row; skip slides by rule | Yes | Yes |  |
| Batch: one deck per value of a column (plant, client) | Yes | Yes |  |
| Data input | Excel / CSV in the documented layout | + SQL Server, Postgres, MySQL, REST, Ignition, PI, OPC UA | You write their queries or reshape their exports |
| Output | PPTX; PDF if PowerPoint or LibreOffice is installed | + server-side PDF, email, network share |  |
| Preview | Slot checklist + rendered preview when LibreOffice is present | Exact PNG preview |  |
| Scheduling | Run from CLI with the OS scheduler | Built-in schedules | You set them up |
| Users and audit | Single user, local run log | Accounts, roles, audit log, LDAP/SSO later |  |
| Template packs | Ops review, KPI scorecard, client monthly report | Same | Custom packs per client |
| Template checker | Reports unsupported parts before a run | Yes |  |

**Lite's data layout (the contract that makes self-setup work):** one Excel workbook, one sheet per slot type.

- `kpis` sheet: columns `key, value, previous, format` (one row per KPI tag)
- One sheet per table tag: header row = the template table's headers
- One sheet per chart tag: first column = categories, other columns = series
- Optional `batch` column on any sheet: rows are split into one deck per distinct value
- Optional `meta` sheet: `period_start, period_end, title` and other free text tags

The app generates a blank workbook in this layout from any tagged template ("Create data file"). The buyer fills it in or points their existing export at it.

## Architecture and stack

One TypeScript monorepo. The engine and the Excel reader are shared packages, Lite is an Electron app around them, and Pro is a server app that reuses the HubKit TS platform code. Lite must never depend on server packages.

| Part | Edition | Tech | Role |
| --- | --- | --- | --- |
| `pptx-engine` | All | TypeScript, JSZip, an XML DOM library, ExcelJS | Read, discover slots, fill, write OOXML; no I/O |
| `data-xlsx` | All | ExcelJS, csv-parse | Read the documented layout; validate it against a template's slots |
| `cli` | Lite + Services | Node, commander | `deckforge fill`, `deckforge batch`, `deckforge init-data` |
| `desktop` | Lite | Electron + React 19 + Vite | Pick template and data, check, preview, generate; license screen |
| `licensing` | Lite + Pro | Ed25519 signed keys, verified offline | Edition, expiry of updates, seat count |
| `server` | Pro | NestJS, Postgres, BullMQ (from HubKit TS) | Accounts, connectors, schedules, audit |
| `connectors` | Pro | `mssql`, `pg`, `mysql2`, `undici`, `node-opcua` | Read-only data sources |
| `pdf` | Lite (optional) / Pro | Detect local PowerPoint or LibreOffice; LibreOffice container in Pro | PPTX → PDF and slide images |

```
deckforge/
├─ packages/
│  ├─ pptx-engine/   # shared core
│  ├─ data-xlsx/     # shared: Lite data layout
│  ├─ licensing/     # shared
│  └─ connectors/    # Pro only
├─ apps/
│  ├─ cli/           # Lite + Services
│  ├─ desktop/       # Lite
│  └─ server/ web/   # Pro (later)
├─ template-packs/   # ops-review, kpi-scorecard, client-monthly
├─ fixtures/templates/
└─ docs/design-log/
```

A lint rule (`eslint-plugin-boundaries`) stops `desktop`, `cli` and the shared packages from importing anything in `server`, `web` or `connectors`.

## PPTX engine design

The engine is the product. It's a pure TypeScript library that takes a template and a data object and returns a deck, written from the ECMA-376 spec and tested on decks you make yourself in PowerPoint.

**Pipeline:**

1. **Parse:** unzip, read `presentation.xml`, slide parts and their relationships, charts and embedded workbooks.
2. **Discover slots:** find `{{tags}}` in text runs (handling tags PowerPoint splits across runs), named tables and named charts. Output a manifest: slot id, type, slide, current sample values.
3. **Bind:** turn source data into the engine's `DeckData` object (values, tables, charts). In Lite, binding is by naming convention and lives in `data-xlsx` (see `docs/design-log/0003-lite-tagging-and-data-contract.md`); in Pro, a report definition maps each slot to a query. Formats (`number`, `percent`, `currency`, `date`, `trend`) are applied by the engine.
4. **Fill:** replace text while keeping the run's formatting; resize tables; rewrite chart series caches and the embedded `.xlsx` so the chart stays editable in PowerPoint.
5. **Slide ops:** delete slides, or duplicate a slide per data row (renumbering ids and relationships).
6. **Write and check:** re-zip, then validate that the file opens (an OOXML schema check plus a LibreOffice round trip in CI).

```ts
const tpl = await Template.load(buffer);
const manifest = tpl.slots();            // what can be filled
const deck = await tpl.fill(data, {      // data keyed by slot id
  locale: 'en-US', onMissing: 'placeholder',
});
await deck.toBuffer();
```

**Test strategy:** golden-file tests. For each fixture template, fill with fixed data and compare the extracted text and chart XML with a stored snapshot. Add a fuzz test with random data sizes (0 rows, 1 row, 500 rows), and open every output in LibreOffice in CI.

Keep this package free of any app code. It could later be sold or open-sourced separately as a developer library.

## Data connectors (Pro)

Lite reads only the Excel/CSV layout above. These connectors are Pro features; build them after Lite launches. Services jobs before then reshape a client's data into the Lite layout instead.

| Order | Connector | Library | Notes |
| --- | --- | --- | --- |
| 1 | CSV / Excel upload or watched folder | ExcelJS, csv-parse | Lowest-friction pilot path |
| 2 | SQL Server | `mssql` | Most common in manufacturing IT |
| 3 | PostgreSQL / MySQL | `pg`, `mysql2` |  |
| 4 | REST JSON | `undici` + JSONPath | ERP and MES APIs |
| 5 | Ignition (Inductive Automation) | Its SQL tables or web API | Popular with mid-size plants |
| 6 | OSIsoft / AVEVA PI | PI Web API | Common at larger sites |
| 7 | OPC UA | `node-opcua` | Direct from controllers; use sparingly, prefer historians |

Rules: connectors are read-only, credentials are encrypted at rest with a key the customer controls, every query has a timeout and row limit, and each run stores the exact data it used, so an old deck can be explained later.

## Packaging, licensing and pricing

Lite is a downloadable desktop app with a one-time license sold from your own site. Pro is a server bundle sold per site. Services are fixed-price jobs. The prices are starting guesses based on competing add-ins; test them at launch.

|  | Lite | Pro | Services |
| --- | --- | --- | --- |
| Delivery | Signed installers: Windows (.exe), Mac (.dmg, notarized); CLI via npm or a single binary | Docker Compose bundle + offline image tarball; Helm later | Scripts or a configured Lite/Pro install |
| License | One-time, perpetual for the version bought; 12 months of updates | Per server, yearly or one-time + maintenance | Per job |
| Price idea | $199 single user; $499 team (5 seats) | $1,500–$3,000 per server | $500–$1,500 per template; hourly for changes |
| Checkout | Lemon Squeezy or Paddle (handles tax, issues license keys) | Invoice through DEQ Labs LLC | Upwork, or invoice through DEQ Labs LLC |
| Support | Email + docs; template checker does most of the work | Email + a set number of support hours | Included in the job |

**Licensing (Lite and Pro):** a license key signed with Ed25519 and checked offline, carrying the edition, seat count and the date updates end. One optional online activation call at first use, with an offline activation file for air-gapped machines. After updates end, the app keeps working on the last version the buyer is entitled to.

**Launch extras for Lite:** a free trial that adds a small "Made with DeckForge" footer to output, a 30-day refund policy, and 3 free template packs.

## Build phases

Build the Lite track first, in phase order. The Services track starts as soon as the CLI works. Pro does not start until Lite has paying users.

### Lite track

#### L0 — Setup

- [x] Repo, strict TypeScript, lint boundaries, CI, CLAUDE.md, design log
- [ ] Five fixture templates you design in PowerPoint (ops review, KPI scorecard, table-heavy, chart-heavy, per-item repeat) — two synthetic stand-ins exist (`pnpm fixtures`); real PowerPoint-authored ones still needed
- [ ] DEQ Labs LLC formed; Lemon Squeezy or Paddle account opened

#### L1 — Engine core + CLI

- [x] Parse, slot discovery (including tags split across runs), text fill keeping formatting
- [x] Native bar and line charts, including the embedded workbook
- [x] Tables (fill, add/trim rows) — moved up from L2 (2026-11-11)
- [x] `data-xlsx` reads the documented layout; `deckforge init-data` writes a blank workbook from a template
- [ ] `deckforge fill` works end to end; golden-file tests; output opens cleanly in PowerPoint and LibreOffice — done except: verify in real PowerPoint (Open XML SDK validation passes; LibreOffice runs in CI)

#### L2 — Engine complete

- [x] Pie and combo charts, trend arrows, number formats (also doughnut, area and 3-D variants)
- [x] Repeat-per-row slides, skip rules, batch mode (one deck per `batch` value) — see design-log 0005
- [x] Template checker with plain-language errors ("Chart 'downtime' has 3 series in the template but 4 columns in the sheet")
- [x] A 30-slide deck in under 5 seconds; fuzz tests on data sizes

#### L3 — Desktop app

- [x] Electron shell: choose template, choose data, run checker, generate, open output folder
- [x] Slot checklist view; rendered preview when LibreOffice or PowerPoint is installed
- [x] Recent jobs saved as "recipes" (template + data path + options) for one-click monthly reruns
- [x] Trial watermark mode

#### L4 — Licensing and installers

- [x] Ed25519 license keys, offline check, activation file flow, seat count — design-log 0007
- [ ] Signed Windows installer and notarized Mac build; auto-update for licensed users — built and wired (electron-builder, release workflow, entitlement-aware updates); needs certificates, Apple account and an update host (docs/release.md)
- [x] Checkout webhook issues keys; a license-recovery page — `apps/license-server`; needs deploying (docs/licensing-operations.md)

#### L5 — Template packs and docs

- [x] Three packs (ops review, KPI scorecard, client monthly report), each with a sample workbook — `template-packs/`
- [ ] Docs: tagging guide, data layout, troubleshooting; a 5-minute getting-started video — docs in `docs/user/`; video script ready (`docs/user/video-script.md`), recording needed
- [ ] A new user goes from download to first deck in under 30 minutes (test with two people who aren't you)

#### L6 — Launch

- [ ] Landing page, demo video, pricing page, refund policy, terms and EULA
- [ ] Posts in 5 communities; 1 case study from a services job
- [ ] First 10 paid licenses

### Services track (starts after L1)

- [ ] Upwork profile and a fixed-price offer: "I'll automate your recurring PowerPoint report"
- [ ] First 3 jobs delivered with the CLI; each template generalized into a pack if reusable
- [ ] A one-page intake form: template, sample data, how often, who receives it

### Pro track (starts after 10+ Lite sales and repeated requests)

- [ ] P1: Server app from HubKit TS (accounts, roles, audit) in single-tenant mode
- [ ] P2: SQL Server, Postgres and Excel-folder connectors with a query test runner
- [ ] P3: Schedules, server-side PDF, email delivery, run history with data snapshots
- [ ] P4: Compose bundle, offline installer, backup/restore, upgrade test
- [ ] P5: Industrial connectors (Ignition, PI, OPC UA) driven by what buyers ask for

## Agentic workflow and risks

Use the HubKit TS setup: a planner, a builder, a reviewer with fresh context and a tester; one task per PR; you merge. Every task file names its edition (`lite`, `pro`, `shared`), so agents never build Pro features into Lite by accident.

**CLAUDE.md:**

```markdown
# DeckForge — agent rules
- Read docs/spec.md first. Current focus: LITE track only.
- Every task file states edition: lite | pro | shared. Do not build pro work
  unless the task says pro.
- Build from ECMA-376, public library docs and files in fixtures/.
- packages/pptx-engine and packages/data-xlsx: no I/O, no app imports,
  pure functions + tests.
- desktop, cli and shared packages must not import server, web or connectors
  (eslint-plugin-boundaries enforces this).
- Every engine change ships a golden-file test and must open in LibreOffice.
- Lite makes no network calls except license activation and update checks,
  both optional and documented.
- Error messages are written for non-developers: say what is wrong and how to
  fix it in the template or the workbook.
- Ask before adding a dependency. Log design decisions in docs/design-log/.
```

**Risks:**

| Risk | Edition | Likelihood | Mitigation |
| --- | --- | --- | --- |
| Nobody finds Lite | Lite | High | Landing page and demo before the app is done; services jobs as case studies; SEO pages for "automate monthly PowerPoint report from Excel" |
| PPTX edge cases (SmartArt, grouped shapes, odd templates) | All | High | Supported subset documented; template checker explains what can't be filled |
| Support load from self-serve buyers | Lite | Medium | Checker + docs + video first; template-setup help sold as a service |
| Code-signing and notarization costs and hassle | Lite | Medium | Budget for a Windows signing certificate and an Apple developer account before L4 |
| Piracy of license keys | Lite | Low–medium | Accept some; keep the price fair; updates need a valid key |
| Competitors (DataPoint, Deckary, think-cell) | Lite | Medium | Stay on batch decks, runs without PowerPoint, Mac, operations packs |
| Long sales cycles | Pro | High | Don't start Pro until Lite buyers ask for it |
| Time with a new baby and a full-time job | All | High | Phases sized for weeknights; services jobs only when there's capacity |
