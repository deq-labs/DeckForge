# 0006 — Desktop app (Lite L3)

- Date: 2026-11-11 · Edition: lite · Status: accepted

## Decisions
- **Process split.** The renderer (React 19, Vite) is sandboxed with context isolation and no
  Node. It talks to the main process only through the typed API in `electron/api.ts`, exposed by
  the preload. All file access, dialogs and conversion happen in the main process.
- **One orchestration path.** The main process calls `@deckforge/jobs` (the same code as the
  CLI), so the app and the CLI can never disagree about what a workbook means.
- **Bundling.** Vite builds the renderer; esbuild bundles main/preload (which import the
  TypeScript workspace packages) into CommonJS. No electron-vite/forge: two small scripts are
  easier to reason about. Packaging (installers, signing) is L4.
- **Opening files.** `openPath`/`showInFolder` only accept paths the user picked or DeckForge
  wrote in this session (or that belong to saved recipes), so a compromised renderer can't ask
  the OS to open arbitrary files.
- **Writes are atomic and all-or-nothing.** A run writes nothing unless every batch filled; each
  deck is written to a temp file and renamed, so a deck open in PowerPoint fails cleanly
  ("close it and generate again") without leaving a half-written file.
- **Preview** converts to PDF with PowerPoint (exact; COM on Windows, AppleScript on Mac) or
  LibreOffice (headless), whichever is installed, and shows it in Chromium's built-in PDF viewer.
  No preview tool → the button is hidden; decks are unaffected.
- **Recipes** are a JSON file in the user's app-data folder, written atomically; a damaged file
  is set aside rather than blocking the app.
- **Trial mode** adds a small "Made with DeckForge" text box to every slide (engine option
  `watermark`). The box's name contains spaces, so it is never read as a slot. L4 turns it off for
  licensed users.
- **UI smoke test.** `DECKFORGE_SMOKE` (tests only) launches the real app offscreen, drives the IPC
  API end to end and saves a screenshot: `pnpm --filter @deckforge/desktop smoke`.
