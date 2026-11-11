# 0002 — Enforce boundaries with `no-restricted-imports`

- Date: 2026-11-11 · Edition: shared · Status: accepted

## Context
The spec names `eslint-plugin-boundaries` to keep Lite from importing Pro code. Packages here
import each other by package name (`@deckforge/...`), which pnpm resolves through symlinks;
boundaries classifies files by path and needs an import resolver to follow those symlinks.

## Decision
Use ESLint's built-in `no-restricted-imports` (see `eslint.config.js`):
- Lite code (`packages/pptx-engine`, `data-xlsx`, `licensing`, `apps/cli`, `apps/desktop`, `tools`)
  may not import `@deckforge/server|web|connectors` or paths under those folders.
- `pptx-engine/src` and `data-xlsx/src` additionally may not import Node I/O modules
  (`fs`, `child_process`, `net`, `http`, ...), app packages, or use `fetch`/`process` globals.

A second guard is structural: pnpm only links a workspace package that is declared in
`package.json`, so an undeclared Pro dependency would not resolve at all.

## Alternatives rejected
- `eslint-plugin-boundaries` + `eslint-import-resolver-typescript`: same result, two more
  dependencies and a resolver config to maintain. Revisit if the rules get more granular than package level.
