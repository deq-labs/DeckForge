# DeckForge — agent rules

- Read `docs/spec.md` first. Current focus: **LITE track only** (L1: engine core + CLI).
- Read `docs/design-log/` before changing engine behaviour; log new design decisions there
  (next number, one decision per file, include *why* and alternatives rejected).
- Every task states edition: `lite` | `pro` | `shared`. Do not build pro work unless the task says pro.
- Build from ECMA-376, public library docs and files in `fixtures/`.
- `packages/pptx-engine` and `packages/data-xlsx`: no I/O, no app imports, pure functions + tests.
  Buffers (`Uint8Array`) in, buffers out. ESLint enforces this.
- desktop, cli and shared packages must not import server, web or connectors (ESLint enforces this).
- Every engine change ships a golden-file test, and output must open in LibreOffice (CI checks)
  and pass the Open XML SDK validator (`tools/ooxml-validate`).
- Lite makes no network calls except license activation and update checks, both optional and documented.
- Error messages are written for non-developers: say what is wrong and how to fix it in the
  template or the workbook. Name the slide number and the tag/shape name.
- Dependencies may be added when needed; note why in the PR description.

## Workflow

- Package manager is **pnpm** (workspace). Never use npm/yarn commands to install.
  - Add a dep to one package: `pnpm --filter @deckforge/pptx-engine add <pkg>`
  - Add a root dev tool: `pnpm add -Dw <pkg>`
  - Link a workspace package: `pnpm --filter @deckforge/cli add "@deckforge/pptx-engine@workspace:*"`
- `pnpm check` = typecheck + lint + test. Run it before committing.
- `pnpm fixtures` regenerates synthetic templates in `fixtures/templates/`.
- `pnpm deckforge <command>` runs the CLI from source.
- Update golden snapshots only deliberately: `pnpm --filter @deckforge/pptx-engine test -u`, then review the diff.

## Code conventions

- TypeScript source is run directly (Node >= 24 type stripping, no build step). So:
  - relative imports use the `.ts` extension;
  - erasable syntax only: no `enum`, `namespace`, or constructor parameter properties.
- Workspace packages export `./src/index.ts`; keep the public API in `index.ts` small.
- One branch + PR per task; the human merges.
