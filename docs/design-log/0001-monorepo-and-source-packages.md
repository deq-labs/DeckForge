# 0001 — pnpm monorepo, TypeScript run from source

- Date: 2026-11-11 · Edition: shared · Status: accepted

## Context
The spec calls for one TypeScript monorepo with shared packages (`pptx-engine`, `data-xlsx`)
used by the CLI now and the Electron app and Pro server later.

## Decision
- pnpm workspace: `packages/*` (shared libraries), `apps/*` (cli, later desktop/server), `tools/*` (dev-only).
- Workspace packages export `./src/index.ts` directly. No build step in development: Node >= 24
  strips types natively, and vitest handles TS itself. `tsc` is used only for type-checking (`noEmit`).
- This requires `.ts` import extensions and `erasableSyntaxOnly` (no enums, namespaces, parameter properties).
- TypeScript is pinned to 6.0.x: typescript-eslint does not yet support TypeScript 7 (the native port has no JS API).
- Distribution builds (single-file CLI, Electron bundle) will use a bundler (esbuild) at L3/L4; that is a packaging concern only.

## Alternatives rejected
- Per-package `tsc` builds to `dist/` with project references: more moving parts, stale-build bugs, slower loop.
- `tsx`/`ts-node` loaders: unnecessary now that Node runs TS natively.
