import { styleText } from 'node:util';
import type { Issue } from '@deckforge/pptx-engine';

export interface Io {
  out: (line: string) => void;
  err: (line: string) => void;
}

export const consoleIo: Io = {
  out: (l) => process.stdout.write(`${l}\n`),
  err: (l) => process.stderr.write(`${l}\n`),
};

const color = (fmt: Parameters<typeof styleText>[0], s: string): string => styleText(fmt, s, { stream: process.stderr });

export interface HintContext {
  /** Keys whose kpis/meta row exists but has an empty value. */
  blankKeys?: string[];
}

/** Lite-specific "how to fix it" hints, keyed by the engine's issue codes. */
function hint(issue: Issue, ctx: HintContext): string | undefined {
  const slot = issue.slot;
  switch (issue.code) {
    case 'missing-value':
      if (slot && ctx.blankKeys?.includes(slot.toLowerCase())) return `Fill in the value for "${slot}" in the kpis sheet (its value cell is empty).`;
      return slot ? `Add a row with key "${slot}" to the kpis sheet (or the meta sheet).` : undefined;
    case 'missing-table':
      return slot ? `Add a sheet named "${slot}" with the table's header row in row 1.` : undefined;
    case 'missing-chart':
      return slot ? `Add a sheet named "${slot}": categories in column A, one series per column from B.` : undefined;
    default:
      return undefined;
  }
}

export function printIssues(issues: Issue[], io: Io, ctx: HintContext = {}): void {
  const errors = issues.filter((i) => i.severity === 'error');
  const warnings = issues.filter((i) => i.severity === 'warning');
  for (const i of [...errors, ...warnings]) {
    const label = i.severity === 'error' ? color('red', 'error') : color('yellow', 'warning');
    io.err(`${label}: ${i.message}`);
    const h = hint(i, ctx);
    if (h) io.err(`       ${color('dim', h)}`);
  }
}

export function plural(n: number, word: string): string {
  return `${n} ${word}${n === 1 ? '' : 's'}`;
}
