import { styleText } from 'node:util';
import type { Issue } from '@deckforge/pptx-engine';
import type { Problem } from '@deckforge/jobs';

export interface Io {
  out: (line: string) => void;
  err: (line: string) => void;
}

export const consoleIo: Io = {
  out: (l) => process.stdout.write(`${l}\n`),
  err: (l) => process.stderr.write(`${l}\n`),
};

const color = (fmt: Parameters<typeof styleText>[0], s: string): string => styleText(fmt, s, { stream: process.stderr });

/** Prints errors first, then warnings, each with its fix-it hint if it has one. */
export function printProblems(problems: (Problem | Issue)[], io: Io): void {
  const errors = problems.filter((i) => i.severity === 'error');
  const warnings = problems.filter((i) => i.severity === 'warning');
  for (const i of [...errors, ...warnings]) {
    const label = i.severity === 'error' ? color('red', 'error') : color('yellow', 'warning');
    io.err(`${label}: ${i.message}`);
    const hint = (i as Problem).hint;
    if (hint) io.err(`       ${color('dim', hint)}`);
  }
}

export function plural(n: number, word: string): string {
  if (n === 1) return `1 ${word}`;
  return `${n} ${/(s|x|ch|sh)$/.test(word) ? `${word}es` : `${word}s`}`;
}
