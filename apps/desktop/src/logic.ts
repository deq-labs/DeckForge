/** UI logic without React or Electron, so it can be unit tested. */
import type { CheckView, JobOptions, ProblemView, SlotView } from '../electron/api.ts';

export const DEFAULT_OPTIONS: JobOptions = { locale: 'en-US', currency: 'USD', missing: 'error' };

export const LOCALES: { code: string; label: string }[] = [
  { code: 'en-US', label: 'English (US): 1,234.5 · Oct 31, 2026' },
  { code: 'en-GB', label: 'English (UK): 1,234.5 · 31 Oct 2026' },
  { code: 'de-DE', label: 'Deutsch: 1.234,5 · 31.10.2026' },
  { code: 'fr-FR', label: 'Français: 1 234,5 · 31 oct. 2026' },
  { code: 'es-ES', label: 'Español: 1234,5 · 31 oct 2026' },
  { code: 'nl-NL', label: 'Nederlands: 1.234,5 · 31 okt 2026' },
  { code: 'pt-BR', label: 'Português (BR): 1.234,5 · 31 de out. de 2026' },
];

/** Splits a path on either separator; the renderer has no access to node:path. */
export function splitPath(p: string): { dir: string; file: string; sep: string } {
  const i = Math.max(p.lastIndexOf('/'), p.lastIndexOf('\\'));
  const sep = p.includes('\\') ? '\\' : '/';
  return { dir: i >= 0 ? p.slice(0, i) : '', file: i >= 0 ? p.slice(i + 1) : p, sep };
}

export function stem(p: string): string {
  return splitPath(p).file.replace(/\.[^.]+$/, '');
}

/** Default output folder: "<template>-decks" next to the template. */
export function defaultOutDir(templatePath: string): string {
  const { dir, sep } = splitPath(templatePath);
  return `${dir}${dir ? sep : ''}${stem(templatePath)}-decks`;
}

export type SlotStatus = 'pending' | 'ok' | 'problem' | 'unsupported';

/** Whether each slot is ready, given the latest check (or none yet). */
export function slotStatus(slot: SlotView, check: CheckView | null): SlotStatus {
  if (!slot.fillable) return 'unsupported';
  if (!check) return 'pending';
  const name = slot.name.toLowerCase();
  const hit = check.problems.some((p) => p.severity === 'error' && p.slot?.toLowerCase() === name);
  return hit ? 'problem' : 'ok';
}

export function countBySeverity(problems: ProblemView[]): { errors: number; warnings: number } {
  const errors = problems.filter((p) => p.severity === 'error').length;
  return { errors, warnings: problems.length - errors };
}

export function plural(n: number, word: string): string {
  return `${n} ${word}${n === 1 ? '' : 's'}`;
}

export function newId(): string {
  return `r-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

export function relativeTime(iso: string | undefined, now = Date.now()): string {
  if (!iso) return 'never run';
  const mins = Math.round((now - new Date(iso).getTime()) / 60000);
  if (mins < 1) return 'just now';
  if (mins < 60) return `${plural(mins, 'minute')} ago`;
  const hours = Math.round(mins / 60);
  if (hours < 24) return `${plural(hours, 'hour')} ago`;
  const days = Math.round(hours / 24);
  if (days < 45) return `${plural(days, 'day')} ago`;
  return new Date(iso).toLocaleDateString();
}
