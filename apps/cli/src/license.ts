/**
 * CLI license handling. Same keys and rules as the desktop app (offline Ed25519 verification);
 * the CLI never makes a network call. Unlicensed runs add the trial footer to every slide.
 *
 * The key is read from DECKFORGE_LICENSE_KEY (handy for scheduled jobs) or from the file written
 * by `deckforge license activate`.
 */
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { PRODUCTION_PUBLIC_KEY, evaluateLicense, normalizeKey, verifyLicense } from '@deckforge/licensing';
import type { LicenseState } from '@deckforge/licensing';

export const WATERMARK = 'Made with DeckForge';

/** Set by the bundler (scripts/build.mjs); absent when running from source. */
declare const __DECKFORGE_RELEASE_DATE__: string | undefined;

/** Running the TypeScript source (development), not a built bundle. */
const fromSource = import.meta.url.endsWith('.ts');

function publicKey(): string {
  // Overriding the key is for development and tests only; built copies always use the product key.
  return (fromSource && process.env.DECKFORGE_LICENSE_PUBLIC_KEY) || PRODUCTION_PUBLIC_KEY;
}

function releaseDate(): string | undefined {
  return typeof __DECKFORGE_RELEASE_DATE__ === 'string' ? __DECKFORGE_RELEASE_DATE__ : undefined;
}

export function configDir(env: NodeJS.ProcessEnv = process.env, platform = process.platform, home = homedir()): string {
  if (env.DECKFORGE_CONFIG_DIR) return env.DECKFORGE_CONFIG_DIR;
  if (platform === 'win32') return join(env.APPDATA ?? join(home, 'AppData', 'Roaming'), 'DeckForge');
  if (platform === 'darwin') return join(home, 'Library', 'Application Support', 'DeckForge');
  return join(env.XDG_CONFIG_HOME ?? join(home, '.config'), 'deckforge');
}

const keyFile = (): string => join(configDir(), 'cli-license.key');

function storedKey(): string | undefined {
  if (process.env.DECKFORGE_LICENSE_KEY) return process.env.DECKFORGE_LICENSE_KEY;
  try {
    return readFileSync(keyFile(), 'utf8');
  } catch {
    return undefined;
  }
}

export async function licenseState(): Promise<LicenseState> {
  return evaluateLicense({ key: storedKey(), publicKey: publicKey(), releaseDate: releaseDate() });
}

/** Saves a key after checking it. Returns the problem, if any. */
export async function saveKey(text: string): Promise<{ ok: true; state: LicenseState } | { ok: false; reason: string }> {
  const key = normalizeKey(text);
  const v = await verifyLicense(key, publicKey());
  if (!v.ok) return { ok: false, reason: publicKey() ? v.reason : 'This build of DeckForge has no license key configured, so it can only run in trial mode.' };
  mkdirSync(configDir(), { recursive: true });
  writeFileSync(keyFile(), key, { mode: 0o600 });
  return { ok: true, state: await licenseState() };
}

export function removeKey(): void {
  rmSync(keyFile(), { force: true });
}

export function keyLocation(): string {
  return process.env.DECKFORGE_LICENSE_KEY ? 'DECKFORGE_LICENSE_KEY environment variable' : keyFile();
}
