/**
 * License state for the desktop app. Verification is offline (Ed25519, @deckforge/licensing).
 * The only network calls are the optional seat registration/release, made only when the user
 * asks and a license server is configured. See docs/design-log/0007-licensing.md.
 */
import { createHash } from 'node:crypto';
import { mkdir, readFile, rename, rm, writeFile } from 'node:fs/promises';
import { arch, hostname, platform, userInfo } from 'node:os';
import { join } from 'node:path';
import { activationRequestCode, evaluateLicense, normalizeKey, recoveryUrl, verifyActivation, verifyLicense } from '@deckforge/licensing';
import type { LicenseState } from '@deckforge/licensing';
import type { LicenseView } from '../api.ts';

export interface LicenseConfig {
  publicKey: string;
  /** This build's release date (YYYY-MM-DD); undefined in development. */
  releaseDate: string | undefined;
  /** License server base URL; '' disables every licensing network call. */
  serverUrl: string;
}

interface Stored {
  key?: string;
  activation?: string;
}

/** A stable, anonymous id for this computer + user: a hash, never the raw values. */
export function machineId(): string {
  let user = '';
  try {
    user = userInfo().username;
  } catch {
    // Some service accounts have no user info; the rest still identifies the machine.
  }
  return createHash('sha256').update(`deckforge|${hostname()}|${user}|${platform()}|${arch()}`).digest('hex').slice(0, 32);
}

type Fetch = (url: string, init: { method: string; headers: Record<string, string>; body: string }) => Promise<{ status: number; json(): Promise<unknown> }>;

export class LicenseService {
  private readonly file: string;
  private readonly config: LicenseConfig;
  private readonly machine: string;
  private readonly fetchImpl: Fetch;

  constructor(dir: string, config: LicenseConfig, machine = machineId(), fetchImpl: Fetch = fetch) {
    this.file = join(dir, 'license.json');
    this.config = config;
    this.machine = machine;
    this.fetchImpl = fetchImpl;
  }

  private async read(): Promise<Stored> {
    try {
      return JSON.parse(await readFile(this.file, 'utf8')) as Stored;
    } catch {
      return {};
    }
  }

  private async write(s: Stored): Promise<void> {
    await mkdir(join(this.file, '..'), { recursive: true });
    await writeFile(`${this.file}.tmp`, JSON.stringify(s, null, 2));
    await rename(`${this.file}.tmp`, this.file);
  }

  async state(): Promise<LicenseState> {
    const s = await this.read();
    return evaluateLicense({ key: s.key, activation: s.activation, publicKey: this.config.publicKey, releaseDate: this.config.releaseDate, machine: this.machine });
  }

  async status(problem?: string): Promise<LicenseView> {
    const st = await this.state();
    const base = this.config.serverUrl.replace(/\/$/, '');
    const recovery = recoveryUrl(base);
    const view: LicenseView = {
      mode: st.mode,
      summary:
        st.mode === 'licensed'
          ? `Licensed to ${st.license!.licensee}.`
          : 'Trial: generated decks carry a small "Made with DeckForge" footer.',
      activated: st.activated,
      canActivateOnline: Boolean(base),
      ...(base ? { offlineActivationUrl: `${base}/activate` } : {}),
      ...(recovery ? { recoveryUrl: recovery } : {}),
    };
    if (st.license) {
      view.licensee = st.license.licensee;
      view.seats = st.license.seats;
      view.updatesUntil = st.license.updatesUntil;
    }
    const p = problem ?? st.problem;
    if (p) view.problem = p;
    return view;
  }

  /** Accepts a license key or an activation file. */
  async activate(text: string, opts: { online?: boolean } = {}): Promise<LicenseView> {
    const token = normalizeKey(text);
    const stored = await this.read();

    if (token.startsWith('DFA1.')) {
      const st = await this.state();
      if (!st.license) return this.status('Activate your license key first, then load the activation file.');
      const v = await verifyActivation(token, this.config.publicKey, { licenseId: st.license.id, machine: this.machine });
      if (!v.ok) return this.status(v.reason);
      await this.write({ ...stored, activation: token });
      return this.status();
    }

    const v = await verifyLicense(token, this.config.publicKey);
    if (!v.ok) return this.status(this.config.publicKey ? v.reason : 'This build of DeckForge has no license key configured, so it can only run in trial mode.');

    let activation: string | undefined;
    let note: string | undefined;
    if (opts.online && this.config.serverUrl) {
      try {
        const res = await this.fetchImpl(`${this.config.serverUrl.replace(/\/$/, '')}/api/activate`, {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ key: token, machine: this.machine }),
        });
        const body = (await res.json()) as { activation?: string; error?: string };
        if (res.status === 200 && body.activation) activation = body.activation;
        // Seat limit reached or license refunded: the server's answer is final.
        else if (res.status === 409 || res.status === 403) return this.status(body.error ?? 'This license cannot be activated on another computer.');
        else note = body.error ?? `The license server answered ${res.status}.`;
      } catch {
        note = "Couldn't reach the license server. DeckForge works fully offline; you can register this computer later, or with an activation file.";
      }
    }
    await this.write({ key: token, ...(activation ? { activation } : {}) });
    return this.status(note);
  }

  async deactivate(): Promise<LicenseView> {
    const s = await this.read();
    if (s.key && s.activation && this.config.serverUrl) {
      // Best effort: free the seat. The local license is removed either way.
      await this.fetchImpl(`${this.config.serverUrl.replace(/\/$/, '')}/api/deactivate`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ key: s.key, machine: this.machine }),
      }).catch(() => undefined);
    }
    await rm(this.file, { force: true });
    return this.status();
  }

  async requestCode(): Promise<string> {
    const st = await this.state();
    if (!st.license) throw new Error('Activate your license key first; the request code is for that license.');
    return activationRequestCode(st.license.id, this.machine);
  }
}
