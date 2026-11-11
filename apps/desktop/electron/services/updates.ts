/**
 * Auto-update (installed builds only). Off by default: DeckForge makes no network call unless
 * the user turns on automatic checks or clicks "Check for updates".
 *
 * Entitlement: a licensed user only gets updates released while their updates are included
 * (license `updatesUntil`); newer releases are announced but not downloaded. Trial users can
 * always update.
 */
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import type { UpdateView } from '../api.ts';

/** The subset of electron-updater's AppUpdater we use (injected so this is testable). */
export interface Updater {
  autoDownload: boolean;
  autoInstallOnAppQuit: boolean;
  checkForUpdates(): Promise<{ updateInfo: { version: string; releaseDate: string } } | null>;
  downloadUpdate(): Promise<unknown>;
  quitAndInstall(): void;
}

/** Whether an update released on `releaseDate` (ISO) is covered by a license's `updatesUntil`. */
export function updateCovered(releaseDate: string, updatesUntil: string | undefined): boolean {
  if (!updatesUntil) return true; // trial: no entitlement limit
  return releaseDate.slice(0, 10) <= updatesUntil;
}

export class UpdateService {
  private view: UpdateView;
  private readonly settingsFile: string;
  private readonly updater: Updater | undefined;
  private readonly updatesUntil: () => Promise<string | undefined>;

  /** `updater` is undefined in development (unpackaged), which disables updates. */
  constructor(dir: string, updater: Updater | undefined, updatesUntil: () => Promise<string | undefined>) {
    this.settingsFile = join(dir, 'settings.json');
    this.updater = updater;
    this.updatesUntil = updatesUntil;
    this.view = updater ? { state: 'idle', auto: false } : { state: 'disabled', auto: false, message: 'Updates are available in installed copies of DeckForge.' };
    if (updater) {
      updater.autoDownload = false;
      updater.autoInstallOnAppQuit = true;
    }
  }

  async init(): Promise<void> {
    try {
      const s = JSON.parse(await readFile(this.settingsFile, 'utf8')) as { autoUpdate?: boolean };
      this.view.auto = s.autoUpdate === true;
    } catch {
      this.view.auto = false;
    }
    if (this.view.auto && this.updater) void this.check();
  }

  status(): UpdateView {
    return { ...this.view };
  }

  async setAuto(auto: boolean): Promise<UpdateView> {
    this.view.auto = auto;
    let settings: Record<string, unknown> = {};
    try {
      settings = JSON.parse(await readFile(this.settingsFile, 'utf8')) as Record<string, unknown>;
    } catch {
      // first save
    }
    await writeFile(this.settingsFile, JSON.stringify({ ...settings, autoUpdate: auto }, null, 2));
    return this.status();
  }

  async check(): Promise<UpdateView> {
    if (!this.updater) return this.status();
    if (this.view.state === 'checking' || this.view.state === 'downloading') return this.status();
    this.view = { state: 'checking', auto: this.view.auto };
    try {
      const r = await this.updater.checkForUpdates();
      const info = r?.updateInfo;
      if (!info) {
        this.view = { state: 'none', auto: this.view.auto, message: 'DeckForge is up to date.' };
        return this.status();
      }
      const until = await this.updatesUntil();
      if (!updateCovered(info.releaseDate, until)) {
        this.view = {
          state: 'not-covered',
          auto: this.view.auto,
          version: info.version,
          message: `Version ${info.version} is available, but your free updates ended on ${until}. Your current version keeps working; renew updates to get this one.`,
        };
        return this.status();
      }
      this.view = { state: 'downloading', auto: this.view.auto, version: info.version, message: `Downloading version ${info.version}…` };
      await this.updater.downloadUpdate();
      this.view = { state: 'ready', auto: this.view.auto, version: info.version, message: `Version ${info.version} is ready. Restart DeckForge to use it.` };
    } catch (e) {
      this.view = { state: 'error', auto: this.view.auto, message: `Couldn't check for updates: ${(e as Error).message}` };
    }
    return this.status();
  }

  install(): void {
    if (this.view.state === 'ready') this.updater?.quitAndInstall();
  }
}
