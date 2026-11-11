import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { armor, generateKeyPair, signActivation, signLicense } from '@deckforge/licensing';
import type { KeyPair, LicensePayload } from '@deckforge/licensing';
import { LicenseService } from '../electron/services/license.ts';
import { UpdateService, updateCovered } from '../electron/services/updates.ts';
import type { Updater } from '../electron/services/updates.ts';

const payload: LicensePayload = { v: 1, id: 'LS-1', product: 'deckforge', edition: 'lite', licensee: 'Acme', seats: 2, issuedAt: '2026-11-11T00:00:00Z', updatesUntil: '2027-11-11' };

let dir: string;
let keys: KeyPair;
let key: string;
beforeEach(async () => {
  dir = mkdtempSync(join(tmpdir(), 'deckforge-license-'));
  keys = await generateKeyPair();
  key = await signLicense(payload, keys.privateKey);
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

type Call = { url: string; body: unknown };
function fakeFetch(status: number, body: unknown, calls: Call[] = [], fail = false) {
  return async (url: string, init: { body: string }) => {
    calls.push({ url, body: JSON.parse(init.body) });
    if (fail) throw new Error('offline');
    return { status, json: async () => body };
  };
}

describe('LicenseService', () => {
  it('starts in trial, activates offline from a key, and survives a restart', async () => {
    const cfg = { publicKey: keys.publicKey, releaseDate: '2027-01-01', serverUrl: '' };
    const svc = new LicenseService(dir, cfg, 'machine-1');
    expect(await svc.status()).toMatchObject({ mode: 'trial', activated: false, canActivateOnline: false });
    // online requested, but no server configured: no network call is possible
    const v = await svc.activate(`\n${armor(key)}\n`, { online: true });
    expect(v).toMatchObject({ mode: 'licensed', licensee: 'Acme', seats: 2, updatesUntil: '2027-11-11', activated: false });
    expect(await new LicenseService(dir, cfg, 'machine-1').status()).toMatchObject({ mode: 'licensed' });
  });

  it('registers the seat online when asked, and stops at the seat limit', async () => {
    const activation = await signActivation({ v: 1, licenseId: 'LS-1', machine: 'machine-1', activatedAt: 'now' }, keys.privateKey);
    const calls: Call[] = [];
    const ok = new LicenseService(dir, { publicKey: keys.publicKey, releaseDate: undefined, serverUrl: 'https://lic.test/' }, 'machine-1', fakeFetch(200, { activation }, calls));
    expect(await ok.activate(key, { online: true })).toMatchObject({ mode: 'licensed', activated: true, canActivateOnline: true, recoveryUrl: 'https://lic.test/recover' });
    expect(calls).toEqual([{ url: 'https://lic.test/api/activate', body: { key, machine: 'machine-1' } }]);

    const dir2 = mkdtempSync(join(tmpdir(), 'deckforge-license-'));
    const full = new LicenseService(dir2, { publicKey: keys.publicKey, releaseDate: undefined, serverUrl: 'https://lic.test' }, 'machine-9', fakeFetch(409, { error: 'already active on 2 computers' }));
    expect(await full.activate(key, { online: true })).toMatchObject({ mode: 'trial', problem: 'already active on 2 computers' });
    rmSync(dir2, { recursive: true, force: true });
  });

  it('still activates when the server is unreachable, and says so', async () => {
    const svc = new LicenseService(dir, { publicKey: keys.publicKey, releaseDate: undefined, serverUrl: 'https://lic.test' }, 'm', fakeFetch(0, {}, [], true));
    expect(await svc.activate(key, { online: true })).toMatchObject({ mode: 'licensed', activated: false, problem: expect.stringMatching(/reach the license server/) });
  });

  it('loads an activation file for an offline computer', async () => {
    const svc = new LicenseService(dir, { publicKey: keys.publicKey, releaseDate: undefined, serverUrl: '' }, 'offline-1');
    await svc.activate(key);
    expect(await svc.requestCode()).toMatch(/^DFR1\./);
    const file = armor(await signActivation({ v: 1, licenseId: 'LS-1', machine: 'offline-1', activatedAt: 'now' }, keys.privateKey), 'ACTIVATION');
    expect(await svc.activate(file)).toMatchObject({ mode: 'licensed', activated: true });
    const wrong = armor(await signActivation({ v: 1, licenseId: 'LS-1', machine: 'other', activatedAt: 'now' }, keys.privateKey), 'ACTIVATION');
    expect((await svc.activate(wrong)).problem).toMatch(/different computer/);
  });

  it('explains keys that are invalid or do not cover this version', async () => {
    const svc = new LicenseService(dir, { publicKey: keys.publicKey, releaseDate: '2028-03-01', serverUrl: '' }, 'm');
    expect((await svc.activate('garbage')).problem).toMatch(/not a DeckForge license key/);
    expect(await svc.activate(key)).toMatchObject({ mode: 'trial', licensee: 'Acme', problem: expect.stringMatching(/renew updates/) });
    expect(await svc.deactivate()).toMatchObject({ mode: 'trial' });
    expect((await svc.status()).licensee).toBeUndefined();
  });
});

describe('UpdateService', () => {
  function fakeUpdater(info: { version: string; releaseDate: string } | null): Updater & { downloaded: boolean } {
    return {
      autoDownload: true,
      autoInstallOnAppQuit: false,
      downloaded: false,
      async checkForUpdates() {
        return info ? { updateInfo: info } : null;
      },
      async downloadUpdate() {
        this.downloaded = true;
      },
      quitAndInstall() {},
    };
  }

  it('is off by default and disabled in development', async () => {
    const dev = new UpdateService(dir, undefined, async () => undefined);
    await dev.init();
    expect(dev.status()).toMatchObject({ state: 'disabled', auto: false });
    const u = fakeUpdater(null);
    const svc = new UpdateService(dir, u, async () => undefined);
    await svc.init();
    expect(svc.status()).toMatchObject({ state: 'idle', auto: false });
    expect(u.autoDownload).toBe(false);
  });

  it('downloads covered updates and announces uncovered ones without downloading', async () => {
    const covered = fakeUpdater({ version: '1.2.0', releaseDate: '2027-05-01T00:00:00.000Z' });
    expect(await new UpdateService(dir, covered, async () => '2027-11-11').check()).toMatchObject({ state: 'ready', version: '1.2.0' });
    expect(covered.downloaded).toBe(true);

    const late = fakeUpdater({ version: '2.0.0', releaseDate: '2028-01-05T00:00:00.000Z' });
    expect(await new UpdateService(dir, late, async () => '2027-11-11').check()).toMatchObject({ state: 'not-covered', message: expect.stringMatching(/free updates ended on 2027-11-11/) });
    expect(late.downloaded).toBe(false);

    expect(updateCovered('2099-01-01T00:00:00Z', undefined)).toBe(true); // trial
  });

  it('remembers the automatic-check setting', async () => {
    const svc = new UpdateService(dir, fakeUpdater(null), async () => undefined);
    await svc.setAuto(true);
    const again = new UpdateService(dir, fakeUpdater(null), async () => undefined);
    await again.init();
    expect(again.status().auto).toBe(true);
  });
});
