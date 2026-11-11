import { coversRelease, verifyActivation, verifyLicense } from './index.ts';
import type { LicensePayload } from './index.ts';

export interface LicenseState {
  mode: 'trial' | 'licensed';
  license?: LicensePayload;
  /** Why a stored key isn't giving a licensed state (invalid, or this version isn't covered). */
  problem?: string;
  /** True when a valid activation for this computer is on file (seat registered). */
  activated: boolean;
}

/**
 * Decides trial vs licensed from what's stored on this computer. Pure and offline.
 * `releaseDate` (YYYY-MM-DD) is this build's release date; omit it in development.
 */
export async function evaluateLicense(opts: {
  key?: string | undefined;
  activation?: string | undefined;
  publicKey: string;
  releaseDate?: string | undefined;
  machine?: string | undefined;
}): Promise<LicenseState> {
  if (!opts.key) return { mode: 'trial', activated: false };
  if (!opts.publicKey) return { mode: 'trial', activated: false, problem: 'This build of DeckForge has no license key configured, so it runs in trial mode.' };
  const v = await verifyLicense(opts.key, opts.publicKey);
  if (!v.ok) return { mode: 'trial', activated: false, problem: v.reason };
  const license = v.payload;
  let activated = false;
  if (opts.activation && opts.machine) {
    activated = (await verifyActivation(opts.activation, opts.publicKey, { licenseId: license.id, machine: opts.machine })).ok;
  }
  if (opts.releaseDate) {
    const c = coversRelease(license, opts.releaseDate);
    if (!c.covered) return { mode: 'trial', license, activated, problem: c.reason };
  }
  return { mode: 'licensed', license, activated };
}
