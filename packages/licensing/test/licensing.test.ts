import { describe, expect, it } from 'vitest';
import {
  activationRequestCode, armor, coversRelease, fromBase64Url, generateKeyPair, parseActivationRequest,
  signActivation, signLicense, toBase64Url, verifyActivation, verifyLicense,
} from '../src/index.ts';
import type { LicensePayload } from '../src/index.ts';

const payload: LicensePayload = {
  v: 1,
  id: 'order-1001',
  product: 'deckforge',
  edition: 'lite',
  licensee: 'Acme Plastics',
  seats: 5,
  issuedAt: '2026-11-11T12:00:00Z',
  updatesUntil: '2027-11-11',
};

describe('base64url', () => {
  it('round-trips arbitrary bytes', () => {
    const b = new Uint8Array([0, 1, 2, 250, 251, 252, 253, 254, 255]);
    expect(fromBase64Url(toBase64Url(b))).toEqual(b);
    expect(toBase64Url(new Uint8Array([251, 255]))).toBe('-_8');
  });
});

describe('license keys', () => {
  it('verify offline with the public key, and survive pasting with armor and line breaks', async () => {
    const kp = await generateKeyPair();
    const key = await signLicense(payload, kp.privateKey);
    expect(key.startsWith('DFL1.')).toBe(true);
    expect(await verifyLicense(key, kp.publicKey)).toEqual({ ok: true, payload });
    expect(await verifyLicense(`  ${armor(key)}  `, kp.publicKey)).toEqual({ ok: true, payload });
  });

  it('rejects tampering, other keys, and junk with plain-language reasons', async () => {
    const kp = await generateKeyPair();
    const other = await generateKeyPair();
    const key = await signLicense(payload, kp.privateKey);
    const [prefix, body, sig] = key.split('.');
    const forged = toBase64Url(new TextEncoder().encode(JSON.stringify({ ...payload, seats: 500 })));

    const bad = async (k: string, pub = kp.publicKey): Promise<string> => {
      const r = await verifyLicense(k, pub);
      if (r.ok) throw new Error('expected failure');
      return r.reason;
    };
    expect(await bad(`${prefix}.${forged}.${sig}`)).toMatch(/not valid/);
    expect(await bad(key, other.publicKey)).toMatch(/not valid/);
    expect(await bad('hello')).toMatch(/not a DeckForge license key/);
    expect(await bad(`DFL1.${body}.%%%`)).toMatch(/damaged/);
    expect(await bad(key, 'not-a-key')).toMatch(/no valid license public key/);
  });

  it('refuses to sign an invalid payload', async () => {
    const kp = await generateKeyPair();
    await expect(signLicense({ ...payload, seats: 0 }, kp.privateKey)).rejects.toThrow(/seats/);
    await expect(signLicense({ ...payload, updatesUntil: 'soon' }, kp.privateKey)).rejects.toThrow(/YYYY-MM-DD/);
  });

  it('covers every version released while updates were included, forever', () => {
    expect(coversRelease(payload, '2027-11-11')).toEqual({ covered: true });
    expect(coversRelease(payload, '2026-12-01')).toEqual({ covered: true });
    const r = coversRelease(payload, '2027-11-12');
    expect(r.covered).toBe(false);
    if (!r.covered) expect(r.reason).toMatch(/released until 2027-11-11; this version was released 2027-11-12/);
  });
});

describe('activations', () => {
  it('bind a license to one computer', async () => {
    const kp = await generateKeyPair();
    const token = await signActivation({ v: 1, licenseId: 'order-1001', machine: 'm1', activatedAt: '2026-11-11T12:00:00Z' }, kp.privateKey);
    expect((await verifyActivation(token, kp.publicKey, { licenseId: 'order-1001', machine: 'm1' })).ok).toBe(true);
    const wrongMachine = await verifyActivation(token, kp.publicKey, { licenseId: 'order-1001', machine: 'm2' });
    expect(wrongMachine).toEqual({ ok: false, reason: expect.stringMatching(/different computer/) });
    const wrongLicense = await verifyActivation(token, kp.publicKey, { licenseId: 'x', machine: 'm1' });
    expect(wrongLicense).toEqual({ ok: false, reason: expect.stringMatching(/different license/) });
  });

  it('round-trips the offline activation request code', () => {
    const code = activationRequestCode('order-1001', 'abc123');
    expect(parseActivationRequest(code)).toEqual({ licenseId: 'order-1001', machine: 'abc123' });
    expect(parseActivationRequest('nope')).toBeUndefined();
  });
});

describe('evaluateLicense', () => {
  it('goes from trial to licensed, and explains why not', async () => {
    const { evaluateLicense } = await import('../src/index.ts');
    const kp = await generateKeyPair();
    const key = await signLicense(payload, kp.privateKey);
    expect(await evaluateLicense({ publicKey: kp.publicKey })).toEqual({ mode: 'trial', activated: false });
    expect(await evaluateLicense({ key, publicKey: '' })).toMatchObject({ mode: 'trial', problem: expect.stringMatching(/no license key configured/) });
    expect(await evaluateLicense({ key, publicKey: kp.publicKey, releaseDate: '2027-01-01' })).toMatchObject({ mode: 'licensed', license: payload, activated: false });
    expect(await evaluateLicense({ key, publicKey: kp.publicKey, releaseDate: '2028-01-01' })).toMatchObject({ mode: 'trial', license: payload, problem: expect.stringMatching(/renew updates/) });
    const activation = await signActivation({ v: 1, licenseId: payload.id, machine: 'm', activatedAt: '2026-11-11T00:00:00Z' }, kp.privateKey);
    expect(await evaluateLicense({ key, activation, machine: 'm', publicKey: kp.publicKey })).toMatchObject({ mode: 'licensed', activated: true });
  });
});
