/**
 * DeckForge license keys: a JSON payload signed with Ed25519, verified offline with the public
 * key built into the app. See docs/design-log/0007-licensing.md.
 *
 *   DFL1.<base64url(payload JSON)>.<base64url(signature)>
 *
 * The signature covers the ASCII text "DFL1.<payload part>". Activation tokens use the same
 * scheme with the prefix DFA1. Pure: WebCrypto only, no Node APIs, no network.
 */

export type Edition = 'lite' | 'pro';

export interface LicensePayload {
  v: 1;
  /** Unique license id, e.g. the order id. */
  id: string;
  product: 'deckforge';
  edition: Edition;
  /** Person or company the license is for, shown in the app. */
  licensee: string;
  seats: number;
  /** ISO timestamp. */
  issuedAt: string;
  /** Last day (YYYY-MM-DD) of free updates: versions released on or before it are covered forever. */
  updatesUntil: string;
}

export interface ActivationPayload {
  v: 1;
  licenseId: string;
  /** Hash identifying one computer (never raw hardware details). */
  machine: string;
  activatedAt: string;
}

export type Verified<T> = { ok: true; payload: T } | { ok: false; reason: string };

const LICENSE_PREFIX = 'DFL1';
const ACTIVATION_PREFIX = 'DFA1';
const ARMOR = /-----(BEGIN|END) DECKFORGE (LICENSE|ACTIVATION)-----/g;
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

// ---------------------------------------------------------------- encoding

export function toBase64Url(bytes: Uint8Array): string {
  let bin = '';
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

export function fromBase64Url(s: string): Uint8Array {
  if (!/^[A-Za-z0-9_-]*$/.test(s)) throw new Error('not base64url');
  const b64 = s.replace(/-/g, '+').replace(/_/g, '/') + '='.repeat((4 - (s.length % 4)) % 4);
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

const utf8 = new TextEncoder();
const fromUtf8 = new TextDecoder();

/** Accepts a pasted key or license file: strips armor lines, spaces and line breaks. */
export function normalizeKey(text: string): string {
  return text.replace(ARMOR, '').replace(/\s+/g, '');
}

/** Wraps a key in armor lines, 64 characters per line, for a downloadable license file. */
export function armor(key: string, kind: 'LICENSE' | 'ACTIVATION' = 'LICENSE'): string {
  const lines = key.match(/.{1,64}/g) ?? [];
  return [`-----BEGIN DECKFORGE ${kind}-----`, ...lines, `-----END DECKFORGE ${kind}-----`, ''].join('\n');
}

// ---------------------------------------------------------------- keys

export interface KeyPair {
  /** Raw 32-byte Ed25519 public key, base64url. Safe to ship in the app. */
  publicKey: string;
  /** PKCS#8 private key, base64url. Keep secret; never commit. */
  privateKey: string;
}

export async function generateKeyPair(): Promise<KeyPair> {
  // Ed25519 always yields a pair; TS's DOM lib types the overload as CryptoKey | CryptoKeyPair.
  const kp = (await crypto.subtle.generateKey({ name: 'Ed25519' }, true, ['sign', 'verify'])) as unknown as { publicKey: CryptoKey; privateKey: CryptoKey };
  const pub = new Uint8Array(await crypto.subtle.exportKey('raw', kp.publicKey));
  const priv = new Uint8Array(await crypto.subtle.exportKey('pkcs8', kp.privateKey));
  return { publicKey: toBase64Url(pub), privateKey: toBase64Url(priv) };
}

function bytes(b: Uint8Array): ArrayBuffer {
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
}

async function importPrivate(privateKey: string): Promise<CryptoKey> {
  return crypto.subtle.importKey('pkcs8', bytes(fromBase64Url(privateKey)), { name: 'Ed25519' }, false, ['sign']);
}

async function importPublic(publicKey: string): Promise<CryptoKey> {
  return crypto.subtle.importKey('raw', bytes(fromBase64Url(publicKey)), { name: 'Ed25519' }, false, ['verify']);
}

async function sign(prefix: string, payload: object, privateKey: string): Promise<string> {
  const body = `${prefix}.${toBase64Url(utf8.encode(JSON.stringify(payload)))}`;
  const sig = new Uint8Array(await crypto.subtle.sign({ name: 'Ed25519' }, await importPrivate(privateKey), bytes(utf8.encode(body))));
  return `${body}.${toBase64Url(sig)}`;
}

async function open(prefix: string, token: string, publicKey: string): Promise<Verified<unknown>> {
  const parts = normalizeKey(token).split('.');
  if (parts.length !== 3 || parts[0] !== prefix) return { ok: false, reason: 'This is not a DeckForge license key. Paste the whole key from your purchase email.' };
  let payload: unknown;
  let sig: Uint8Array;
  try {
    payload = JSON.parse(fromUtf8.decode(fromBase64Url(parts[1]!)));
    sig = fromBase64Url(parts[2]!);
  } catch {
    return { ok: false, reason: 'This license key is damaged. Copy it again from your purchase email, all of it.' };
  }
  let key: CryptoKey;
  try {
    key = await importPublic(publicKey);
  } catch {
    return { ok: false, reason: 'This copy of DeckForge has no valid license public key configured.' };
  }
  const valid = await crypto.subtle.verify({ name: 'Ed25519' }, key, bytes(sig), bytes(utf8.encode(`${parts[0]}.${parts[1]}`)));
  if (!valid) return { ok: false, reason: 'This license key is not valid. Check that you copied all of it, or use the license-recovery page to get it again.' };
  return { ok: true, payload };
}

// ---------------------------------------------------------------- licenses

export function validatePayload(p: LicensePayload): string | undefined {
  if (p.v !== 1 || p.product !== 'deckforge') return 'unknown license format';
  if (!p.id || !p.licensee) return 'missing id or licensee';
  if (p.edition !== 'lite' && p.edition !== 'pro') return 'unknown edition';
  if (!Number.isInteger(p.seats) || p.seats < 1) return 'seats must be a whole number of at least 1';
  if (!DATE_RE.test(p.updatesUntil)) return 'updatesUntil must be YYYY-MM-DD';
  return undefined;
}

export async function signLicense(payload: LicensePayload, privateKey: string): Promise<string> {
  const problem = validatePayload(payload);
  if (problem) throw new Error(`Cannot sign license: ${problem}`);
  return sign(LICENSE_PREFIX, payload, privateKey);
}

export async function verifyLicense(key: string, publicKey: string): Promise<Verified<LicensePayload>> {
  const r = await open(LICENSE_PREFIX, key, publicKey);
  if (!r.ok) return r;
  const p = r.payload as LicensePayload;
  if (validatePayload(p)) return { ok: false, reason: 'This license key is from a newer or different product version than this app understands.' };
  return { ok: true, payload: p };
}

/**
 * Whether a license covers an app version released on `releaseDate` (YYYY-MM-DD): perpetual for
 * every version released while updates were included.
 */
export function coversRelease(license: LicensePayload, releaseDate: string): { covered: true } | { covered: false; reason: string } {
  if (releaseDate <= license.updatesUntil) return { covered: true };
  return {
    covered: false,
    reason:
      `Your license includes versions released until ${license.updatesUntil}; this version was released ${releaseDate}. ` +
      'Keep using the version you have, or renew updates to use this one.',
  };
}

// ---------------------------------------------------------------- activations

export async function signActivation(payload: ActivationPayload, privateKey: string): Promise<string> {
  return sign(ACTIVATION_PREFIX, payload, privateKey);
}

export async function verifyActivation(token: string, publicKey: string, expect: { licenseId: string; machine: string }): Promise<Verified<ActivationPayload>> {
  const r = await open(ACTIVATION_PREFIX, token, publicKey);
  if (!r.ok) return { ok: false, reason: r.reason.replace('license key', 'activation file') };
  const p = r.payload as ActivationPayload;
  if (p.v !== 1 || p.licenseId !== expect.licenseId) return { ok: false, reason: 'This activation file is for a different license.' };
  if (p.machine !== expect.machine) return { ok: false, reason: 'This activation file is for a different computer. Create a new activation request on this computer.' };
  return { ok: true, payload: p };
}

/**
 * The code a user carries from an offline computer to the activation page:
 * license id + machine hash, so the server can issue an activation file.
 */
export function activationRequestCode(licenseId: string, machine: string): string {
  return `DFR1.${toBase64Url(utf8.encode(JSON.stringify({ licenseId, machine })))}`;
}

export function parseActivationRequest(code: string): { licenseId: string; machine: string } | undefined {
  const parts = normalizeKey(code).split('.');
  if (parts.length !== 2 || parts[0] !== 'DFR1') return undefined;
  try {
    const r = JSON.parse(fromUtf8.decode(fromBase64Url(parts[1]!))) as { licenseId?: unknown; machine?: unknown };
    return typeof r.licenseId === 'string' && typeof r.machine === 'string' ? { licenseId: r.licenseId, machine: r.machine } : undefined;
  } catch {
    return undefined;
  }
}

export { PRODUCTION_PUBLIC_KEY, LICENSE_SERVER_URL, recoveryUrl } from './config.ts';
export { evaluateLicense } from './state.ts';
export type { LicenseState } from './state.ts';
