import { createHmac } from 'node:crypto';
import { createServer } from 'node:http';
import type { Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { activationRequestCode, generateKeyPair, normalizeKey, verifyActivation, verifyLicense } from '@deckforge/licensing';
import type { KeyPair } from '@deckforge/licensing';
import { addMonths, createHandler, issueLicense } from '../src/app.ts';
import type { Config } from '../src/app.ts';
import { LogMailer } from '../src/mail.ts';
import { Store } from '../src/store.ts';

let server: Server;
let base: string;
let store: Store;
let mailer: LogMailer;
let keys: KeyPair;
let config: Config;
const SECRET = 'whsec_test';

beforeEach(async () => {
  keys = await generateKeyPair();
  store = new Store(':memory:');
  mailer = new LogMailer();
  config = { ...keys, webhookSecret: SECRET, variants: { '111': { seats: 1 }, '555': { seats: 5, updateMonths: 12 } }, now: () => new Date('2026-11-11T10:00:00Z') };
  const handler = createHandler({ store, mailer, config });
  server = createServer((req, res) => void handler(req, res));
  await new Promise<void>((r) => server.listen(0, '127.0.0.1', () => r()));
  base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
});
afterEach(async () => {
  await new Promise((r) => server.close(r));
  store.close();
});

function order(id: string, variant: number, extra: Record<string, unknown> = {}): string {
  return JSON.stringify({
    meta: { event_name: 'order_created' },
    data: { id, attributes: { status: 'paid', user_email: 'Buyer@Example.com', user_name: 'Acme Plastics', first_order_item: { variant_id: variant }, ...extra } },
  });
}
const sign = (body: string, secret = SECRET): string => createHmac('sha256', secret).update(body).digest('hex');
const post = (path: string, body: string, headers: Record<string, string> = {}): Promise<Response> =>
  fetch(`${base}${path}`, { method: 'POST', body, headers: { 'content-type': 'application/json', ...headers } });

describe('checkout webhook', () => {
  it('issues a signed key for a paid order, emails it, and is idempotent', async () => {
    const body = order('9001', 555);
    const r = await post('/webhooks/lemonsqueezy', body, { 'x-signature': sign(body) });
    expect(r.status).toBe(200);
    expect(await r.json()).toEqual({ license: 'LS-9001' });
    const rec = store.byId('LS-9001')!;
    expect(rec).toMatchObject({ email: 'buyer@example.com', licensee: 'Acme Plastics', seats: 5, updatesUntil: '2027-11-11' });
    const v = await verifyLicense(rec.key, keys.publicKey);
    expect(v.ok && v.payload.seats).toBe(5);

    expect(mailer.sent).toHaveLength(1);
    expect(mailer.sent[0]!.text).toContain(rec.key);
    expect(mailer.sent[0]!.attachment!.content).toContain('-----BEGIN DECKFORGE LICENSE-----');

    await post('/webhooks/lemonsqueezy', body, { 'x-signature': sign(body) }); // retried delivery
    expect(store.all()).toHaveLength(1);
    expect(mailer.sent).toHaveLength(1);
  });

  it('rejects bad signatures, unpaid orders and unknown products', async () => {
    const body = order('1', 111);
    expect((await post('/webhooks/lemonsqueezy', body, { 'x-signature': sign(body, 'wrong') })).status).toBe(401);
    expect((await post('/webhooks/lemonsqueezy', body)).status).toBe(401);
    const pending = order('2', 111, { status: 'pending' });
    expect(await (await post('/webhooks/lemonsqueezy', pending, { 'x-signature': sign(pending) })).json()).toEqual({ ignored: 'status pending' });
    const unknown = order('3', 999);
    expect((await post('/webhooks/lemonsqueezy', unknown, { 'x-signature': sign(unknown) })).status).toBe(422);
    expect(store.all()).toEqual([]);
  });

  it('revokes on refund', async () => {
    const body = order('77', 111);
    await post('/webhooks/lemonsqueezy', body, { 'x-signature': sign(body) });
    const refund = JSON.stringify({ meta: { event_name: 'order_refunded' }, data: { id: '77', attributes: {} } });
    await post('/webhooks/lemonsqueezy', refund, { 'x-signature': sign(refund) });
    expect(store.byId('LS-77')!.revoked).toBe(true);
  });
});

describe('activation', () => {
  async function license(seats: number): Promise<string> {
    return (await issueLicense({ store, mailer, config }, { orderId: `o${seats}`, email: 'a@b.c', licensee: 'A', seats, updateMonths: 12 })).key;
  }

  it('activates up to the seat limit, re-activates the same computer, and frees seats', async () => {
    const key = await license(2);
    const act = (machine: string): Promise<Response> => post('/api/activate', JSON.stringify({ key, machine }));
    const r1 = await act('machine-aaaa');
    expect(r1.status).toBe(200);
    const { activation } = (await r1.json()) as { activation: string };
    expect((await verifyActivation(activation, keys.publicKey, { licenseId: 'LS-o2', machine: 'machine-aaaa' })).ok).toBe(true);
    expect((await act('machine-bbbb')).status).toBe(200);
    expect((await act('machine-aaaa')).status).toBe(200); // same computer again: no new seat
    const full = await act('machine-cccc');
    expect(full.status).toBe(409);
    expect(((await full.json()) as { error: string }).error).toMatch(/already active on 2 computers/);

    expect(await (await post('/api/deactivate', JSON.stringify({ key, machine: 'machine-aaaa' }))).json()).toEqual({ removed: true });
    expect((await act('machine-cccc')).status).toBe(200);
  });

  it('refuses forged and revoked keys', async () => {
    const other = await generateKeyPair();
    const { signLicense } = await import('@deckforge/licensing');
    const forged = await signLicense({ v: 1, id: 'LS-x', product: 'deckforge', edition: 'lite', licensee: 'x', seats: 99, issuedAt: '', updatesUntil: '2099-01-01' }, other.privateKey);
    expect((await post('/api/activate', JSON.stringify({ key: forged, machine: 'machine-aaaa' }))).status).toBe(400);
    const key = await license(1);
    store.setRevoked('LS-o1', true);
    expect((await post('/api/activate', JSON.stringify({ key, machine: 'machine-aaaa' }))).status).toBe(403);
  });

  it('serves an activation file for an offline computer', async () => {
    const key = await license(1);
    const form = new URLSearchParams({ request: activationRequestCode('LS-o1', 'offline-machine-1'), key });
    const r = await fetch(`${base}/activate`, { method: 'POST', body: form, headers: { 'content-type': 'application/x-www-form-urlencoded' } });
    expect(r.status).toBe(200);
    expect(r.headers.get('content-disposition')).toMatch(/deckforge-activation\.txt/);
    const file = await r.text();
    expect(file).toMatch(/^-----BEGIN DECKFORGE ACTIVATION-----/);
    expect((await verifyActivation(normalizeKey(file), keys.publicKey, { licenseId: 'LS-o1', machine: 'offline-machine-1' })).ok).toBe(true);
  });
});

describe('recovery', () => {
  it('emails keys for a known address and answers the same for unknown ones', async () => {
    await issueLicense({ store, mailer, config }, { orderId: 'r1', email: 'Known@Example.com', licensee: 'K', seats: 1, updateMonths: 12 });
    mailer.sent.length = 0;
    const ask = (email: string): Promise<Response> => fetch(`${base}/recover`, { method: 'POST', body: new URLSearchParams({ email }), headers: { 'content-type': 'application/x-www-form-urlencoded' } });
    const known = await (await ask('known@example.com')).text();
    const unknown = await (await ask('nobody@example.com')).text();
    expect(known).toBe(unknown);
    expect(mailer.sent.map((m) => m.to)).toEqual(['known@example.com']);
    expect((await fetch(`${base}/recover`)).status).toBe(200);
  });

  it('rate-limits', async () => {
    const ask = (): Promise<Response> => fetch(`${base}/recover`, { method: 'POST', body: new URLSearchParams({ email: 'x@y.z' }), headers: { 'content-type': 'application/x-www-form-urlencoded' } });
    const statuses = [];
    for (let i = 0; i < 6; i++) statuses.push((await ask()).status);
    expect(statuses.at(-1)).toBe(429);
  });
});

describe('addMonths', () => {
  it('adds calendar months in UTC', () => {
    expect(addMonths(new Date('2026-11-11T23:30:00Z'), 12)).toBe('2027-11-11');
    expect(addMonths(new Date('2026-01-31T00:00:00Z'), 1)).toBe('2026-03-03');
  });
});
