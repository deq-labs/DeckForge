/**
 * HTTP handlers. Kept framework-free (node:http) so the whole backend is a few hundred lines
 * with no dependencies besides @deckforge/licensing.
 *
 *   POST /webhooks/lemonsqueezy   order_created -> issue + email key; order_refunded -> revoke
 *   GET  /recover, POST /recover  email lost keys (same answer whether or not the email exists)
 *   POST /api/activate            {key, machine} -> {activation} (seat-limited)
 *   POST /api/deactivate          {key, machine} -> frees the seat
 *   GET  /activate, POST /activate  offline activation: request code + key -> activation file
 *   GET  /health
 */
import { createHmac, timingSafeEqual } from 'node:crypto';
import type { IncomingMessage, ServerResponse } from 'node:http';
import { armor, parseActivationRequest, signActivation, signLicense, verifyLicense } from '@deckforge/licensing';
import type { LicensePayload } from '@deckforge/licensing';
import type { LicenseRecord, Store } from './store.ts';
import { licenseEmail } from './mail.ts';
import type { Mailer } from './mail.ts';

export interface Variant {
  seats: number;
  /** Months of free updates. Default 12. */
  updateMonths?: number;
  label?: string;
}

export interface Config {
  publicKey: string;
  privateKey: string;
  /** Lemon Squeezy webhook signing secret. */
  webhookSecret: string;
  /** Lemon Squeezy variant id -> what it sells. */
  variants: Record<string, Variant>;
  now?: () => Date;
}

const MAX_BODY = 256 * 1024;

export function addMonths(d: Date, months: number): string {
  const x = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + months, d.getUTCDate()));
  return x.toISOString().slice(0, 10);
}

/** Issues (signs, stores, emails) a license. Idempotent per order id. */
export async function issueLicense(
  deps: { store: Store; mailer: Mailer; config: Config },
  input: { orderId: string | null; email: string; licensee: string; seats: number; updateMonths: number },
): Promise<LicenseRecord> {
  if (input.orderId) {
    const existing = deps.store.byOrder(input.orderId);
    if (existing) return existing;
  }
  const now = (deps.config.now ?? (() => new Date()))();
  const id = input.orderId ? `LS-${input.orderId}` : `M-${now.getTime().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
  const payload: LicensePayload = {
    v: 1,
    id,
    product: 'deckforge',
    edition: 'lite',
    licensee: input.licensee,
    seats: input.seats,
    issuedAt: now.toISOString(),
    updatesUntil: addMonths(now, input.updateMonths),
  };
  const key = await signLicense(payload, deps.config.privateKey);
  const record: LicenseRecord = {
    id,
    orderId: input.orderId,
    email: input.email,
    licensee: input.licensee,
    seats: input.seats,
    updatesUntil: payload.updatesUntil,
    key,
    issuedAt: payload.issuedAt,
    revoked: false,
  };
  deps.store.insertLicense(record);
  await deps.mailer.send(licenseEmail(input.email, [{ ...record, file: armor(key) }]));
  return record;
}

// ---------------------------------------------------------------- http helpers

async function readBody(req: IncomingMessage): Promise<string> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req) {
    size += (chunk as Buffer).length;
    if (size > MAX_BODY) throw Object.assign(new Error('body too large'), { status: 413 });
    chunks.push(chunk as Buffer);
  }
  return Buffer.concat(chunks).toString('utf8');
}

function send(res: ServerResponse, status: number, body: string, type = 'application/json; charset=utf-8', extra: Record<string, string> = {}): void {
  res.writeHead(status, { 'content-type': type, 'cache-control': 'no-store', 'x-content-type-options': 'nosniff', ...extra });
  res.end(body);
}
const json = (res: ServerResponse, status: number, body: unknown): void => send(res, status, JSON.stringify(body));

const esc = (s: string): string => s.replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);

function page(title: string, body: string): string {
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${esc(title)} · DeckForge</title>
<style>body{font:16px/1.5 system-ui,sans-serif;max-width:560px;margin:48px auto;padding:0 16px;color:#1f2328}
input,textarea{font:inherit;width:100%;padding:8px;margin:4px 0 12px;border:1px solid #ccd;border-radius:6px}
button{font:inherit;padding:8px 16px;border-radius:6px;border:0;background:#1f3864;color:#fff}
.muted{color:#646d78}@media(prefers-color-scheme:dark){body{background:#15181c;color:#e6e8eb}input,textarea{background:#1d2126;color:#e6e8eb;border-color:#2e343b}}</style>
</head><body><h1>${esc(title)}</h1>${body}</body></html>`;
}

/** Tiny fixed-window rate limiter, per key. */
class RateLimit {
  private readonly hits = new Map<string, { n: number; reset: number }>();
  private readonly limit: number;
  private readonly windowMs: number;
  constructor(limit: number, windowMs: number) {
    this.limit = limit;
    this.windowMs = windowMs;
  }
  allow(key: string, now = Date.now()): boolean {
    const h = this.hits.get(key);
    if (!h || h.reset < now) {
      this.hits.set(key, { n: 1, reset: now + this.windowMs });
      return true;
    }
    h.n++;
    return h.n <= this.limit;
  }
}

export function verifyWebhookSignature(rawBody: string, signature: string | undefined, secret: string): boolean {
  if (!signature || !secret) return false;
  const expected = createHmac('sha256', secret).update(rawBody).digest('hex');
  const a = Buffer.from(expected, 'utf8');
  const b = Buffer.from(signature, 'utf8');
  return a.length === b.length && timingSafeEqual(a, b);
}

// ---------------------------------------------------------------- app

export function createHandler(deps: { store: Store; mailer: Mailer; config: Config }): (req: IncomingMessage, res: ServerResponse) => Promise<void> {
  const { store, config } = deps;
  const recoverLimit = new RateLimit(5, 60 * 60 * 1000);
  const activateLimit = new RateLimit(30, 60 * 60 * 1000);
  const now = (): Date => (config.now ?? (() => new Date()))();

  /** Shared by online and offline activation. */
  async function activate(key: string, machine: string): Promise<{ status: number; body: { activation?: string; error?: string } }> {
    const v = await verifyLicense(key, config.publicKey);
    if (!v.ok) return { status: 400, body: { error: v.reason } };
    const rec = store.byId(v.payload.id);
    if (!rec) return { status: 404, body: { error: 'This license is not in our records. Contact support with your order email.' } };
    if (rec.revoked) return { status: 403, body: { error: 'This license was refunded or revoked, so it can no longer be activated.' } };
    if (!/^[A-Za-z0-9_-]{8,128}$/.test(machine)) return { status: 400, body: { error: 'The computer identifier is missing or malformed.' } };
    const active = store.activations(rec.id);
    if (!active.some((a) => a.machine === machine) && active.length >= rec.seats) {
      return {
        status: 409,
        body: {
          error:
            `This license is already active on ${active.length} computer${active.length === 1 ? '' : 's'}, its limit. ` +
            'Remove it from one of them (License > Remove from this computer), or contact support to move it.',
        },
      };
    }
    const at = now().toISOString();
    store.addActivation(rec.id, machine, at);
    const activation = await signActivation({ v: 1, licenseId: rec.id, machine, activatedAt: at }, config.privateKey);
    return { status: 200, body: { activation } };
  }

  return async (req, res) => {
    const url = new URL(req.url ?? '/', 'http://localhost');
    const ip = (req.socket.remoteAddress ?? 'unknown').replace(/^::ffff:/, '');
    try {
      if (req.method === 'GET' && url.pathname === '/health') return json(res, 200, { ok: true });

      if (req.method === 'POST' && url.pathname === '/webhooks/lemonsqueezy') {
        const raw = await readBody(req);
        if (!verifyWebhookSignature(raw, req.headers['x-signature'] as string | undefined, config.webhookSecret)) return json(res, 401, { error: 'bad signature' });
        const event = JSON.parse(raw) as {
          meta?: { event_name?: string };
          data?: { id?: string; attributes?: { status?: string; user_email?: string; user_name?: string; first_order_item?: { variant_id?: number | string } } };
        };
        const name = event.meta?.event_name;
        const orderId = String(event.data?.id ?? '');
        const attrs = event.data?.attributes;
        if (name === 'order_created') {
          if (attrs?.status !== 'paid') return json(res, 200, { ignored: `status ${attrs?.status}` });
          const variant = config.variants[String(attrs.first_order_item?.variant_id ?? '')];
          if (!variant) return json(res, 422, { error: `unknown variant ${attrs.first_order_item?.variant_id}` });
          if (!attrs.user_email) return json(res, 422, { error: 'order has no email' });
          const rec = await issueLicense(deps, {
            orderId,
            email: attrs.user_email,
            licensee: attrs.user_name || attrs.user_email,
            seats: variant.seats,
            updateMonths: variant.updateMonths ?? 12,
          });
          return json(res, 200, { license: rec.id });
        }
        if (name === 'order_refunded') {
          const rec = store.byOrder(orderId);
          if (rec) store.setRevoked(rec.id, true);
          return json(res, 200, { revoked: rec?.id ?? null });
        }
        return json(res, 200, { ignored: name });
      }

      if (url.pathname === '/recover') {
        if (req.method === 'GET') {
          return send(
            res,
            200,
            page(
              'Recover your license key',
              `<p class="muted">Enter the email address you used to buy DeckForge. We'll email every license key for it.</p>
<form method="post"><label>Email<input type="email" name="email" required autocomplete="email"></label><button>Email my keys</button></form>`,
            ),
            'text/html; charset=utf-8',
          );
        }
        if (req.method === 'POST') {
          const email = (new URLSearchParams(await readBody(req)).get('email') ?? '').trim().toLowerCase();
          const done = page('Check your email', '<p>If that address has a DeckForge license, the keys are on their way. Check your spam folder too.</p>');
          if (!recoverLimit.allow(ip) || !recoverLimit.allow(`e:${email}`)) return send(res, 429, page('Too many requests', '<p>Please try again in an hour.</p>'), 'text/html; charset=utf-8');
          const recs = email ? store.byEmail(email).filter((r) => !r.revoked) : [];
          if (recs.length) await deps.mailer.send(licenseEmail(email, recs.map((r) => ({ ...r, file: armor(r.key) }))));
          // Same answer either way, so the page can't be used to discover customers.
          return send(res, 200, done, 'text/html; charset=utf-8');
        }
      }

      if (req.method === 'POST' && (url.pathname === '/api/activate' || url.pathname === '/api/deactivate')) {
        if (!activateLimit.allow(ip)) return json(res, 429, { error: 'Too many activation attempts. Try again later.' });
        let body: { key?: string; machine?: string };
        try {
          body = JSON.parse(await readBody(req)) as { key?: string; machine?: string };
        } catch {
          return json(res, 400, { error: 'Expected JSON.' });
        }
        if (url.pathname === '/api/deactivate') {
          const v = await verifyLicense(body.key ?? '', config.publicKey);
          if (!v.ok) return json(res, 400, { error: v.reason });
          return json(res, 200, { removed: store.removeActivation(v.payload.id, body.machine ?? '') });
        }
        const r = await activate(body.key ?? '', body.machine ?? '');
        return json(res, r.status, r.body);
      }

      if (url.pathname === '/activate') {
        const form = (error = ''): string =>
          page(
            'Activate an offline computer',
            `${error ? `<p style="color:#c62828">${esc(error)}</p>` : ''}
<p class="muted">On the computer without internet, open DeckForge &gt; License &gt; "Activate offline" and copy the request code. Paste it here with your license key to get an activation file, then load that file on the offline computer.</p>
<form method="post"><label>Request code<input name="request" required></label><label>License key<textarea name="key" rows="4" required></textarea></label><button>Get activation file</button></form>`,
          );
        if (req.method === 'GET') return send(res, 200, form(), 'text/html; charset=utf-8');
        if (req.method === 'POST') {
          if (!activateLimit.allow(ip)) return send(res, 429, form('Too many attempts. Try again later.'), 'text/html; charset=utf-8');
          const f = new URLSearchParams(await readBody(req));
          const request = parseActivationRequest(f.get('request') ?? '');
          if (!request) return send(res, 400, form('That request code is not valid. Copy it again from DeckForge.'), 'text/html; charset=utf-8');
          const r = await activate(f.get('key') ?? '', request.machine);
          if (!r.body.activation) return send(res, r.status, form(r.body.error ?? 'Activation failed.'), 'text/html; charset=utf-8');
          return send(res, 200, armor(r.body.activation, 'ACTIVATION'), 'text/plain; charset=utf-8', {
            'content-disposition': 'attachment; filename="deckforge-activation.txt"',
          });
        }
      }

      return json(res, 404, { error: 'not found' });
    } catch (e) {
      const status = (e as { status?: number }).status ?? 500;
      if (status === 500) console.error(e);
      return json(res, status, { error: status === 500 ? 'Something went wrong on our side. Please try again.' : (e as Error).message });
    }
  };
}
