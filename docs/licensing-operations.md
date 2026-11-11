# Licensing operations

How to run DeckForge licensing for real. Design: `docs/design-log/0007-licensing.md`.

## 1. Generate the production key pair (once, ever)

```sh
pnpm --filter @deckforge/license-server admin keys --out <folder OUTSIDE the repo>
```

- `deckforge-private.key` signs every license. **Back it up in two places** (password manager +
  offline copy). If it leaks, anyone can mint keys. If it's lost, you can't issue keys that
  existing app versions accept.
- Paste the printed public key into `packages/licensing/src/config.ts` → `PRODUCTION_PUBLIC_KEY`,
  commit, and release. Until then, every build runs in trial mode.

## 2. Lemon Squeezy

1. Create the product with two variants: Single user ($199) and Team, 5 seats ($499).
2. Settings → Webhooks: URL `https://<your server>/webhooks/lemonsqueezy`, events
   `order_created` and `order_refunded`, and a signing secret.
3. Note each variant id for `DECKFORGE_VARIANTS`.
4. Turn off Lemon Squeezy's own license keys for this product: DeckForge issues its own.

## 3. Deploy the license server

Any small Node ≥ 24 host works (Fly.io, Railway, a $5 VPS). It needs persistent disk for SQLite.

```sh
DECKFORGE_KEY_DIR=/secrets/deckforge \
DECKFORGE_DB=/data/licenses.db \
LEMONSQUEEZY_WEBHOOK_SECRET=... \
DECKFORGE_VARIANTS='{"<single id>":{"seats":1},"<team id>":{"seats":5}}' \
RESEND_API_KEY=... MAIL_FROM='DeckForge <licenses@yourdomain>' \
node apps/license-server/src/main.ts
```

- Put it behind HTTPS. Back up `licenses.db` daily.
- Then set `LICENSE_SERVER_URL` in `packages/licensing/src/config.ts` to its https URL and release.
  This turns on "Register this computer online", the "Lost your key?" link, and the
  offline-activation page link in the app.
- Without `RESEND_API_KEY`, emails are only logged, which is fine for a dry run.

## 4. Day-to-day

```sh
# all from apps/license-server, with DECKFORGE_KEY_DIR and DECKFORGE_DB set
pnpm admin list                                  # licenses, seats in use
pnpm admin issue --licensee "Acme" --email a@acme.com --seats 5   # services clients, replacements
pnpm admin activations LS-1234                   # which computers
pnpm admin free LS-1234 <machine>                # move a seat
pnpm admin revoke LS-1234                        # blocks new activations
```

Renewals: issue a new key with a later `updatesUntil` (`--months`). Old keys keep working for
the versions they cover.

## Test before launch

- [ ] Webhook test from Lemon Squeezy (test mode) → key email arrives → key activates in the app
- [ ] Refund in test mode → online activation refused
- [ ] Recovery page emails the key
- [ ] Offline activation: request code → `/activate` → file → app shows "Registered"
- [ ] A build dated after a test key's `updatesUntil` shows the "renew updates" message
