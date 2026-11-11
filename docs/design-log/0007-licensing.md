# 0007 — Licensing, activation and updates (Lite L4)

- Date: 2026-11-11 · Edition: lite (licensing package: shared) · Status: accepted

## License keys
- `DFL1.<base64url(JSON payload)>.<base64url(Ed25519 signature)>`, signature over `DFL1.<payload>`.
  Payload: id, edition, licensee, seats, issuedAt, `updatesUntil` (YYYY-MM-DD).
- Verified **offline** with the public key compiled into the app (`packages/licensing/src/config.ts`).
  WebCrypto only, no dependencies. Keys can be pasted with line breaks or as an armored
  license file (`-----BEGIN DECKFORGE LICENSE-----`).
- **Perpetual for covered versions:** each build carries its release date (set by the build
  script). A key licenses every build released on or before `updatesUntil`, forever. A newer
  build runs in trial mode and says why ("renew updates, or keep using your version").

## Seats and activation
- A valid key alone makes the app licensed. Seat registration ("activation") is **optional**, as
  the spec requires: online (one POST to the license server, only when the user ticks the box and
  a server is configured) or with an **activation file** for air-gapped computers (request code →
  activation page on another computer → file back).
- So seats are enforced for users who register online, and are an honor system otherwise. That
  is the deliberate trade-off from the spec's risk table (accept some piracy, keep the price fair).
  The server refuses registration past the seat limit and for refunded orders.
- The machine id is a SHA-256 of hostname, OS user, platform and arch, truncated. No raw hardware
  details leave the computer.

## Network policy
- With `LICENSE_SERVER_URL` empty (the default in the repo), the app makes **no** licensing calls.
- Update checks are **off by default**. The user turns on "check when DeckForge starts" or clicks
  "Check for updates". Air-gapped and regulated sites get no surprise traffic. Revisit this
  default after launch if support load shows people running stale versions.
- Licensed users only download updates released while their updates are included. Newer ones are
  announced, not downloaded.

## Overrides
- The public key, release date and server URL can be overridden by environment variables **only**
  in development (unpackaged desktop app, or the CLI run from TypeScript source), so tests can use
  throwaway keys. Packaged/bundled builds ignore them.

## Vendor backend (`apps/license-server`)
- Lemon Squeezy `order_created` (paid) → sign key → store → email key + license file.
  Idempotent per order. `order_refunded` → revoke (blocks future activations; keys already on
  computers keep working offline, which is inherent to offline licensing).
- `/recover` emails all keys for an address, with the same response whether or not it exists, and
  is rate-limited. `/activate` is the offline-activation page. `/api/activate|deactivate` are for the app.
- node:http + node:sqlite: no framework, no ORM, nothing to patch but Node.

## Rejected
- A hosted license SaaS (Keygen, Cryptlex): monthly cost and a dependency for something ~400
  lines can do; also their client SDKs phone home.
- Lemon Squeezy's own license keys: they need an online check, which breaks air-gapped use.
- Mandatory online activation: it contradicts "no data leaves the network" for the Pro-leaning buyers.
