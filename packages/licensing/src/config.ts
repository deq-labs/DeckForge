/**
 * Product licensing configuration, shared by the desktop app and the CLI.
 *
 * Both values are empty until you set them up (see docs/licensing-operations.md):
 * - PRODUCTION_PUBLIC_KEY: generate the real key pair with
 *     pnpm --filter @deckforge/license-server admin keys --out <folder outside the repo>
 *   and paste the printed public key here. Keep the private key out of the repo.
 *   While it is empty, every copy runs in trial mode.
 * - LICENSE_SERVER_URL: the https URL where apps/license-server is deployed. While it is empty,
 *   the app never makes a network call for licensing; offline activation files still work.
 *   (Deliberately not a placeholder domain: a build must never send keys to a host you don't own.)
 */
export const PRODUCTION_PUBLIC_KEY = '';

export const LICENSE_SERVER_URL = '';

export function recoveryUrl(base = LICENSE_SERVER_URL): string | undefined {
  return base ? `${base.replace(/\/$/, '')}/recover` : undefined;
}
