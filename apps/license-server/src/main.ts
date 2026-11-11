/**
 * Starts the license server. Configuration comes from the environment:
 *
 *   PORT                        default 8787
 *   DECKFORGE_DB                SQLite file, default ./licenses.db
 *   DECKFORGE_KEY_DIR           folder with deckforge-private.key and deckforge-public.key
 *                               (from `pnpm admin keys --out <dir>`); keep it outside the repo
 *   LEMONSQUEEZY_WEBHOOK_SECRET signing secret of the store's webhook
 *   DECKFORGE_VARIANTS          JSON: {"<variant id>": {"seats": 1, "updateMonths": 12}, ...}
 *   RESEND_API_KEY, MAIL_FROM   optional; without them, emails are only logged
 */
import { createServer } from 'node:http';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { createHandler } from './app.ts';
import type { Config } from './app.ts';
import { Store } from './store.ts';
import { LogMailer, ResendMailer } from './mail.ts';

function required(name: string): string {
  const v = process.env[name];
  if (!v) {
    console.error(`Missing environment variable ${name}. See the top of src/main.ts.`);
    process.exit(1);
  }
  return v;
}

const keyDir = required('DECKFORGE_KEY_DIR');
const config: Config = {
  privateKey: readFileSync(join(keyDir, 'deckforge-private.key'), 'utf8').trim(),
  publicKey: readFileSync(join(keyDir, 'deckforge-public.key'), 'utf8').trim(),
  webhookSecret: required('LEMONSQUEEZY_WEBHOOK_SECRET'),
  variants: JSON.parse(required('DECKFORGE_VARIANTS')) as Config['variants'],
};
const store = new Store(process.env.DECKFORGE_DB ?? 'licenses.db');
const mailer = process.env.RESEND_API_KEY ? new ResendMailer(process.env.RESEND_API_KEY, required('MAIL_FROM')) : new LogMailer();
const port = Number(process.env.PORT ?? 8787);

const handler = createHandler({ store, mailer, config }); // once: it holds the rate limiters
createServer((req, res) => void handler(req, res)).listen(port, () => {
  console.log(`DeckForge license server on :${port} (${mailer instanceof LogMailer ? 'emails logged only' : 'emails via Resend'})`);
});
