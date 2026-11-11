/**
 * Admin tasks, run on your machine (never on customers' machines).
 *
 *   pnpm --filter @deckforge/license-server admin keys --out <dir>
 *       Generates the production key pair. Prints the public key to paste into
 *       packages/licensing/src/config.ts. Keep <dir> outside the repo and backed up.
 *   pnpm ... admin issue --licensee "Acme" --email a@b.com [--seats 5] [--months 12]
 *       Issues a key by hand (services clients, replacements). Needs DECKFORGE_KEY_DIR and DECKFORGE_DB.
 *   pnpm ... admin list | revoke <id> | unrevoke <id> | activations <id> | free <id> <machine>
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { generateKeyPair } from '@deckforge/licensing';
import { issueLicense } from './app.ts';
import { LogMailer } from './mail.ts';
import { Store } from './store.ts';

const [cmd, ...rest] = process.argv.slice(2);
const flag = (name: string): string | undefined => {
  const i = rest.indexOf(`--${name}`);
  return i >= 0 ? rest[i + 1] : undefined;
};
const fail = (msg: string): never => {
  console.error(msg);
  process.exit(1);
};

function keyDir(): string {
  return process.env.DECKFORGE_KEY_DIR ?? fail('Set DECKFORGE_KEY_DIR to the folder with the key pair.');
}

async function main(): Promise<void> {
  switch (cmd) {
    case 'keys': {
      const out = resolve(flag('out') ?? fail('Usage: admin keys --out <folder outside the repo>'));
      if (existsSync(join(out, 'deckforge-private.key'))) fail(`${out} already has a key pair. Refusing to overwrite it: keys already sold would stop working.`);
      if (existsSync(join(process.cwd(), '.git')) && out.startsWith(process.cwd())) fail('Choose a folder outside the repository for the private key.');
      mkdirSync(out, { recursive: true });
      const kp = await generateKeyPair();
      writeFileSync(join(out, 'deckforge-private.key'), `${kp.privateKey}\n`, { mode: 0o600 });
      writeFileSync(join(out, 'deckforge-public.key'), `${kp.publicKey}\n`);
      console.log(`Wrote ${out}/deckforge-private.key (secret: back it up, never commit it)`);
      console.log(`Wrote ${out}/deckforge-public.key`);
      console.log(`\nPaste this into packages/licensing/src/config.ts as PRODUCTION_PUBLIC_KEY:\n\n  ${kp.publicKey}\n`);
      return;
    }
    case 'issue': {
      const dir = keyDir();
      const store = new Store(process.env.DECKFORGE_DB ?? 'licenses.db');
      const mailer = new LogMailer();
      const rec = await issueLicense(
        {
          store,
          mailer,
          config: {
            privateKey: readFileSync(join(dir, 'deckforge-private.key'), 'utf8').trim(),
            publicKey: readFileSync(join(dir, 'deckforge-public.key'), 'utf8').trim(),
            webhookSecret: '',
            variants: {},
          },
        },
        {
          orderId: null,
          email: flag('email') ?? fail('--email is required'),
          licensee: flag('licensee') ?? fail('--licensee is required'),
          seats: Number(flag('seats') ?? 1),
          updateMonths: Number(flag('months') ?? 12),
        },
      );
      console.log(`Issued ${rec.id} to ${rec.licensee} (${rec.seats} seat(s), updates until ${rec.updatesUntil}).\n\n${rec.key}\n`);
      console.log('(Not emailed: send the key yourself, or have the customer use the recovery page.)');
      return;
    }
    case 'list': {
      const store = new Store(process.env.DECKFORGE_DB ?? 'licenses.db');
      for (const r of store.all()) {
        console.log(`${r.id}\t${r.email}\t${r.licensee}\tseats=${r.seats}\tactive=${store.activations(r.id).length}\tupdates<=${r.updatesUntil}${r.revoked ? '\tREVOKED' : ''}`);
      }
      return;
    }
    case 'revoke':
    case 'unrevoke': {
      const store = new Store(process.env.DECKFORGE_DB ?? 'licenses.db');
      const id = rest[0] ?? fail(`Usage: admin ${cmd} <license id>`);
      console.log(store.setRevoked(id, cmd === 'revoke') ? `${cmd}d ${id}` : `no license ${id}`);
      return;
    }
    case 'activations': {
      const store = new Store(process.env.DECKFORGE_DB ?? 'licenses.db');
      for (const a of store.activations(rest[0] ?? fail('Usage: admin activations <license id>'))) console.log(`${a.machine}\t${a.activatedAt}`);
      return;
    }
    case 'free': {
      const store = new Store(process.env.DECKFORGE_DB ?? 'licenses.db');
      const [id, machine] = rest;
      if (!id || !machine) fail('Usage: admin free <license id> <machine>');
      console.log(store.removeActivation(id!, machine!) ? 'seat freed' : 'no such activation');
      return;
    }
    default:
      console.log('Commands: keys --out <dir> | issue --licensee <name> --email <email> [--seats n] [--months n] | list | revoke <id> | unrevoke <id> | activations <id> | free <id> <machine>');
  }
}

await main();
