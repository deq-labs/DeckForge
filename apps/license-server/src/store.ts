/** License and activation records in SQLite (Node's built-in node:sqlite). */
import { DatabaseSync } from 'node:sqlite';

export interface LicenseRecord {
  id: string;
  orderId: string | null;
  email: string;
  licensee: string;
  seats: number;
  updatesUntil: string;
  key: string;
  issuedAt: string;
  revoked: boolean;
}

interface LicenseRow {
  id: string;
  order_id: string | null;
  email: string;
  licensee: string;
  seats: number;
  updates_until: string;
  key: string;
  issued_at: string;
  revoked: number;
}

const fromRow = (r: LicenseRow): LicenseRecord => ({
  id: r.id,
  orderId: r.order_id,
  email: r.email,
  licensee: r.licensee,
  seats: r.seats,
  updatesUntil: r.updates_until,
  key: r.key,
  issuedAt: r.issued_at,
  revoked: r.revoked === 1,
});

export class Store {
  private readonly db: DatabaseSync;

  constructor(path: string) {
    this.db = new DatabaseSync(path);
    this.db.exec(`
      PRAGMA journal_mode = WAL;
      CREATE TABLE IF NOT EXISTS licenses (
        id TEXT PRIMARY KEY,
        order_id TEXT UNIQUE,
        email TEXT NOT NULL,
        licensee TEXT NOT NULL,
        seats INTEGER NOT NULL,
        updates_until TEXT NOT NULL,
        key TEXT NOT NULL,
        issued_at TEXT NOT NULL,
        revoked INTEGER NOT NULL DEFAULT 0
      );
      CREATE INDEX IF NOT EXISTS licenses_email ON licenses (email);
      CREATE TABLE IF NOT EXISTS activations (
        license_id TEXT NOT NULL REFERENCES licenses (id),
        machine TEXT NOT NULL,
        activated_at TEXT NOT NULL,
        PRIMARY KEY (license_id, machine)
      );
    `);
  }

  insertLicense(r: LicenseRecord): void {
    this.db
      .prepare('INSERT INTO licenses (id, order_id, email, licensee, seats, updates_until, key, issued_at, revoked) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)')
      .run(r.id, r.orderId, r.email.toLowerCase(), r.licensee, r.seats, r.updatesUntil, r.key, r.issuedAt, r.revoked ? 1 : 0);
  }

  byId(id: string): LicenseRecord | undefined {
    const r = this.db.prepare('SELECT * FROM licenses WHERE id = ?').get(id) as LicenseRow | undefined;
    return r ? fromRow(r) : undefined;
  }

  byOrder(orderId: string): LicenseRecord | undefined {
    const r = this.db.prepare('SELECT * FROM licenses WHERE order_id = ?').get(orderId) as LicenseRow | undefined;
    return r ? fromRow(r) : undefined;
  }

  byEmail(email: string): LicenseRecord[] {
    return (this.db.prepare('SELECT * FROM licenses WHERE email = ? ORDER BY issued_at').all(email.toLowerCase()) as unknown as LicenseRow[]).map(fromRow);
  }

  all(): LicenseRecord[] {
    return (this.db.prepare('SELECT * FROM licenses ORDER BY issued_at').all() as unknown as LicenseRow[]).map(fromRow);
  }

  setRevoked(id: string, revoked: boolean): boolean {
    return Number(this.db.prepare('UPDATE licenses SET revoked = ? WHERE id = ?').run(revoked ? 1 : 0, id).changes) > 0;
  }

  activations(licenseId: string): { machine: string; activatedAt: string }[] {
    return (this.db.prepare('SELECT machine, activated_at FROM activations WHERE license_id = ? ORDER BY activated_at').all(licenseId) as unknown as { machine: string; activated_at: string }[]).map(
      (r) => ({ machine: r.machine, activatedAt: r.activated_at }),
    );
  }

  addActivation(licenseId: string, machine: string, at: string): void {
    this.db.prepare('INSERT OR IGNORE INTO activations (license_id, machine, activated_at) VALUES (?, ?, ?)').run(licenseId, machine, at);
  }

  removeActivation(licenseId: string, machine: string): boolean {
    return Number(this.db.prepare('DELETE FROM activations WHERE license_id = ? AND machine = ?').run(licenseId, machine).changes) > 0;
  }

  close(): void {
    this.db.close();
  }
}
