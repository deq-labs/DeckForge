/** Outgoing email. The provider is your choice; Resend's HTTP API is built in, logging is the default. */

export interface Mail {
  to: string;
  subject: string;
  text: string;
  /** Optional attachment, e.g. the license file. */
  attachment?: { filename: string; content: string };
}

export interface Mailer {
  send(mail: Mail): Promise<void>;
}

/** Development: prints mail instead of sending it. */
export class LogMailer implements Mailer {
  readonly sent: Mail[] = [];
  async send(mail: Mail): Promise<void> {
    this.sent.push(mail);
    console.log(`[mail] to=${mail.to} subject=${JSON.stringify(mail.subject)}`);
  }
}

/** https://resend.com — set RESEND_API_KEY and MAIL_FROM. */
export class ResendMailer implements Mailer {
  private readonly apiKey: string;
  private readonly from: string;
  constructor(apiKey: string, from: string) {
    this.apiKey = apiKey;
    this.from = from;
  }
  async send(mail: Mail): Promise<void> {
    const body: Record<string, unknown> = { from: this.from, to: [mail.to], subject: mail.subject, text: mail.text };
    if (mail.attachment) {
      body.attachments = [{ filename: mail.attachment.filename, content: Buffer.from(mail.attachment.content, 'utf8').toString('base64') }];
    }
    const res = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: { authorization: `Bearer ${this.apiKey}`, 'content-type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!res.ok) throw new Error(`email provider returned ${res.status}: ${await res.text()}`);
  }
}

export function licenseEmail(to: string, licenses: { licensee: string; seats: number; updatesUntil: string; key: string; file: string }[]): Mail {
  const parts = licenses.map(
    (l, i) =>
      `${licenses.length > 1 ? `License ${i + 1}\n` : ''}Licensed to: ${l.licensee}\nSeats: ${l.seats}\nFree updates until: ${l.updatesUntil}\n\nYour license key:\n\n${l.key}\n`,
  );
  return {
    to,
    subject: licenses.length > 1 ? 'Your DeckForge license keys' : 'Your DeckForge license key',
    text:
      'Thank you for buying DeckForge.\n\n' +
      parts.join('\n') +
      '\nTo activate: open DeckForge, click "Trial" in the top-right corner, and paste the key.\n' +
      'The attached license file works too (handy for computers without internet access).\n\n' +
      'Keep this email: you can always get your key again from the license-recovery page.\n',
    ...(licenses[0] ? { attachment: { filename: 'deckforge-license.txt', content: licenses.map((l) => l.file).join('\n') } } : {}),
  };
}
