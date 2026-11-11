// Generates the legal pages (refund, terms, EULA, privacy) from the texts below, sharing one
// layout. These are DRAFTS written from the product plan, not legal advice: have them reviewed
// for DEQ Labs LLC before launch, then remove the draft banner (DRAFT = false).
// Run: node site/legal/build-legal.mjs
import { writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const DRAFT = true;
const UPDATED = '2026-11-11';
const out = join(dirname(fileURLToPath(import.meta.url)), '..');

const pages = {
  'refund.html': {
    title: 'Refund policy',
    body: `
<p>If DeckForge doesn't work for you, you can have your money back.</p>
<h2>30 days, no questions</h2>
<p>Ask for a refund within 30 days of purchase and we'll refund the full price. Reply to your
receipt email or write to us; you don't need to give a reason.</p>
<h2>After a refund</h2>
<p>The license key is deactivated for new activations. Please remove DeckForge from your
computers; the app returns to trial mode when the license is removed.</p>
<h2>Try first</h2>
<p>The free trial has every feature; it only adds a small "Made with DeckForge" note to each
slide. We encourage you to test it on your own template before buying.</p>
<p>Purchases are processed by our reseller (Merchant of Record), whose refund process we use.</p>`,
  },
  'terms.html': {
    title: 'Terms of sale',
    body: `
<p>These terms cover buying DeckForge from DEQ Labs LLC ("we"). Using the software is covered by
the <a href="eula.html">End User License Agreement</a>.</p>
<h2>What you buy</h2>
<p>A perpetual license to use the version of DeckForge Lite you buy, and any version released
during the 12 months of updates included with your purchase, on the number of computers in your
plan (Single user: 1; Team: 5).</p>
<h2>Price, payment and taxes</h2>
<p>Prices are in USD. Payment, invoicing and sales taxes are handled by our reseller, who acts as
the Merchant of Record.</p>
<h2>Delivery</h2>
<p>Your license key is emailed right after payment. You can get it again any time from the
license-recovery page.</p>
<h2>Updates and support</h2>
<p>Updates are included for 12 months from purchase. After that, the versions you already have
keep working; renewing updates is optional. Support is by email.</p>
<h2>Refunds</h2>
<p>See the <a href="refund.html">refund policy</a> (30 days).</p>
<h2>Liability</h2>
<p>To the extent the law allows, our total liability is limited to the amount you paid. Always
review generated decks before you share them.</p>
<h2>Contact</h2>
<p>DEQ Labs LLC · <a data-email href="#">contact us by email</a></p>`,
  },
  'eula.html': {
    title: 'End User License Agreement',
    body: `
<p>This agreement is between you and DEQ Labs LLC for DeckForge Lite ("the software").</p>
<h2>1. License</h2>
<p>We grant you a non-exclusive, non-transferable, perpetual license to install and use the
software on the number of computers your license allows, for your own or your organization's
business. The license covers the versions released on or before the end of your included update
period. Unlicensed (trial) use is free and adds a notice to generated slides.</p>
<h2>2. Your content</h2>
<p>Your templates, workbooks and generated decks are yours. The software processes them on your
computer and does not send them to us.</p>
<h2>3. Restrictions</h2>
<p>You may not resell, sublicense or redistribute the software or license keys, or remove or work
around the license checks.</p>
<h2>4. Network use</h2>
<p>The software works offline. It contacts our servers only when you choose to register a
computer for your license or to check for updates, and sends only your license key, an anonymous
computer identifier and the app version.</p>
<h2>5. Warranty disclaimer</h2>
<p>The software is provided "as is", without warranties of any kind, to the extent the law allows.
Review generated decks before you rely on or distribute them.</p>
<h2>6. Limitation of liability</h2>
<p>To the extent the law allows, we are not liable for indirect or consequential damages, and our
total liability is limited to the price you paid for the license.</p>
<h2>7. Termination</h2>
<p>This license ends if you breach it. Refunded licenses end on refund.</p>
<h2>8. Third-party components</h2>
<p>The software includes open-source components under their own licenses, listed in the
application's notices.</p>
<h2>9. Governing law</h2>
<p>[To be completed: the governing law and venue for DEQ Labs LLC.]</p>`,
  },
  'privacy.html': {
    title: 'Privacy policy',
    body: `
<p>Short version: DeckForge runs on your computer, and your files never leave it.</p>
<h2>The app</h2>
<ul>
<li>Templates, workbooks and generated decks stay on your computer. We never receive them.</li>
<li>No analytics, tracking or telemetry.</li>
<li>Optional network calls: registering a computer for your license (sends your license key and an
anonymous computer identifier: a one-way hash, not your name or hardware details) and checking
for updates (sends the app version). Both can be left off.</li>
</ul>
<h2>When you buy</h2>
<p>Our reseller processes your payment and shares your name and email with us so we can send your
license key, support you and let you recover the key. We keep these records for as long as your
license exists, and for tax and accounting periods the law requires.</p>
<h2>Email list</h2>
<p>If you sign up for early access, we use your email only to tell you when DeckForge is available.
Unsubscribe any time; we never sell or share the list.</p>
<h2>This website</h2>
<p>No cookies and no third-party trackers.</p>
<h2>Your rights</h2>
<p>Ask us to see, correct or delete your data at any time by email.</p>`,
  },
};

const layout = (title, body) => `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${title} · DeckForge</title>
  <link rel="stylesheet" href="styles.css">
</head>
<body>
  <header class="top">
    <div class="wrap">
      <a class="brand" href="./"><span>◆</span> DeckForge</a>
      <nav><a href="./#how">How it works</a><a href="pricing.html">Pricing</a><a href="./#faq">FAQ</a></nav>
    </div>
  </header>
  <main class="wrap legal">
    <h1>${title}</h1>
    ${DRAFT ? '<div class="draft"><strong>Draft:</strong> not yet reviewed by a lawyer. Not in effect.</div>' : ''}
    <p class="small">Last updated ${UPDATED}</p>
    ${body.trim()}
  </main>
  <footer>
    <div class="wrap">
      <span>© 2026 DEQ Labs LLC</span>
      <span><a href="refund.html">Refunds</a> · <a href="terms.html">Terms</a> · <a href="eula.html">EULA</a> · <a href="privacy.html">Privacy</a></span>
    </div>
  </footer>
  <script src="config.js"></script>
  <script src="site.js"></script>
</body>
</html>
`;

for (const [file, p] of Object.entries(pages)) {
  writeFileSync(join(out, file), layout(p.title, p.body));
  console.log(`wrote site/${file}`);
}
