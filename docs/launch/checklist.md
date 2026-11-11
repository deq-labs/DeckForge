# Lite launch checklist (L6)

The code for launch is in place (installers, licensing, packs, docs, `site/`). What's left is
mostly yours. In rough order:

## Before the landing page goes up (validation, spec §Launch validation)
- [ ] Host `site/` (Cloudflare Pages / Netlify / any static host). Set values in `site/config.js`:
      `signupEndpoint` (Buttondown, ConvertKit…), `supportEmail`, `checkoutPreorder`.
- [ ] Record the 2-minute demo and the 5-minute getting-started video (`docs/user/video-script.md`); set `demoVideo`.
- [ ] Post in 5 communities (drafts below). Show before/after, not a pitch.
- [ ] Go / no-go after 6 weeks: 100+ signups or 5+ pre-orders; 2+ services jobs; 3 people describing the same deck.

## Before taking money
- [ ] DEQ Labs LLC formed; Lemon Squeezy store (Merchant of Record) set up — `docs/licensing-operations.md`.
- [ ] Production key pair generated; public key in `packages/licensing/src/config.ts`.
- [ ] License server deployed; `LICENSE_SERVER_URL` set; end-to-end test purchase in test mode.
- [ ] Lawyer review of `site/legal/build-legal.mjs` texts (refund, terms, EULA, privacy); set
      `DRAFT = false`, regenerate; copy the final EULA into `apps/desktop/build/eula.txt`.
- [ ] Code signing (Windows) + Apple notarization; update host; app icon — `docs/release.md`.
- [ ] Two people who aren't you: download → first deck in under 30 minutes.

## Launch
- [ ] Tag `v1.0.0`; review the draft release; upload installers + `latest*.yml`; set download links in `site/config.js`.
- [ ] 1 case study from a services job (before/after time, with permission).
- [ ] First 10 paid licenses.

## Community post drafts

**r/manufacturing / LinkedIn ops groups**
> Every month I rebuilt the same ops review deck: export from the MES, paste into Excel, copy
> 30 numbers and four charts into PowerPoint, fix what I broke. I built a small desktop tool that
> fills our own PowerPoint template from the Excel file instead — native charts, tables that
> grow, and one deck per plant from one workbook. Nothing goes to the cloud (a must for us).
> Before/after screenshots below. Curious how others handle the monthly deck — and what breaks.

**r/PowerPoint / r/excel**
> Tip for anyone rebuilding the same report deck monthly: tag your template (`{{oee}}`, name
> charts in the Selection Pane) and fill it from a workbook. I made a tool for this
> (DeckForge) — here's a 2-minute demo of Excel → 20 client decks. Happy to answer how the
> chart data stays editable.
