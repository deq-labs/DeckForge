// Site settings. Leave a value empty and the page falls back gracefully:
// - no checkout URLs: buy buttons scroll to the email signup instead;
// - no signup endpoint: the form opens an email to SUPPORT_EMAIL;
// - no video: the demo section shows a still instead.
window.DECKFORGE_SITE = {
  // Lemon Squeezy checkout links, e.g. https://<store>.lemonsqueezy.com/buy/<variant-id>
  checkoutSingle: '',
  checkoutTeam: '',
  // Pre-order (discounted) checkout, during validation.
  checkoutPreorder: '',
  // Email-list form endpoint (Buttondown, ConvertKit, ...) that accepts POST with an "email" field.
  signupEndpoint: '',
  // Public support address.
  supportEmail: '',
  // Demo video embed URL (e.g. a YouTube "embed" link).
  demoVideo: '',
  // Download links (filled in at launch).
  downloadWindows: '',
  downloadMac: '',
};
