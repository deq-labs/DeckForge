// Wires the configurable links in config.js into the page. No tracking, no third-party scripts.
(function () {
  var cfg = window.DECKFORGE_SITE || {};

  document.querySelectorAll('[data-link]').forEach(function (el) {
    var url = cfg[el.getAttribute('data-link')];
    if (url) {
      el.setAttribute('href', url);
    } else {
      el.setAttribute('href', '#signup');
      if (el.hasAttribute('data-fallback')) el.textContent = el.getAttribute('data-fallback');
    }
  });

  document.querySelectorAll('[data-email]').forEach(function (el) {
    if (cfg.supportEmail) {
      el.setAttribute('href', 'mailto:' + cfg.supportEmail);
      if (!el.textContent.trim()) el.textContent = cfg.supportEmail;
    }
  });

  var video = document.getElementById('demo-video');
  if (video && cfg.demoVideo) {
    var frame = document.createElement('iframe');
    frame.src = cfg.demoVideo;
    frame.title = 'DeckForge in two minutes';
    frame.allow = 'fullscreen';
    frame.style.cssText = 'width:100%;aspect-ratio:16/9;border:0;display:block';
    video.replaceChildren(frame);
  }

  var form = document.getElementById('signup-form');
  if (form) {
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var email = form.elements.email.value.trim();
      var status = document.getElementById('signup-status');
      if (!email) return;
      if (!cfg.signupEndpoint) {
        if (cfg.supportEmail) window.location.href = 'mailto:' + cfg.supportEmail + '?subject=DeckForge%20early%20access&body=' + encodeURIComponent(email);
        else status.textContent = 'Signup opens soon. Please check back.';
        return;
      }
      var body = new FormData();
      body.append('email', email);
      fetch(cfg.signupEndpoint, { method: 'POST', body: body, mode: 'no-cors' })
        .then(function () { status.textContent = "Thanks! We'll email you when DeckForge is ready, and nothing else."; form.reset(); })
        .catch(function () { status.textContent = 'That did not go through. Please try again.'; });
    });
  }
})();
