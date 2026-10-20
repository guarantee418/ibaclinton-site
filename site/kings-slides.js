// The Kings Corner's purple banner as a slider: the welcome slide, then the office's announcements
// (content/parents.json). Adds the arrows and dots, turns every 7 seconds (not while the mouse or keyboard is on it,
// nor for visitors who ask for less motion), and hides announcements outside their show dates (Eastern time),
// because the page is only rebuilt when something is saved. Without JavaScript the slides swipe sideways.
(function () {
  var slider = document.querySelector(".kc-slider");
  if (!slider) return;
  var track = slider.querySelector(".kc-track");
  var today = new Date().toLocaleDateString("en-CA", { timeZone: "America/New_York" }); // YYYY-MM-DD
  Array.prototype.forEach.call(track.children, function (li) {
    var from = li.getAttribute("data-from"), until = li.getAttribute("data-until");
    if ((from && from > today) || (until && until < today)) li.remove();
  });
  var slides = Array.prototype.slice.call(track.children);
  if (slides.length < 2) { slider.classList.add("kc-single"); return; }

  var svg = function (d) {
    return '<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" ' +
      'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="' + d + '"/></svg>';
  };
  function button(cls, label, html) {
    var b = document.createElement("button");
    b.type = "button";
    b.className = cls;
    b.setAttribute("aria-label", label);
    b.innerHTML = html;
    return b;
  }
  var prev = button("kc-nav kc-prev", "Previous slide", svg("M15 5l-7 7 7 7"));
  var next = button("kc-nav kc-next", "Next slide", svg("M9 5l7 7-7 7"));
  var dots = document.createElement("div");
  dots.className = "kc-dots";
  slides.forEach(function (s, i) {
    s.setAttribute("aria-label", (i + 1) + " of " + slides.length);
    var d = button("kc-dot", "Show slide " + (i + 1), "");
    d.addEventListener("click", function () { go(i); });
    dots.appendChild(d);
  });
  slider.appendChild(prev);
  slider.appendChild(next);
  slider.appendChild(dots);
  slider.classList.add("kc-slider-js");

  var current = 0;
  function mark(i) {
    current = i;
    slides.forEach(function (s, j) {
      s.toggleAttribute("inert", j !== i); // links on other slides are skipped by Tab
      dots.children[j].setAttribute("aria-current", j === i ? "true" : "false");
    });
  }
  function go(i) {
    i = (i + slides.length) % slides.length;
    track.scrollTo({ left: slides[i].offsetLeft - track.offsetLeft, behavior: "smooth" });
    mark(i);
  }
  prev.addEventListener("click", function () { go(current - 1); restart(); });
  next.addEventListener("click", function () { go(current + 1); restart(); });
  slider.addEventListener("keydown", function (e) {
    if (e.key === "ArrowLeft") { go(current - 1); restart(); }
    if (e.key === "ArrowRight") { go(current + 1); restart(); }
  });
  var settle;
  track.addEventListener("scroll", function () { // swiping on a phone
    clearTimeout(settle);
    settle = setTimeout(function () {
      var i = Math.round(track.scrollLeft / track.clientWidth);
      if (i !== current && slides[i]) mark(i);
    }, 120);
  }, { passive: true });
  window.addEventListener("resize", function () { track.scrollLeft = slides[current].offsetLeft - track.offsetLeft; });

  var timer = null, paused = false;
  var still = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  function restart() {
    clearInterval(timer);
    if (!still && !paused) timer = setInterval(function () { go(current + 1); }, 7000);
  }
  function pause(on) { paused = on; restart(); }
  slider.addEventListener("mouseenter", function () { pause(true); });
  slider.addEventListener("mouseleave", function () { pause(false); });
  slider.addEventListener("focusin", function () { pause(true); });
  slider.addEventListener("focusout", function (e) { if (!slider.contains(e.relatedTarget)) pause(false); });
  track.addEventListener("touchstart", function () { pause(true); }, { passive: true });
  document.addEventListener("visibilitychange", function () { if (document.hidden) clearInterval(timer); else restart(); });
  mark(0);
  restart();
})();
