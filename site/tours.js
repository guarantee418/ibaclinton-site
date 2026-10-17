// School tour sign-up (see tours.py and worker/). Tours, spots and deadlines come from
// tours/schedule.json and the sign-up counts from the sign-up service, which checks everything
// again, so nothing here is trusted.
(function () {
  var app = document.getElementById("tour-app");
  if (!app || !window.fetch) return;
  var api = app.getAttribute("data-api");
  var list = app.querySelector(".tour-list");
  var panel = document.getElementById("tour-panel");
  var schedule = null, taken = null, loadError = false, done = {}, open = null;

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function left(t) { return Math.max(t.spots - ((taken && taken[t.id]) || 0), 0); }
  function closed(t) { return Date.now() > Date.parse(t.deadline); }

  function loadTaken() {
    return fetch(api + "/tours", { cache: "no-store" }).then(function (r) { return r.json(); }).then(function (d) {
      if (!d.ok) throw new Error(d.error || "unavailable");
      taken = d.taken || {}; loadError = false;
    }).catch(function () { loadError = true; });
  }

  function render(message) {
    var tours = schedule.tours.filter(function (t) { return Date.now() < Date.parse(t.starts); });
    if (!tours.length) {
      list.innerHTML = "";
      panel.innerHTML = '<p class="notice">No tours are scheduled right now. Please call the office at <a href="tel:3018561616">(301) 856-1616</a> to set up a visit.</p>';
      return;
    }
    list.innerHTML = tours.map(function (t) {
      var n = left(t), status, button = "";
      if (done[t.id]) status = '<span class="tour-status ok">You’re signed up</span>';
      else if (loadError) status = '<span class="tour-status">Open spots unknown</span>';
      else if (!n) status = '<span class="tour-status full">Full</span>';
      else if (closed(t)) status = '<span class="tour-status full">Sign-up closed</span>';
      else {
        status = '<span class="tour-status">' + n + " of " + t.spots + " spots open</span>";
        button = '<button type="button" class="btn tour-pick" data-id="' + esc(t.id) + '">Sign up</button>';
      }
      return '<li class="tour-card' + (open === t.id ? " open" : "") + '"><div><b>' + esc(t.label) + "</b>" +
        (t.note ? "<span>" + esc(t.note) + "</span>" : "") + status + "</div>" + button + "</li>";
    }).join("");
    var t = open && tours.filter(function (x) { return x.id === open; })[0];
    if (loadError) {
      panel.innerHTML = '<div class="notice err">We couldn’t load the open spots. Please reload the page in a minute, or call the office at <a href="tel:3018561616">(301) 856-1616</a>.</div>';
    } else if (t && done[t.id]) {
      panel.innerHTML = '<div class="notice ok" role="status"><b>You’re signed up</b> for the tour on ' + esc(t.label) +
        ". A confirmation email is on its way" + (done[t.id].email_sent ? "" : " (if it doesn’t arrive, please call the office)") +
        '. Please check in at the school office when you arrive. To change or cancel, call <a href="tel:3018561616">(301) 856-1616</a>.</div>';
    } else if (t) {
      panel.innerHTML = (message || "") + form(t);
      wire(t);
    } else {
      panel.innerHTML = message || "";
    }
  }

  function form(t) {
    return '<form class="info tour-form" id="tour-form" novalidate>' +
      "<h3>Sign up for " + esc(t.label) + "</h3>" +
      "<fieldset><legend>Parent/Guardian</legend>" +
      '<label>First name<input name="parent_first" autocomplete="given-name" maxlength="60" required></label>' +
      '<label>Last name<input name="parent_last" autocomplete="family-name" maxlength="60" required></label>' +
      '<label>Email<input type="email" name="email" autocomplete="email" maxlength="120" required></label>' +
      '<label>Phone<input type="tel" name="phone" autocomplete="tel" maxlength="30" required></label>' +
      "</fieldset><fieldset><legend>Your children</legend>" +
      '<label class="full">Names and the grades they’d start in<textarea name="children" maxlength="400" rows="2" required placeholder="For example: Sam, 3rd grade; Ally, K4"></textarea></label>' +
      '<label class="full">Anything you’d like us to know? <span class="opt">(optional)</span><textarea name="notes" maxlength="500" rows="2"></textarea></label>' +
      "</fieldset>" +
      '<label class="hp" aria-hidden="true">Leave this field blank<input name="website" tabindex="-1" autocomplete="off"></label>' +
      '<div id="tour-status" role="status" aria-live="polite"></div>' +
      '<p><button class="btn" type="submit">Sign up</button></p>' +
      '<p class="lunch-fine">You’ll get a confirmation email. Each tour takes ' + t.spots + " families.</p>" +
      "</form>";
  }

  function wire(t) {
    var f = document.getElementById("tour-form");
    var status = document.getElementById("tour-status");
    function show(cls, html) { status.className = "notice " + cls; status.innerHTML = html; }
    f.addEventListener("input", function () { if (status.classList.contains("err")) { status.className = ""; status.innerHTML = ""; } });
    f.addEventListener("submit", function (e) {
      e.preventDefault();
      var el = f.elements;
      var b = {
        tour: t.id, parent_first: el.parent_first.value.trim(), parent_last: el.parent_last.value.trim(),
        email: el.email.value.trim(), phone: el.phone.value.trim(), children: el.children.value.trim(),
        notes: el.notes.value.trim(), website: el.website.value
      };
      var problems = [];
      if (!b.parent_first || !b.parent_last) problems.push("your first and last name");
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(b.email)) problems.push("a valid email address");
      if (!b.phone) problems.push("a phone number");
      if (!b.children) problems.push("your children’s names and grades");
      if (problems.length) { show("err", "Please enter " + problems.join(", ") + "."); return; }
      var btn = f.querySelector("button[type=submit]");
      btn.disabled = true; btn.textContent = "Signing up…";
      fetch(api + "/tour", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(b) })
        .then(function (r) { return r.json(); })
        .then(function (d) {
          if (d.ok) {
            done[t.id] = { email_sent: d.email_sent };
            taken[t.id] = (taken[t.id] || 0) + 1;
            render();
            panel.scrollIntoView({ block: "start" });
          } else if (d.code === "full" || d.code === "busy") {
            loadTaken().then(function () { render('<div class="notice err">' + esc(d.error) + "</div>"); });
          } else {
            show("err", esc(d.error || "Something went wrong. Please try again or call the office."));
            btn.disabled = false; btn.textContent = "Sign up";
          }
        })
        .catch(function () {
          show("err", "We couldn’t reach the sign-up service. Please check your connection and try again, or call the office at (301) 856-1616.");
          btn.disabled = false; btn.textContent = "Sign up";
        });
    });
  }

  list.addEventListener("click", function (e) {
    var b = e.target.closest(".tour-pick");
    if (!b) return;
    open = b.getAttribute("data-id");
    render();
    panel.scrollIntoView({ block: "start", behavior: "smooth" });
    var first = document.querySelector("#tour-form [name=parent_first]");
    if (first) first.focus({ preventScroll: true });
  });

  Promise.all([
    fetch(app.getAttribute("data-schedule"), { cache: "no-store" }).then(function (r) { return r.json(); }).then(function (s) { schedule = s; }),
    loadTaken()
  ]).then(function () { render(); }).catch(function () { /* the plain list of tours stays */ });
})();
