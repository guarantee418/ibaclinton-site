// Parent-teacher conference sign-up (see conferences.py and worker/). Teachers, times and the
// deadline come from conferences/schedule.json and the booked times from the booking service,
// which checks everything again, so nothing here is trusted.
(function () {
  var app = document.getElementById("conf-app");
  if (!app || !window.fetch) return;
  var api = app.getAttribute("data-api");
  var pick = document.getElementById("conf-pick");
  var list = app.querySelector(".conf-teachers");
  var panel = document.getElementById("conf-panel");
  var schedule = null, taken = null, loadError = false, booked = {};

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function closed() { return !schedule || !schedule.deadline || Date.now() > Date.parse(schedule.deadline); }
  function teacher(id) {
    for (var i = 0; schedule && i < schedule.teachers.length; i++) if (schedule.teachers[i].id === id) return schedule.teachers[i];
    return null;
  }
  function isTaken(tid, sid) { return !!(taken && taken[tid] && taken[tid].indexOf(sid) >= 0); }

  function loadTaken() {
    return fetch(api + "/conferences", { cache: "no-store" }).then(function (r) { return r.json(); }).then(function (d) {
      if (!d.ok) throw new Error(d.error || "unavailable");
      taken = d.taken || {}; loadError = false;
    }).catch(function () { loadError = true; });
  }

  function route() {
    var id = decodeURIComponent(location.hash.replace(/^#/, ""));
    var t = schedule && teacher(id);
    pick.hidden = list.hidden = !!t;
    panel.hidden = !t;
    if (t) { render(t); panel.scrollIntoView({ block: "start" }); }
  }

  function render(t, message) {
    var free = 0, slots = schedule.slots.map(function (s) {
      var gone = isTaken(t.id, s.id);
      if (!gone) free++;
      return '<button type="button" class="conf-slot" data-slot="' + esc(s.id) + '"' +
        (gone || closed() || loadError ? " disabled" : "") + ' aria-pressed="false">' +
        esc(s.label) + "<small>" + (gone ? "Booked" : "Open") + "</small></button>";
    }).join("");
    var mine = booked[t.id];
    var html = '<p><a href="#" class="conf-back">&larr; All teachers</a></p>' +
      "<h2>" + esc(t.name) + "</h2>" + (t.subject ? '<p class="conf-subject">' + esc(t.subject) + "</p>" : "") +
      (message || "");
    if (mine) {
      html += '<div class="notice ok" role="status"><b>You’re signed up.</b> ' + esc(t.name) + ", " +
        esc(schedule.date_label) + " at " + esc(mine.time) + ". A confirmation email is on its way" +
        (mine.email_sent ? "" : " (if it doesn’t arrive, please call the office)") +
        ". To change or cancel, please call the office at <a href=\"tel:3018561616\">(301) 856-1616</a>.</div>";
    } else if (loadError) {
      html += '<div class="notice err">We couldn’t load the open times. Please reload the page in a minute, or call the office at <a href="tel:3018561616">(301) 856-1616</a>.</div>';
    } else if (closed()) {
      html += '<div class="notice">Online sign-up is closed. To book a time, please call the office at <a href="tel:3018561616">(301) 856-1616</a>.</div>';
    } else if (!free) {
      html += '<div class="notice">All of ' + esc(t.name) + '’s times are booked. Please call the office at <a href="tel:3018561616">(301) 856-1616</a>.</div>';
    } else {
      html += '<p class="conf-count">' + free + " of " + schedule.slots.length + " times open, " + schedule.minutes + " minutes each</p>";
    }
    html += '<div class="conf-slots" role="group" aria-label="Times">' + slots + "</div>";
    if (!mine) html +=
      '<form class="info conf-form" id="conf-form" novalidate hidden>' +
      '<h3>Book <span id="conf-time"></span> with ' + esc(t.name) + "</h3>" +
      "<fieldset><legend>Parent/Guardian</legend>" +
      '<label>First name<input name="parent_first" autocomplete="given-name" maxlength="60" required></label>' +
      '<label>Last name<input name="parent_last" autocomplete="family-name" maxlength="60" required></label>' +
      '<label>Email<input type="email" name="email" autocomplete="email" maxlength="120" required></label>' +
      '<label>Phone<input type="tel" name="phone" autocomplete="tel" maxlength="30" required></label>' +
      "</fieldset><fieldset><legend>Student</legend>" +
      '<label class="full">Student’s full name<input name="student" autocomplete="off" maxlength="80" required></label>' +
      '<label class="full">Anything you’d like to talk about? <span class="opt">(optional)</span><textarea name="notes" maxlength="500" rows="3"></textarea></label>' +
      "</fieldset>" +
      '<label class="hp" aria-hidden="true">Leave this field blank<input name="website" tabindex="-1" autocomplete="off"></label>' +
      '<div id="conf-status" role="status" aria-live="polite"></div>' +
      '<p><button class="btn" type="submit">Book this time</button></p>' +
      '<p class="lunch-fine">You’ll get a confirmation email. To change or cancel after booking, please call the office.</p>' +
      "</form>";
    panel.innerHTML = html;
    wire(t);
  }

  function wire(t) {
    var form = document.getElementById("conf-form"), chosen = null;
    if (!form) return;
    var status = document.getElementById("conf-status");
    function show(cls, html) { status.className = "notice " + cls; status.innerHTML = html; }
    panel.querySelector(".conf-slots").addEventListener("click", function (e) {
      var b = e.target.closest(".conf-slot");
      if (!b || b.disabled) return;
      chosen = b.getAttribute("data-slot");
      Array.prototype.forEach.call(panel.querySelectorAll(".conf-slot"), function (x) { x.setAttribute("aria-pressed", String(x === b)); });
      document.getElementById("conf-time").textContent = b.firstChild.nodeValue;
      form.hidden = false;
      form.scrollIntoView({ block: "start", behavior: "smooth" });
      form.elements.parent_first.focus({ preventScroll: true });
    });
    form.addEventListener("input", function () { if (status.classList.contains("err")) { status.className = ""; status.innerHTML = ""; } });
    form.addEventListener("submit", function (e) {
      e.preventDefault();
      var f = form.elements;
      var b = {
        teacher: t.id, slot: chosen, date: schedule.date,
        parent_first: f.parent_first.value.trim(), parent_last: f.parent_last.value.trim(),
        email: f.email.value.trim(), phone: f.phone.value.trim(), student: f.student.value.trim(),
        notes: f.notes.value.trim(), website: f.website.value
      };
      var problems = [];
      if (!b.parent_first || !b.parent_last) problems.push("your first and last name");
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(b.email)) problems.push("a valid email address");
      if (!b.phone) problems.push("a phone number");
      if (!b.student) problems.push("your student’s name");
      if (problems.length) { show("err", "Please enter " + problems.join(", ") + "."); return; }
      var btn = form.querySelector("button[type=submit]");
      btn.disabled = true; btn.textContent = "Booking…";
      fetch(api + "/conference", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(b) })
        .then(function (r) { return r.json(); })
        .then(function (d) {
          if (d.ok) {
            booked[t.id] = { time: d.time, email_sent: d.email_sent };
            (taken[t.id] = taken[t.id] || []).push(chosen);
            render(t);
            panel.scrollIntoView({ block: "start" });
          } else if (d.code === "taken") {
            loadTaken().then(function () {
              render(t, '<div class="notice err">' + esc(d.error) + "</div>");
            });
          } else {
            show("err", esc(d.error || "Something went wrong. Please try again or call the office."));
            btn.disabled = false; btn.textContent = "Book this time";
          }
        })
        .catch(function () {
          show("err", "We couldn’t reach the sign-up service. Please check your connection and try again, or call the office at (301) 856-1616.");
          btn.disabled = false; btn.textContent = "Book this time";
        });
    });
  }

  panel.addEventListener("click", function (e) {
    if (e.target.classList.contains("conf-back")) { e.preventDefault(); history.pushState(null, "", location.pathname); route(); pick.scrollIntoView({ block: "start" }); }
  });
  window.addEventListener("hashchange", route);
  window.addEventListener("popstate", route);

  Promise.all([
    fetch(app.getAttribute("data-schedule"), { cache: "no-store" }).then(function (r) { return r.json(); }).then(function (s) { schedule = s; }),
    loadTaken()
  ]).then(route).catch(function () { /* the plain list of teachers stays */ });
})();
