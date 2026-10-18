// Transcript request form (see worker/src/transcripts.js). The fees and the payment link come from the
// request service, which checks every request again, so nothing here is trusted.
(function () {
  var f = document.getElementById("tr-form");
  if (!f || !window.fetch) return;
  var api = f.getAttribute("data-api");
  var status = document.getElementById("tr-status");
  var done = document.getElementById("tr-done");
  var fees = null;
  f.hidden = false;

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function money(n) { return "$" + Number(n).toFixed(2).replace(/\.00$/, ""); }
  function val(name) { var el = f.elements[name]; return el ? String(el.value || "").trim() : ""; }
  function who() { var c = f.querySelector("input[name=who]:checked"); return c ? c.value : ""; }
  function show(cls, html) { status.className = "notice " + cls; status.innerHTML = html; }
  function toggle(sel, on) {
    Array.prototype.forEach.call(f.querySelectorAll(sel), function (el) { el.hidden = !on; });
  }

  function feeLine() {
    var line = document.getElementById("tr-fee-line");
    if (!fees || !line) return;
    var parts = [];
    if (fees.fee_student) parts.push("Current students: " + money(fees.fee_student) + ", added to your Gradelink account.");
    if (fees.fee_alumni) parts.push("Alumni: " + money(fees.fee_alumni) + (fees.pay_url ? ", paid online after you submit." : ", collected by the school office."));
    line.textContent = parts.length ? parts.join(" ") + " We send the transcript once the fee is paid." : "There is no fee for transcripts.";
  }

  function update() {
    var w = who(), d = val("delivery");
    toggle(".tr-rest", !!w);
    toggle(".tr-student", w === "student");
    toggle(".tr-alumni", w === "alumni");
    toggle(".tr-parent", w === "student" && val("relation") === "parent");
    toggle(".tr-other", val("reason") === "Other");
    toggle(".tr-send", d === "mail" || d === "email");
    toggle(".tr-mail", d === "mail");
    toggle(".tr-email", d === "email");
    var consent = f.querySelector(".tr-consent-text");
    consent.textContent = (w === "student" && val("relation") === "parent")
      ? "I am this student’s parent or guardian, and I authorize Independent Baptist Academy to release this transcript as described above."
      : "I am the student named above, and I authorize Independent Baptist Academy to release my transcript as described above.";
  }

  f.addEventListener("change", update);
  f.addEventListener("input", function () { if (status.classList.contains("err")) { status.className = ""; status.innerHTML = ""; } });

  f.addEventListener("submit", function (e) {
    e.preventDefault();
    var w = who(), d = val("delivery"), parent = w === "student" && val("relation") === "parent";
    var b = { who: w, consent: f.elements.consent.checked, website: f.elements.website.value };
    ["student_first", "student_last", "former_name", "birth_date", "grade", "class_year", "relation", "requester_name",
      "email", "phone", "kind", "copies", "reason", "reason_other", "delivery", "send_to", "attention", "recipient_email",
      "street", "street2", "city", "state", "zip", "notes", "signature"].forEach(function (n) { b[n] = val(n); });
    var problems = [];
    if (!b.student_first || !b.student_last) problems.push("the student’s first and last name");
    if (!b.birth_date) problems.push("the date of birth");
    if (w === "student" && !b.grade) problems.push("the grade");
    if (w === "alumni" && !/^\d{4}$/.test(b.class_year)) problems.push("the year you graduated or last attended");
    if (parent && !b.requester_name) problems.push("your name");
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(b.email)) problems.push("a valid email address");
    if (!b.phone) problems.push("a phone number");
    if (!b.kind) problems.push("official or unofficial");
    if (!b.reason) problems.push("the reason");
    if (b.reason === "Other" && !b.reason_other) problems.push("the reason for the request");
    if (!d) problems.push("how to deliver it");
    if ((d === "mail" || d === "email") && !b.send_to) problems.push("who it goes to");
    if (d === "mail" && !(b.street && b.city && b.state && b.zip)) problems.push("the full mailing address");
    if (d === "email" && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(b.recipient_email)) problems.push("the recipient’s email address");
    if (!b.consent || !b.signature) problems.push("the release check box and your typed signature");
    if (problems.length) { show("err", "Please enter " + esc(problems.join(", ")) + "."); return; }

    var btn = f.querySelector("button[type=submit]");
    btn.disabled = true; btn.textContent = "Sending…";
    fetch(api + "/transcript", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(b) })
      .then(function (r) { return r.json(); })
      .then(function (r) {
        if (!r.ok) {
          show("err", esc(r.error || "Something went wrong. Please try again or call the office."));
          btn.disabled = false; btn.textContent = "Submit request";
          return;
        }
        var fee = w === "alumni" ? r.fee_alumni : r.fee_student, pay = "";
        if (fee && w === "student") pay = "<p>The " + money(fee) + " transcript fee will be added to your Gradelink account. We’ll send the transcript once it is paid.</p>";
        else if (fee && r.pay_url) pay = "<p>Please pay the " + money(fee) + " transcript fee now. Enter your request number, <b>" + esc(r.request_id) +
          "</b>, in the payment notes.</p><p><a class=\"btn\" href=\"" + esc(r.pay_url) + "\" target=\"_blank\" rel=\"noopener\">Pay " + money(fee) + "</a></p>";
        else if (fee) pay = "<p>There is a " + money(fee) + " transcript fee. The school office will contact you to collect it.</p>";
        f.hidden = true;
        done.innerHTML = '<div class="notice ok" role="status"><p><b>Thank you! We received your request.</b> Your request number is <b>' +
          esc(r.request_id) + "</b>.</p>" + pay + "<p>A confirmation email is on its way" +
          (r.email_sent ? "" : " (if it doesn’t arrive, please call the office)") +
          '. Questions? Call <a href="tel:3018561616">(301) 856-1616</a>.</p></div>';
        done.scrollIntoView({ block: "start" });
      })
      .catch(function () {
        show("err", "We couldn’t reach the request service. Please check your connection and try again, or call the office at (301) 856-1616.");
        btn.disabled = false; btn.textContent = "Submit request";
      });
  });

  update();
  fetch(api + "/transcripts", { cache: "no-store" }).then(function (r) { return r.json(); })
    .then(function (d) { if (d.ok) { fees = d; feeLine(); } }).catch(function () { /* the general fee line stays */ });
})();
