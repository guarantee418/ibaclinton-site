// Feedback for Administration form on the Parent Portal (see worker/src/feedback.js). The service checks every
// response again and saves it on the Feedback tab of the office workbook; no email is sent.
(function () {
  var f = document.getElementById("fb-form"), status = document.getElementById("fb-status");
  if (!f || !window.fetch) return;
  f.hidden = false;
  function show(cls, html) { status.className = "notice " + cls; status.innerHTML = html; status.scrollIntoView({block: "center"}); }
  f.addEventListener("submit", function (e) {
    e.preventDefault();
    var missing = Array.prototype.filter.call(f.querySelectorAll("[required]"), function (el) { return !el.value.trim(); });
    if (missing.length || !f.checkValidity()) {
      show("err", "Please choose a topic and write your feedback" + (f.elements.email.validity.valid ? "." : ", and check your email address (or leave it blank)."));
      (missing[0] || f.querySelector(":invalid")).focus();
      return;
    }
    var b = {};
    Array.prototype.forEach.call(f.elements, function (el) { if (el.name) b[el.name] = el.value.trim(); });
    var btn = f.querySelector("button[type=submit]");
    btn.disabled = true; btn.textContent = "Sending…";
    fetch(f.getAttribute("data-api") + "/feedback", {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(b)})
      .then(function (r) { return r.json(); })
      .then(function (d) {
        if (!d.ok) throw new Error(d.error || "");
        f.reset(); f.hidden = true;
        show("ok", "<b>Thank you!</b> Your feedback was sent to the administration." +
          (b.email || b.phone ? " We’ll be in touch if a reply is needed." : ""));
      })
      .catch(function (err) {
        btn.disabled = false; btn.textContent = "Send feedback";
        var msg = String(err.message || "").replace(/[<>&]/g, "");
        show("err", "Sorry, your feedback didn’t go through" + (msg ? " (" + msg + ")" : "") +
          ". Please try again, or email <a href=\"mailto:iba@ibcministries.org\">iba@ibcministries.org</a>.");
      });
  });
})();
