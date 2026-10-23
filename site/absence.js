// Report Student Absence form in The King's Court (see worker/src/absences.js). The service checks every report
// again, saves it on the Absences tab of the office workbook and emails the parent, the office and the teacher.
(function () {
  var f = document.getElementById("ab-form"), status = document.getElementById("ab-status");
  if (!f || !window.fetch) return;
  f.hidden = false;
  function show(cls, html) { status.className = "notice " + cls; status.innerHTML = html; status.scrollIntoView({block: "center"}); }
  f.addEventListener("submit", function (e) {
    e.preventDefault();
    var b = {};
    Array.prototype.forEach.call(f.elements, function (el) { if (el.name) b[el.name] = el.value.trim(); });
    var missing = Array.prototype.filter.call(f.querySelectorAll("[required]"), function (el) { return !el.value.trim(); });
    if (b.reason === "Other" && !b.details) missing.push(f.elements.details);
    if (missing.length || !f.checkValidity()) {
      show("err", "Please fill in every field (the last day and details are only needed when they apply).");
      (missing[0] || f.querySelector(":invalid")).focus();
      return;
    }
    var btn = f.querySelector("button[type=submit]");
    btn.disabled = true; btn.textContent = "Sending…";
    fetch(f.getAttribute("data-api") + "/absence", {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(b)})
      .then(function (r) { return r.json(); })
      .then(function (d) {
        if (!d.ok) throw new Error(d.error || "");
        f.reset(); f.hidden = true;
        show("ok", "<b>Thank you!</b> The absence was reported to the school office and your child’s teacher. " +
          "A confirmation email is on its way" + (d.email_sent ? "." : " (if it doesn’t arrive, please call the office)."));
      })
      .catch(function (err) {
        btn.disabled = false; btn.textContent = "Report absence";
        var msg = String(err.message || "").replace(/[<>&]/g, "");
        show("err", "Sorry, your report didn’t go through" + (msg ? " (" + msg + ")" : "") +
          ". Please try again, or call <a href=\"tel:3018561616\">(301) 856-1616</a>.");
      });
  });
})();
