// Report Student Absence form in The Kings Corner (see worker/src/absences.js). The service checks every report
// again, saves it on the Absences tab of the office workbook and emails the parent, the office and the teacher.
(function () {
  var f = document.getElementById("ab-form"), status = document.getElementById("ab-status");
  if (!f || !window.fetch) return;
  f.hidden = false;
  function show(cls, html) { status.className = "notice " + cls; status.innerHTML = html; status.scrollIntoView({block: "center"}); }
  // A missing or wrong field gets a red border and its own line under it until it is fixed.
  function mark(el, msg) {
    var label = el.closest("label"), note = label && label.querySelector(".field-err");
    el.setAttribute("aria-invalid", msg ? "true" : "false");
    if (msg && !note && label) {
      note = document.createElement("span");
      note.className = "field-err";
      note.id = "ab-err-" + el.name;
      label.appendChild(note);
      el.setAttribute("aria-describedby", note.id);
    }
    if (note) { note.textContent = msg; note.hidden = !msg; }
  }
  f.addEventListener("input", function (e) { if (e.target.getAttribute("aria-invalid") === "true") mark(e.target, ""); });
  f.addEventListener("change", function (e) { if (e.target.getAttribute("aria-invalid") === "true") mark(e.target, ""); });
  f.addEventListener("submit", function (e) {
    e.preventDefault();
    var b = {};
    Array.prototype.forEach.call(f.elements, function (el) { if (el.name) b[el.name] = el.value.trim(); });
    var bad = [];
    Array.prototype.forEach.call(f.querySelectorAll("input[name], select[name], textarea[name]"), function (el) {
      if (el.name === "website") return;
      var msg = "";
      if (el.required && !el.value.trim()) msg = "Please fill this in.";
      else if (el.name === "details" && b.reason === "Other" && !b.details) msg = "Please tell us the reason.";
      else if (el.type === "email" && el.value && !el.checkValidity()) msg = "Please enter an email address like name@example.com.";
      else if (!el.checkValidity()) msg = el.validationMessage || "Please check this.";
      mark(el, msg);
      if (msg) bad.push(el);
    });
    if (bad.length) {
      show("err", bad.length === 1 ? "Please fix the field marked in red." : "Please fix the " + bad.length + " fields marked in red.");
      bad[0].focus();
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
