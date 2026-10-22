// Forms the office adds to The Kings Corner sections (forms.py builds them; worker/src/forms.js saves them on the
// form's own tab of the office workbook and emails a confirmation). The service checks every answer again.
(function () {
  if (window.cfForms || !window.fetch) return;
  window.cfForms = true;
  Array.prototype.forEach.call(document.querySelectorAll("form.cf-form"), function (f) {
    var status = document.getElementById("cf-status-" + f.getAttribute("data-form"));
    f.hidden = false;
    function show(cls, html) { status.className = "cf-status notice " + cls; status.innerHTML = html; status.scrollIntoView({block: "center"}); }
    function labelOf(el) { return (el.getAttribute("data-label") || "").replace(/[<>&]/g, ""); } // goes into innerHTML
    f.addEventListener("submit", function (e) {
      e.preventDefault();
      var problems = [], first = null, answers = {};
      function need(el, what) { problems.push(what); first = first || el; }
      var name = f.elements.name.value.trim(), email = f.elements.email.value.trim();
      if (!name) need(f.elements.name, "your name");
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) need(f.elements.email, "a valid email address");
      Array.prototype.forEach.call(f.elements, function (el) {
        if (!/^q\d+$/.test(el.name || "")) return;
        if (el.type === "checkbox") {
          var group = el.closest(".cf-group[role=group]");
          if (group) {
            answers[el.name] = answers[el.name] || [];
            if (el.checked) answers[el.name].push(el.value);
          } else {
            answers[el.name] = el.checked;
            if (el.required && !el.checked) need(el, "“" + labelOf(el) + "”");
          }
          return;
        }
        answers[el.name] = el.value.trim();
        if (el.required && !answers[el.name]) need(el, "“" + labelOf(el) + "”");
        else if (answers[el.name] && !el.checkValidity()) need(el, "a valid answer for “" + labelOf(el) + "”");
      });
      Array.prototype.forEach.call(f.querySelectorAll(".cf-group[data-required]"), function (g) {
        var box = g.querySelector("input");
        if (!(answers[box.name] || []).length) need(box, "at least one choice for “" + labelOf(box) + "”");
      });
      if (problems.length) { show("err", "Please add " + problems.join(", ") + "."); if (first) first.focus(); return; }

      var btn = f.querySelector("button[type=submit]"), label = btn.textContent;
      btn.disabled = true; btn.textContent = "Sending…";
      var body = {form: f.getAttribute("data-form"), version: f.getAttribute("data-version"),
                  name: name, email: email, answers: answers, website: f.elements.website.value};
      // A Teacher Portal form goes to its own address, with the teacher's sign-in (site/teachers.js).
      var endpoint = f.getAttribute("data-endpoint"), headers = {"Content-Type": "application/json"};
      if (endpoint && window.tpAuth) headers.Authorization = "Bearer " + window.tpAuth();
      fetch(f.getAttribute("data-api") + (endpoint || "/form"), {method: "POST", headers: headers, body: JSON.stringify(body)})
        .then(function (r) { return r.json(); })
        .then(function (d) {
          if (!d.ok) throw new Error(d.error || "");
          f.reset(); f.hidden = true;
          show("ok", "<b>Thank you!</b> Your response was sent (number " + String(d.response_id).replace(/[<>&]/g, "") + "). " +
            (d.email_sent ? "A confirmation has been emailed to you." : "We couldn’t send the confirmation email, but your response was received."));
        })
        .catch(function (err) {
          btn.disabled = false; btn.textContent = label;
          var msg = String(err.message || "").replace(/[<>&]/g, "");
          show("err", "Sorry, your response didn’t go through" + (msg ? " (" + msg + ")" : "") +
            ". Please try again, or email <a href=\"mailto:iba@ibcministries.org\">iba@ibcministries.org</a>.");
        });
    });
  });
})();
