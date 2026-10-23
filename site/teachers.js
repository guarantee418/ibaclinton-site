// The Teacher Portal (ibaclinton.com/teachers/): the page only holds the sign-in box. With the right password the
// order service (worker/src/teachers.js) gives a sign-in token, and with it the portal itself (built by
// teachers.py from the private repo). It then works like The Kings Corner: kings.js opens each tile's pop-up,
// kings-slides.js turns the announcements and forms.js sends the forms (with the token).
(function () {
  var box = document.getElementById("tp"), holder = document.getElementById("tp-portal");
  if (!box || !window.fetch) return;
  var api = box.getAttribute("data-api"), root = box.getAttribute("data-root");
  var form = document.getElementById("tp-form"), status = document.getElementById("tp-status");
  var KEY = "ibaTeacherPortal";

  function stored() {
    try { return localStorage.getItem(KEY) || sessionStorage.getItem(KEY) || ""; } catch (e) { return ""; }
  }
  function keep(token, remember) {
    try {
      localStorage.removeItem(KEY); sessionStorage.removeItem(KEY);
      if (token) (remember ? localStorage : sessionStorage).setItem(KEY, token);
    } catch (e) { /* private browsing: signed in for this page only */ }
    window.tpToken = token;
  }
  window.tpAuth = function () { return window.tpToken || stored(); }; // forms.js sends it with Teacher Portal forms

  function say(text, bad) { status.textContent = text; status.className = "tp-status" + (bad ? " tp-bad" : ""); }
  function showLogin(text) {
    form.hidden = false;
    if (text) say(text, true);
  }
  function addScript(name) {
    var s = document.createElement("script");
    s.src = root + name;
    document.body.appendChild(s);
  }

  function show(d) {
    // Uploaded files go through the service with a short-lived pass.
    var files = api + "/teachers/file/" + encodeURIComponent(d.files) + "/";
    holder.innerHTML = '<div class="tp-bar"><div class="kc-in"><span>Signed in to The King’s Court</span>' +
      '<button type="button" class="tp-out">Sign out</button></div></div>' + String(d.html).split("__TP_FILES__/").join(files);
    holder.querySelector(".tp-out").addEventListener("click", function () { keep(""); location.replace(location.pathname); });
    box.parentNode.removeChild(box); // kings.js works on the first .kc on the page: the portal's
    holder.hidden = false;
    addScript("kings.js");
    if (holder.querySelector(".kc-slider")) addScript("kings-slides.js");
    if (holder.querySelector("form.cf-form")) addScript("forms.js");
  }

  function load() {
    var token = window.tpAuth();
    if (!token) return showLogin();
    say("Opening the portal…");
    fetch(api + "/teachers/portal", {headers: {Authorization: "Bearer " + token}})
      .then(function (r) { return r.json().then(function (d) { d.status = r.status; return d; }); })
      .then(function (d) {
        if (d.ok) return show(d);
        if (d.signin) { keep(""); return showLogin("Please sign in again."); }
        throw new Error(d.error || "");
      })
      .catch(function (err) {
        showLogin("Sorry, the portal didn’t open" + (err.message ? " (" + err.message + ")" : "") + ". Please try again.");
      });
  }

  form.addEventListener("submit", function (e) {
    e.preventDefault();
    var pass = form.elements.password.value, remember = form.elements.remember.checked;
    if (!pass.trim()) { say("Please type the password.", true); form.elements.password.focus(); return; }
    var btn = form.querySelector("button");
    btn.disabled = true;
    say("Signing in…");
    fetch(api + "/teachers/login", {method: "POST", headers: {"Content-Type": "application/json"},
                                    body: JSON.stringify({password: pass, remember: remember})})
      .then(function (r) { return r.json(); })
      .then(function (d) {
        btn.disabled = false;
        if (!d.ok) { say(d.error || "That didn’t work; please try again.", true); form.elements.password.select(); return; }
        keep(d.token, remember);
        load();
      })
      .catch(function () { btn.disabled = false; say("Couldn’t reach the portal; please check your connection and try again.", true); });
  });

  load();
})();
