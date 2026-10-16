// Hot lunch online ordering (see lunch.py and worker/). Prices, days and the deadline come
// from lunch/menu.json; the order service checks them again, so nothing here is trusted.
(function () {
  var form = document.getElementById("lunch-order");
  if (!form || !window.fetch) return;
  var list = document.getElementById("lunch-children");
  var status = document.getElementById("lunch-status");
  var totalEl = document.getElementById("lunch-total");
  var addBtn = document.getElementById("lunch-add");
  // Staff order from ?staff (not linked on the site): add a "Staff" choice to the grade list.
  if (/[?&]staff\b/.test(location.search)) {
    var opt = document.createElement("option"); opt.value = "STAFF"; opt.textContent = "Staff";
    list.querySelector("[name=grade]").appendChild(opt);
  }
  var template = list.firstElementChild.cloneNode(true);
  var MAX = 8, menu = null;

  function money(n) { return "$" + n.toFixed(2); }
  function show(cls, html) { status.className = "notice " + cls; status.innerHTML = html; status.scrollIntoView({block: "center"}); }
  function children() { return Array.prototype.slice.call(list.querySelectorAll(".lunch-child")); }
  function price(grade) {
    if (!menu || !grade) return 0;
    for (var i = 0; i < menu.grades.length; i++)
      if (menu.grades[i].value === grade)
        return menu.grades[i].staff ? menu.prices.staff : menu.grades[i].preschool ? menu.prices.preschool : menu.prices.regular;
    return 0;
  }
  function refresh() {
    var kids = children(), total = 0;
    kids.forEach(function (c, i) {
      c.querySelector("legend").textContent = kids.length > 1 ? "Student " + (i + 1) : "Student";
      c.querySelector(".lunch-remove").hidden = kids.length < 2;
      total += price(c.querySelector("[name=grade]").value) * c.querySelectorAll("[name=days]:checked").length;
    });
    totalEl.textContent = money(total);
    addBtn.hidden = kids.length >= MAX;
  }
  function close() {
    document.getElementById("lunch-closed").hidden = false;
    Array.prototype.forEach.call(form.elements, function (el) { el.disabled = true; });
  }

  addBtn.addEventListener("click", function () {
    var c = template.cloneNode(true);
    list.appendChild(c); refresh(); c.querySelector("[name=student]").focus();
  });
  list.addEventListener("click", function (e) {
    if (e.target.classList.contains("lunch-remove")) { e.target.closest(".lunch-child").remove(); refresh(); }
  });
  form.addEventListener("change", refresh);
  form.addEventListener("input", function () {
    if (status.classList.contains("err")) { status.className = ""; status.innerHTML = ""; }
  });

  fetch(form.getAttribute("data-menu"), {cache: "no-store"}).then(function (r) { return r.json(); }).then(function (m) {
    menu = m; refresh();
    if (!m.deadline || Date.now() > Date.parse(m.deadline)) close();
  }).catch(function () { /* the order service checks everything anyway */ });

  form.addEventListener("submit", function (e) {
    e.preventDefault();
    var f = form.elements, problems = [];
    var order = {
      parent_first: f.parent_first.value.trim(), parent_last: f.parent_last.value.trim(),
      email: f.email.value.trim(), phone: f.phone.value.trim(), website: f.website.value,
      week_start: menu ? menu.week_start : "", children: []
    };
    if (!order.parent_first || !order.parent_last) problems.push("your first and last name");
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(order.email)) problems.push("a valid email address");
    if (!order.phone) problems.push("a phone number");
    children().forEach(function (c, i) {
      var days = Array.prototype.map.call(c.querySelectorAll("[name=days]:checked"), function (x) { return x.value; });
      var kid = {name: c.querySelector("[name=student]").value.trim(), grade: c.querySelector("[name=grade]").value, days: days};
      var who = children().length > 1 ? "student " + (i + 1) : "the student";
      if (!kid.name) problems.push("the name of " + who);
      if (!kid.grade) problems.push("the grade of " + who);
      if (!days.length) problems.push("at least one day for " + (kid.name || who));
      order.children.push(kid);
    });
    if (problems.length) { show("err", "Please add " + problems.join(", ") + "."); return; }

    var btn = form.querySelector("button[type=submit]");
    btn.disabled = true; btn.textContent = "Placing order…";
    fetch(form.getAttribute("data-api") + "/order", {
      method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(order)
    }).then(function (r) { return r.json().then(function (d) { return {ok: r.ok, data: d}; }); })
      .then(function (r) {
        if (!r.ok || !r.data.ok) throw new Error(r.data.error || "");
        form.hidden = true;
        var p = document.createElement("div");
        p.className = "notice ok"; p.setAttribute("role", "status");
        p.innerHTML = "<b>Thank you! Your order is placed.</b> Order number " + r.data.order_id +
          ", total " + money(r.data.total) + ". " + (r.data.email_sent
            ? "A confirmation has been emailed to you."
            : "We couldn’t send the confirmation email, but your order was received.") +
          " To change it, please call the office at <a href=\"tel:3018561616\">(301) 856-1616</a>.";
        form.parentNode.insertBefore(p, form.nextSibling); p.scrollIntoView({block: "center"});
      })
      .catch(function (err) {
        btn.disabled = false; btn.textContent = "Place order";
        show("err", "Sorry, your order didn’t go through" + (err.message ? " (" + err.message.replace(/</g, "&lt;") + ")" : "") +
          ". Please try again, or call the office at <a href=\"tel:3018561616\">(301) 856-1616</a>.");
      });
  });
  refresh();
})();
