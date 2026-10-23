// The Kings Corner (parents page): each tile opens its section in a pop-up instead of scrolling down the page.
// Each pop-up has a back button at the top and bottom; the browser's back button and Esc close it too.
// Old links like parents/#hot-lunch (and the staff link parents/?staff#hot-lunch) open that section's pop-up.
// Without JavaScript (or without <dialog> support) the sections stay listed below the tiles, as before.
(function () {
  var kc = document.querySelector(".kc");
  if (!kc || typeof HTMLDialogElement !== "function") return;
  var label = kc.getAttribute("data-back") || "Back";
  var arrow = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" ' +
    'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M20 12H4"/><path d="M10 6l-6 6 6 6"/></svg>';
  var pops = {}, open = null;

  function backButton(where) {
    var b = document.createElement("button");
    b.type = "button";
    b.className = "kc-back kc-back-" + where;
    b.innerHTML = arrow + "<span></span>";
    b.lastChild.textContent = label;
    b.addEventListener("click", function () { close(true); });
    return b;
  }

  Array.prototype.forEach.call(kc.querySelectorAll(".kc-sections .block[id]"), function (block) {
    var d = document.createElement("dialog");
    var h = block.querySelector("h2");
    d.className = "kc-pop";
    if (h) { h.id = h.id || block.id + "-title"; d.setAttribute("aria-labelledby", h.id); }
    var bar = document.createElement("div");
    bar.className = "kc-pop-bar";
    bar.appendChild(backButton("top"));
    d.appendChild(bar);
    d.appendChild(block); // moves the section (and any form in it, with its listeners) into the pop-up
    var foot = document.createElement("div");
    foot.className = "kc-pop-foot";
    foot.appendChild(backButton("bottom"));
    d.appendChild(foot);
    d.addEventListener("cancel", function (e) { e.preventDefault(); close(true); }); // Esc
    d.addEventListener("click", function (e) { if (e.target === d) close(true); }); // click outside the box
    kc.appendChild(d);
    pops[block.id] = d;
  });
  kc.classList.add("kc-js"); // hides the leftover section list

  function show(id, push) {
    var d = pops[id];
    if (!d) return false;
    if (open && open !== d) open.close();
    if (!d.open) d.showModal();
    d.scrollTop = 0;
    setTimeout(function () { d.scrollTop = 0; }, 0); // after the browser's own jump to the #section
    open = d;
    document.documentElement.classList.add("kc-lock");
    if (push) history.pushState({ kc: id }, "", location.pathname + location.search + "#" + id);
    return true;
  }

  function close(viaButton) {
    if (!open) return;
    var d = open;
    open = null;
    d.close();
    document.documentElement.classList.remove("kc-lock");
    if (viaButton) {
      if (history.state && history.state.kc) history.back(); // undo the entry the tile added
      else history.replaceState(null, "", location.pathname + location.search); // opened from a link: drop the #
    }
    var tile = kc.querySelector('.kc-tile[href="#' + d.querySelector(".block").id + '"]');
    if (tile) tile.focus();
  }

  kc.addEventListener("click", function (e) {
    var a = e.target.closest && e.target.closest('a.kc-tile[href^="#"], a.kc-btn[href^="#"], a.kc-flyer[href^="#"]');
    if (a && show(a.getAttribute("href").slice(1), true)) e.preventDefault();
  });

  window.addEventListener("popstate", function () {
    var id = location.hash.slice(1);
    if (pops[id]) show(id, false); else close(false);
  });

  if (location.hash && pops[location.hash.slice(1)]) show(location.hash.slice(1), false);
})();
