(function () {
  // Parse a cell into a sort key: W-L record, numeric, or text.
  function cellKey(td) {
    var t = (td.textContent || "").trim();
    var wl = t.match(/^(\d+)[\u2013-](\d+)$/);           // e.g. "3-0" -> rank by wins, fewer losses breaks ties
    if (wl) return parseInt(wl[1], 10) - parseInt(wl[2], 10) / 100;
    var money = t.match(/^\$(\d+(?:\.\d+)?)([KM])$/);       // "$60K", "$820K", "$26.3M" -> thousands
    if (money) return parseFloat(money[1]) * (money[2] === "M" ? 1000 : 1);
    var n = Number(t.replace(/[+]/g, "").replace(/[\u2013\u2212]/g, "-"));  // ".558", "17.0", "+49"
    if (t !== "" && !isNaN(n)) return n;
    return t.toLowerCase();
  }

  function makeSortable(table) {
    var ths = table.tHeaders && table.tHeaders.length ? null : null;
    var headRow = table.querySelector("thead tr");
    if (!headRow) return;
    var headers = Array.prototype.slice.call(headRow.children);
    var tbody = table.querySelector("tbody");

    headers.forEach(function (th, col) {
      th.classList.add("sortable");
      th.setAttribute("tabindex", "0");
      th.setAttribute("role", "columnheader");
      var caret = document.createElement("span");
      caret.className = "caret";
      th.appendChild(caret);

      function sortBy() {
        // Totals rows are locked to the bottom and never take part in a sort.
        // They belong in a tfoot; this catches any that sit in the tbody.
        var all = Array.prototype.slice.call(tbody.rows);
        var rows = all.filter(function (r) { return !r.classList.contains("tot"); });
        var tots = all.filter(function (r) { return r.classList.contains("tot"); });
        if (!rows.length) return;
        var isText = typeof cellKey(rows[0].cells[col]) === "string";
        var prev = th.getAttribute("aria-sort");
        // first click: text ascends A->Z, numbers descend high->low
        var dir = prev === "ascending" ? "descending"
                : prev === "descending" ? "ascending"
                : (isText ? "ascending" : "descending");

        headers.forEach(function (h) { h.removeAttribute("aria-sort"); });
        th.setAttribute("aria-sort", dir);
        var sign = dir === "ascending" ? 1 : -1;

        rows.sort(function (a, b) {
          var ka = cellKey(a.cells[col]), kb = cellKey(b.cells[col]);
          if (ka < kb) return -1 * sign;
          if (ka > kb) return 1 * sign;
          return 0;
        });
        rows.concat(tots).forEach(function (r) { tbody.appendChild(r); });
      }

      th.addEventListener("click", sortBy);
      th.addEventListener("keydown", function (e) {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); sortBy(); }
      });
    });
  }

  document.querySelectorAll(".tbl-wrap table").forEach(makeSortable);
})();

(function () {
  var tabs = Array.prototype.slice.call(document.querySelectorAll(".tab"));
  var panels = Array.prototype.slice.call(document.querySelectorAll(".panel"));
  function activate(id) {
    tabs.forEach(function (t) {
      var on = t.getAttribute("data-panel") === id;
      t.classList.toggle("is-active", on);
      t.setAttribute("aria-selected", on ? "true" : "false");
    });
    panels.forEach(function (p) { p.classList.toggle("is-active", p.id === id); });
    window.scrollTo(0, 0);
  }
  tabs.forEach(function (t) {
    t.addEventListener("click", function () { activate(t.getAttribute("data-panel")); });
  });
  document.addEventListener("click", function (e) {
    var g = e.target.closest ? e.target.closest("[data-goto]") : null;
    if (g) activate(g.getAttribute("data-goto"));
  });
  var home = document.querySelector(".mast-home");
  if (home) home.addEventListener("click", function () { activate("p-overview"); });
})();

(function () {
  /* GitHub Pages serves this page with cache-control: max-age=600 and gives us
     no way to change that, so a phone can sit on a stale copy. version.txt is
     fetched with no-store, which bypasses the HTTP cache entirely; when it names
     a newer build than this document, reload through a ?v= URL the browser has
     no cache entry for. The ?v= guard means we redirect at most once per build,
     so a mismatched version.txt can never loop. */
  var meta = document.querySelector('meta[name="build"]');
  var BUILD = meta ? meta.getAttribute("content") : "";
  function currentV() {
    var m = /[?&]v=([^&]*)/.exec(location.search);
    return m ? decodeURIComponent(m[1]) : null;
  }
  function check() {
    if (!window.fetch || location.protocol === "file:") return;
    fetch("version.txt", { cache: "no-store" }).then(function (r) {
      return r.ok ? r.text() : null;
    }).then(function (v) {
      if (!v) return;
      v = v.trim();
      if (!v || v === BUILD) {
        if (currentV()) history.replaceState(null, "", location.pathname);
        return;
      }
      if (currentV() === v) return;   /* already tried this build; do not loop */
      location.replace(location.pathname + "?v=" + encodeURIComponent(v));
    }).catch(function () {});
  }
  check();
  document.addEventListener("visibilitychange", function () {
    if (!document.hidden) check();
  });
  /* Offline shell. Registered after the version check so it never delays it. */
  if ("serviceWorker" in navigator &&
      (location.protocol === "https:" || location.hostname === "127.0.0.1" || location.hostname === "localhost")) {
    window.addEventListener("load", function () {
      navigator.serviceWorker.register("sw.js").catch(function () {});
    });
  }
  window.addEventListener("pageshow", function (e) { if (e.persisted) check(); });
})();
