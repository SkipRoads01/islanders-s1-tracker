/* Islanders S1 Tracker - offline shell.
   Network-first for the page so an online phone always gets the current build
   and the version.txt check in index.html keeps working untouched; the cache is
   only the fallback when there is no connection. version.txt itself is never
   intercepted: offline it must fail so the page's own catch ignores it. */
var CACHE = "islanders-s1-v1";
var PAGE  = "./";

self.addEventListener("install", function (e) {
  e.waitUntil(
    caches.open(CACHE).then(function (c) { return c.add(PAGE); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener("activate", function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (k) {
        return k === CACHE ? null : caches.delete(k);
      }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener("fetch", function (e) {
  var req = e.request;
  if (req.method !== "GET") return;
  var url;
  try { url = new URL(req.url); } catch (err) { return; }
  if (url.origin !== self.location.origin) return;
  if (url.pathname.indexOf("version.txt") !== -1) return;   /* always live */

  if (req.mode === "navigate") {
    /* Store under the bare path so a ?v=<build> URL still resolves offline. */
    e.respondWith(
      fetch(req).then(function (res) {
        var copy = res.clone();
        caches.open(CACHE).then(function (c) { c.put(PAGE, copy); });
        return res;
      }).catch(function () {
        return caches.match(PAGE).then(function (r) { return r || Response.error(); });
      })
    );
    return;
  }

  e.respondWith(
    caches.match(req).then(function (hit) {
      return hit || fetch(req).then(function (res) {
        var copy = res.clone();
        caches.open(CACHE).then(function (c) { c.put(req, copy); });
        return res;
      });
    })
  );
});
