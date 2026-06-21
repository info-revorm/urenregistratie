/* ============================================================
   revorm urenregistratie — service worker (PWA, offline)
   - Cachet alleen de app-shell (HTML/JS/icons).
   - KRITISCH: verzoeken naar Microsoft-login/Graph en de
     auth-redirect (?code=/state=) NOOIT cachen -> netwerk-only,
     anders breekt de login stilletjes.
   ============================================================ */

const CACHE = "revorm-uren-v4";
const SHELL = [
  "./",
  "./index.html",
  "./msal-browser.min.js",
  "./msal-config.js",
  "./sync.js",
  "./manifest.webmanifest",
  "./icons/icon-192.png",
  "./icons/icon-512.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

function isMicrosoft(host){
  return /login\.microsoftonline\.com|graph\.microsoft\.com|msauth\.net|msftauth\.net|login\.live\.com|login\.windows\.net/.test(host);
}

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if(req.method !== "GET") return;

  let url;
  try{ url = new URL(req.url); }catch(e){ return; }

  // Microsoft-login en Graph: altijd rechtstreeks naar het netwerk, nooit cachen.
  if(isMicrosoft(url.hostname)) return;

  // Auth-redirect terug naar onze app (?code=/state=/error=): netwerk-only.
  if(url.search.indexOf("code=") > -1 || url.search.indexOf("state=") > -1 || url.search.indexOf("error=") > -1){
    event.respondWith(fetch(req));
    return;
  }

  // Andere oorsprong: laat de browser het standaard afhandelen.
  if(url.origin !== self.location.origin) return;

  // Paginanavigatie: netwerk eerst, val terug op de gecachete app-shell (offline).
  if(req.mode === "navigate"){
    event.respondWith(
      fetch(req).catch(() => caches.match("./index.html").then((r) => r || caches.match("./")))
    );
    return;
  }

  // Statische assets: stale-while-revalidate (direct uit cache, op de achtergrond verversen).
  event.respondWith(
    caches.open(CACHE).then((cache) =>
      cache.match(req).then((cached) => {
        const network = fetch(req).then((resp) => {
          if(resp && resp.status === 200 && resp.type === "basic") cache.put(req, resp.clone());
          return resp;
        }).catch(() => cached);
        return cached || network;
      })
    )
  );
});
