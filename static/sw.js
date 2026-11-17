// Service worker: rende l'app installabile e veloce. I dati (/api) non vengono mai messi in cache.
const V = "soh-v31", SHELL = ["/", "/static/icon-192.png", "/static/logo.png", "/manifest.webmanifest"];
self.addEventListener("install", e => { e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== V).map(x => caches.delete(x)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== "GET" || u.origin !== location.origin || u.pathname.startsWith("/api/") || u.pathname.startsWith("/cal/")) return;
  if (r.mode === "navigate") {   // pagina: rete prima, copia locale se offline
    e.respondWith(fetch(r).then(x => { const c = x.clone(); caches.open(V).then(ca => ca.put("/", c)); return x; }).catch(() => caches.match("/")));
    return;
  }
  e.respondWith(caches.match(r).then(m => m || fetch(r).then(x => { if (x.ok) { const c = x.clone(); caches.open(V).then(ca => ca.put(r, c)); } return x; })));
});
