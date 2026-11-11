// Guarda la app en el móvil para que abra rápido. Los datos siempre se piden a la hoja.
const CACHE = "vacaciones-v6";
const ARCHIVOS = ["./", "index.html", "manifest.webmanifest", "icono-192.png", "icono-512.png", "icono-180.png", "logo.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARCHIVOS))); self.skipWaiting(); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))); self.clients.claim(); });
self.addEventListener("fetch", e => {
  const r = e.request;
  if (r.method !== "GET" || new URL(r.url).origin !== location.origin) return;   // la hoja y las fuentes van directas
  // primero la red (para coger cambios de la app); si no hay conexión, la copia guardada
  e.respondWith(fetch(r, {cache: "no-cache"}).then(resp => { const c = resp.clone(); caches.open(CACHE).then(k => k.put(r, c)); return resp; }).catch(() => caches.match(r)));
});
