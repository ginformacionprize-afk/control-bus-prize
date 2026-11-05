// Control Bus Prize - service worker
// La página y los datos se piden SIEMPRE primero a internet (así cada publicación nueva llega sola).
// Si no hay señal, se usa la última copia guardada para que la app siga funcionando.
const CACHE = "controlbus-v13";
const CORE = ["./", "./index.html", "./manifest.webmanifest", "./icons/icon-192.png", "./icons/icon-512.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

async function networkFirst(req){
  const c = await caches.open(CACHE);
  try{
    const res = await fetch(req, {cache: "no-store"});
    if(res && res.ok) c.put(req, res.clone());
    return res;
  }catch(e){
    return (await c.match(req, {ignoreSearch: true})) || (await c.match("./index.html")) || Response.error();
  }
}
async function cacheFirst(req){
  const c = await caches.open(CACHE);
  const hit = await c.match(req);
  if(hit) return hit;
  const res = await fetch(req);
  if(res && (res.ok || res.type === "opaque")) c.put(req, res.clone());
  return res;
}

self.addEventListener("fetch", e => {
  const req = e.request;
  if(req.method !== "GET") return;
  const url = new URL(req.url);
  const same = url.origin === self.location.origin;
  if(same && url.pathname.startsWith("/api/")) return; // el servidor siempre en vivo, nunca desde caché
  if(same && (req.mode === "navigate" || url.pathname.endsWith(".html") ||url.pathname.endsWith(".webmanifest"))){
    e.respondWith(networkFirst(req)); return;
  }
  // Librerías y fuentes externas (versión fija): se guardan para funcionar sin señal
  if(/cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|fonts\.(googleapis|gstatic)\.com/.test(url.hostname) || same){
    e.respondWith(cacheFirst(req));
  }
});
