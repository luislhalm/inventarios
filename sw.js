// Service worker de la PWA "Inventarios CAPA".
//
// IMPORTANTE al publicar una versión nueva de index.html: sube también este archivo
// cambiando el número de CACHE_VERSION de abajo. Si no lo cambias, algunos dispositivos
// podrían seguir viendo la versión anterior cacheada durante un tiempo.
const CACHE_VERSION = "v6";
const CACHE_NAME = `capa-inventarios-${CACHE_VERSION}`;

// Lo mínimo para que la app arranque (la "carcasa") aunque no haya conexión.
// Los datos reales (artículos, registros...) siempre se piden en directo a SharePoint;
// este service worker no los cachea, de eso ya se encarga la propia app (cola de
// pendientes en IndexedDB) cuando falla la red.
const ARCHIVOS_CARCASA = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./icon-192.png",
  "./icon-512.png",
];

// CDNs de las librerías (MSAL, XLSX): se cachean también para poder abrir la app offline,
// pero como son de otro origen, la petición se hace en modo "no-cors" (respuesta "opaca":
// no se puede leer el contenido, pero sí se puede guardar y volver a servir).
const ARCHIVOS_CDN = [
  "https://cdn.jsdelivr.net/npm/@azure/msal-browser@3/lib/msal-browser.min.js",
  "https://unpkg.com/@azure/msal-browser@3/lib/msal-browser.min.js",
  "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
  "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
];

self.addEventListener("install", (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    await cache.addAll(ARCHIVOS_CARCASA).catch(()=>{});
    await Promise.all(
      ARCHIVOS_CDN.map(url => fetch(url, { mode: "no-cors" }).then(r => cache.put(url, r)).catch(()=>{}))
    );
    self.skipWaiting();
  })());
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const nombres = await caches.keys();
    await Promise.all(nombres.filter(n => n !== CACHE_NAME).map(n => caches.delete(n)));
    self.clients.claim();
  })());
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  const esMismoOrigen = url.origin === self.location.origin;
  const esCDNConocido = ARCHIVOS_CDN.some(cdn => event.request.url.startsWith(cdn.split("?")[0].replace(/\/[^/]*$/, "")) || event.request.url === cdn);

  // Nunca interceptamos llamadas a Microsoft Graph, SharePoint, login/autenticación, ni
  // descargas de contenido (downloadUrl) — esas siempre deben ir en directo a la red, la
  // propia app ya sabe qué hacer si fallan por falta de conexión.
  if (!esMismoOrigen && !esCDNConocido) return;

  if (event.request.mode === "navigate"){
    // Página principal: primero intentamos traer la versión más reciente de la red
    // (para que veáis los cambios en cuanto los publico), y si falla, servimos la copia
    // guardada para poder al menos abrir la app sin conexión.
    event.respondWith((async () => {
      try{
        const respuesta = await fetch(event.request);
        const cache = await caches.open(CACHE_NAME);
        cache.put("./index.html", respuesta.clone());
        return respuesta;
      }catch(e){
        const cache = await caches.open(CACHE_NAME);
        return (await cache.match("./index.html")) || (await cache.match("./"));
      }
    })());
    return;
  }

  // Resto de archivos propios y de los CDN: primero caché (rápido), red de refresco/reserva.
  event.respondWith((async () => {
    const cache = await caches.open(CACHE_NAME);
    const enCache = await cache.match(event.request);
    if (enCache) return enCache;
    try{
      const respuesta = await fetch(event.request);
      cache.put(event.request, respuesta.clone());
      return respuesta;
    }catch(e){
      return enCache; // undefined si tampoco había nada guardado
    }
  })());
});
