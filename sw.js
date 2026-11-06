/* Offline support: the app shell, fonts and the Excel module are cached after the first visit. */
const VERSION = "cans-v5";
const SHELL = [
  "./", "./index.html", "./style.css", "./facts.js", "./cloud.js", "./app.js", "./manifest.webmanifest",
  "./icon-192.png", "./icon-512.png", "./apple-touch-icon.png",
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Serve from cache right away, refresh the cache in the background.
self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET" || !req.url.startsWith("http")) return;
  // the cloud (sign-in, data, photos) always goes to the network
  if (new URL(req.url).hostname.endsWith(".supabase.co")) return;
  e.respondWith(
    caches.open(VERSION).then(async (cache) => {
      const cached = await cache.match(req, { ignoreSearch: req.mode === "navigate" });
      const network = fetch(req)
        .then((res) => { if (res && (res.ok || res.type === "opaque")) cache.put(req, res.clone()); return res; })
        .catch(() => cached || (req.mode === "navigate" ? cache.match("./index.html") : Response.error()));
      if (cached) { e.waitUntil(network.catch(() => {})); return cached; }
      return network;
    })
  );
});
