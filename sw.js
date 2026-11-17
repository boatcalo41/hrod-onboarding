// HROD service worker — exists mainly so the app is installable
// (Add to Home Screen / "Install app") on Android/desktop Chrome, which
// requires a registered service worker with a fetch handler. Strategy is
// deliberately simple: stale-while-revalidate for same-origin GET requests,
// so the app still opens when offline or on a flaky connection, but always
// tries to fetch a fresher copy in the background for next time.
//
// Bump CACHE_NAME whenever app-shell files change so old caches get dropped.
const CACHE_NAME = "hrod-shell-v2";
const APP_SHELL = [
  "./",
  "./index.html",
  "./manifest.json",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  // Only handle same-origin GET — never touch the Google Apps Script API
  // calls (cross-origin) or anything that isn't a simple GET.
  if (req.method !== "GET" || new URL(req.url).origin !== self.location.origin) {
    return;
  }
  event.respondWith(
    caches.open(CACHE_NAME).then(async (cache) => {
      const cached = await cache.match(req);
      const network = fetch(req)
        .then((res) => {
          if (res && res.ok) cache.put(req, res.clone());
          return res;
        })
        .catch(() => null);
      return cached || (await network) || new Response("Offline", { status: 503 });
    })
  );
});
