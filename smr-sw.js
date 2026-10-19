// Keep CACHE_NAME and APP_VERSION in smr-app.html the same.
const CACHE_NAME = 'smr-v26';
const APP = 'smr-app.html';
const LOCAL_ASSETS = [APP, 'smr-manifest.json'];
const CDN_ASSETS = ['https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'];
// How long to wait for the network before opening the saved copy of the app.
const NETWORK_WAIT_MS = 3000;

self.addEventListener('install', e => {
  // A CDN hiccup must not stop the update from installing (roundtable entry 060).
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache =>
      cache.addAll(LOCAL_ASSETS).then(() =>
        Promise.all(CDN_ASSETS.map(url => cache.add(url).catch(() => {})))
      )
    )
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// The app page comes from the network when there is one, so a new version
// shows on the next open. With no signal, or a slow one, the saved copy opens.
function appPage() {
  const network = fetch(APP, {cache: 'no-cache'}).then(res => {
    if (res.ok) {
      const copy = res.clone();
      caches.open(CACHE_NAME).then(cache => cache.put(APP, copy));
    }
    return res;
  });
  const saved = () => caches.match(APP);
  const slow = new Promise(resolve => setTimeout(resolve, NETWORK_WAIT_MS))
    .then(saved)
    .then(res => res || network);
  return Promise.race([network.catch(saved), slow]);
}

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (url.origin === self.location.origin && url.pathname.endsWith('/' + APP)) {
    e.respondWith(appPage());
    return;
  }
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request))
  );
});
