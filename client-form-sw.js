// Minimal offline shell — caches just the admin console page itself so it
// opens instantly and the UI is viewable offline. Data (the Google Sheet)
// still needs a real network connection; this only covers the app shell,
// not live data, since caching stale prospect data would be actively
// misleading for an operations console.
const CACHE_NAME = 'dp-admin-shell-v1';
const SHELL_FILES = [
  './deeppocket-admin.html',
  './client-form-manifest.json'
];

self.addEventListener('install', function(event){
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache){
      return cache.addAll(SHELL_FILES);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(event){
  event.waitUntil(
    caches.keys().then(function(names){
      return Promise.all(names.filter(function(n){ return n!==CACHE_NAME; }).map(function(n){ return caches.delete(n); }));
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(event){
  var url = new URL(event.request.url);
  // Only ever serve the shell files from cache — every other request
  // (Apps Script API calls, Gemini, anything else) always goes to the
  // network untouched, since serving stale data here would be worse than
  // no offline support at all.
  var isShellFile = SHELL_FILES.some(function(f){ return url.pathname.endsWith(f.replace('./','')); });
  if(!isShellFile){ return; }

  event.respondWith(
    caches.match(event.request).then(function(cached){
      var fetchPromise = fetch(event.request).then(function(networkResponse){
        caches.open(CACHE_NAME).then(function(cache){ cache.put(event.request, networkResponse.clone()); });
        return networkResponse;
      }).catch(function(){ return cached; });
      return cached || fetchPromise;
    })
  );
});
