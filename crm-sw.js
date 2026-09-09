// Offline app shell for the main CRM login/dashboard page only — this
// caches crm.html itself so it opens instantly and the UI is viewable
// offline. It deliberately does NOT cache anything from the backend
// (leads, clients, reminders) — showing stale client data while offline
// would be actively misleading for a working CRM, so all real data
// requires a live connection to the Apps Script backend.
//
// The 5 client-facing forms (upload/gatherinfo/review/findoc/feedback)
// are intentionally NOT included here — those are one-time-visit links a
// client opens once, not something worth installing or caching.
const CACHE_NAME = 'wealth-matrix-crm-shell-v1';
const SHELL_FILES = [
  './crm.html',
  './crm-manifest.json'
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
  var isShellFile = SHELL_FILES.some(function(f){ return url.pathname.endsWith(f.replace('./','')); });
  if(!isShellFile){ return; } // everything else (backend calls, AI calls, etc.) always goes to the network untouched

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
