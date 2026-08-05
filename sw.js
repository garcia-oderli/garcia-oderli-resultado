/* Service Worker — PPCP Capacidade & Eficiência
   Network-first (sem cache HTTP) para o shell do app, permite abrir offline
   com os últimos dados sincronizados. */

var CACHE_NAME = 'ppcp-dashboard-v8';
var URLS_TO_CACHE = [
  '/',
  '/manifest.json',
  '/icon-192.png',
  '/icon-512.png',
  '/logo-patrimar.jpg'
];

/* Sem skipWaiting automático: o worker novo fica em espera e só assume
   quando o usuário clica "Atualizar" no aviso. Trocar por baixo recarregava
   a página sozinha e derrubava filtro e rolagem no meio de uma reunião. */
self.addEventListener('message', function(event){
  if (event.data && event.data.tipo === 'ATIVAR_AGORA') self.skipWaiting();
});

self.addEventListener('install', function(event){
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache){
      return cache.addAll(URLS_TO_CACHE);
    })
  );
});

self.addEventListener('activate', function(event){
  event.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(
        keys.filter(function(k){ return k !== CACHE_NAME; })
            .map(function(k){ return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(event){
  /* Network-first pro HTML principal (sempre busca versão nova, ignorando
     cache HTTP do navegador — evita tela desatualizada após deploy),
     cache-first pro resto (ícones, manifest) */
  var url = new URL(event.request.url);

  if (event.request.mode === 'navigate' || url.pathname === '/') {
    event.respondWith(
      fetch(event.request, { cache: 'no-store' })
        .then(function(resp){
          var clone = resp.clone();
          caches.open(CACHE_NAME).then(function(cache){ cache.put(event.request, clone); });
          return resp;
        })
        .catch(function(){ return caches.match(event.request).then(function(r){ return r || caches.match('/'); }); })
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then(function(cached){
      return cached || fetch(event.request);
    })
  );
});
