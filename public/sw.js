const CACHE_NAME = 'cerne-cache-v2';

self.addEventListener('install', () => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const { request } = event;
  if (request.method !== 'GET') return;
  // Nunca guarda em cache chamadas ao Dropbox — precisam sempre ser atuais.
  if (request.url.includes('dropboxapi.com') || request.url.includes('dropbox.com')) return;

  // Navegação (o index.html): busca na rede primeiro. Se cachearmos o HTML, depois de um novo
  // deploy ele pode apontar para um bundle .js com hash antigo que não existe mais — tela branca
  // até a segunda visita. Cache só entra como fallback se a rede falhar (modo offline).
  if (request.mode === 'navigate') {
    event.respondWith(
      fetch(request)
        .then((response) => {
          caches.open(CACHE_NAME).then((cache) => cache.put(request, response.clone()));
          return response;
        })
        .catch(() => caches.match(request))
    );
    return;
  }

  // Demais arquivos (JS/CSS com hash, imagens etc.): stale-while-revalidate normalmente.
  event.respondWith(
    caches.open(CACHE_NAME).then(async (cache) => {
      const cached = await cache.match(request);
      const networkFetch = fetch(request)
        .then((response) => {
          if (response && response.status === 200) cache.put(request, response.clone());
          return response;
        })
        .catch(() => cached);
      return cached || networkFetch;
    })
  );
});
