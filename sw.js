// Service worker: app funciona sem sinal (arquivos do app + mapa já visto ou baixado).
const VERSAO = 'v3-2026-11-03';
const APP = 'app-' + VERSAO;
const TILES = 'tiles-v1';
const ARQUIVOS = [
  './', './index.html', './config.js', './data.json', './manifest.webmanifest',
  './icon-192.png', './icon-512.png',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js',
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(APP).then(c => c.addAll(ARQUIVOS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(
    ks.filter(k => k.startsWith('app-') && k !== APP).map(k => caches.delete(k))
  )).then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = req.url;
  // Backend (planilha): sempre rede, nunca cache.
  if (url.includes('script.google.com') || url.includes('googleusercontent.com')) return;
  // Mapa: cache primeiro; o que vier da rede é guardado para uso offline.
  if (url.includes('arcgisonline.com')) {
    e.respondWith(caches.open(TILES).then(async c => {
      const hit = await c.match(url);
      if (hit) return hit;
      try {
        const r = await fetch(req);
        if (r.ok || r.type === 'opaque') c.put(url, r.clone());
        return r;
      } catch (err) {
        return new Response('', { status: 504 });
      }
    }));
    return;
  }
  // Arquivos do app: rede primeiro (pega atualização), cache se estiver sem sinal.
  e.respondWith(fetch(req).then(r => {
    if (r.ok && (url.startsWith(self.location.origin) || url.includes('unpkg.com'))) {
      const cp = r.clone();
      caches.open(APP).then(c => c.put(req, cp));
    }
    return r;
  }).catch(() => caches.match(req, { ignoreSearch: true }).then(h => h || caches.match('./index.html'))));
});
