const cacheName = 'z97-office-shell-v3'
const root = new URL(self.registration.scope)
const relativePaths = [
  './', './index.html', './office.css', './office.js', './editors.js', './formats.js',
  './grid.js', './storage.js', './legacy.js', './calendar.js', './ooxml.js', './interchange.js',
]
const urls = new Set(relativePaths.map(path => new URL(path, root).href))

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(cacheName)
    await cache.addAll([...urls])
    await self.skipWaiting()
  })())
})

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys()
    await Promise.all(keys.filter(key => key.startsWith('z97-office-shell-') && key !== cacheName)
      .map(key => caches.delete(key)))
    await self.clients.claim()
  })())
})

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return
  const url = new URL(event.request.url)
  url.search = ''
  url.hash = ''
  if (!urls.has(url.href)) return
  event.respondWith((async () => {
    try {
      return await fetch(event.request)
    } catch {
      return await caches.match(url.href, { cacheName }) ?? Response.error()
    }
  })())
})
