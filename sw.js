/**
 * sw.js — POS File Checker Service Worker
 *
 * Strategy: Cache-first for all local app assets (shell, CSS, JS).
 * Network-first for Google Fonts (nice-to-have, not required).
 * CDN export libraries (jsPDF, SheetJS) are NOT cached — they are
 * loaded only when the user triggers an export, and graceful error
 * messages are shown if they are unavailable offline.
 *
 * The app's core functionality (upload, parse, validate) works
 * fully offline because all business logic is in script.js which
 * is in the precache list.
 */

const CACHE_NAME = 'pos-checker-v1';

/**
 * App shell: all files required for the app to load and run.
 * Paths are relative to the service worker's scope (the site root).
 * Update CACHE_NAME above when any of these files change, so
 * existing users get the new version on next visit.
 */
const PRECACHE_URLS = [
  './',
  './index.html',
  './style.css',
  './script.js',
  './manifest.json',
  './icons/favicon.svg',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png',
];

/* ── Install: precache the app shell ─────────────────────────────────── */
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(PRECACHE_URLS))
      .then(() => self.skipWaiting())   // Activate immediately
  );
});

/* ── Activate: delete old caches ─────────────────────────────────────── */
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys
          .filter(key => key !== CACHE_NAME)
          .map(key => caches.delete(key))
      ))
      .then(() => self.clients.claim())  // Take control of open pages
  );
});

/* ── Fetch: serve from cache, fall back to network ───────────────────── */
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Only handle GET requests
  if (event.request.method !== 'GET') return;

  // Let CDN requests (jsPDF, SheetJS) go straight to network.
  // These are large libraries only needed for export — don't cache them.
  if (url.hostname === 'cdnjs.cloudflare.com') return;

  // Google Fonts: network-first so font updates are picked up,
  // fall back to cache if offline.
  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
          return response;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  // App shell: cache-first
  event.respondWith(
    caches.match(event.request)
      .then(cached => {
        if (cached) return cached;
        // Not in cache — fetch, cache, and return
        return fetch(event.request).then(response => {
          if (!response || response.status !== 200 || response.type === 'opaque') {
            return response;
          }
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
          return response;
        });
      })
      .catch(() => {
        // If both cache and network fail for a navigation request,
        // serve index.html so the SPA shell still loads.
        if (event.request.mode === 'navigate') {
          return caches.match('./index.html');
        }
      })
  );
});
