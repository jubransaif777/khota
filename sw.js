/* خدمة العمل بلا إنترنت: تعرض آخر نسخة محفوظة من الصفحات العامة وبياناتها حين ينقطع الاتصال.
   • الصفحات والبيانات: «الشبكة أولاً» (تُجلَب الأحدث دائماً)، والنسخة المحفوظة عند الانقطاع أو البطء فقط.
   • لا تُخزَّن أبداً: ألف، التقرير، الإدارة، الموظفون، التقييمات، الإحصاءات، أو أي طلب غير GET أو فيه معاملات (مثل ?audit=).
   لتعطيلها فوراً عند أي مشكلة: افتح أي صفحة بإضافة  ?nosw=1  إلى الرابط. */
var VERSION = 'v1', SHELL = 'ssuae-shell-' + VERSION, DATA = 'ssuae-data-' + VERSION;
var PAGES = ['/', '/summary', '/schedule', '/exams', '/news', '/links'];
var DATA_PATHS = ['/.netlify/functions/plan', '/.netlify/functions/exams', '/.netlify/functions/schedule', '/.netlify/functions/content'];
var ASSETS = ['/site-class.js', '/site-content.js', '/site-rating.js', '/site-track.js', '/site-merge.js', '/site-pwa.js', '/manifest.webmanifest', '/icon-192.png', '/icon-512.png', '/apple-touch-icon.png'];
var NEVER = /^\/(admin|staff|alef|report|certs|docs|files)(\/|\.html$|$)|^\/\.netlify\/functions\/(alef|report|certs|docs|staff|admin|rating|track|health|health-cron)(\/|$)/;
var PAGE_TIMEOUT = 4000, DATA_TIMEOUT = 5000;

function pageKey(p) { p = p.replace(/\/index\.html$/, '/').replace(/\.html$/, ''); return p.length > 1 ? p.replace(/\/+$/, '') : p; }
// يقرّر المعالجة: bypass | page | data | asset
function route(pathname, search, method, mode, sameOrigin) {
  if (!sameOrigin || method !== 'GET' || NEVER.test(pathname)) return 'bypass';
  if (DATA_PATHS.indexOf(pathname) >= 0) return search ? 'bypass' : 'data';          // أي معاملات (?audit=…) ← لا تخزين
  if (mode === 'navigate' && PAGES.indexOf(pageKey(pathname)) >= 0) return 'page';
  if (ASSETS.indexOf(pathname) >= 0) return 'asset';
  return 'bypass';
}
function offlineHtml() {
  return new Response('<!doctype html><html lang="ar" dir="rtl"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>غير متصل</title><body style="font-family:system-ui,sans-serif;text-align:center;padding:48px 20px;background:#F4F7F6;color:#15302D"><div style="font-size:48px">📡</div><h2>أنت غير متصل بالإنترنت</h2><p style="line-height:1.9">لم تُحفظ هذه الصفحة على جهازك بعد.<br>افتحها مرة واحدة وأنت متصل، وستظهر لك بعدها حتى بلا إنترنت.</p><button onclick="location.reload()" style="padding:10px 24px;border:0;border-radius:12px;background:#0E5A54;color:#fff;font-size:16px">إعادة المحاولة</button></body></html>', { status: 503, headers: { 'Content-Type': 'text/html; charset=utf-8' } });
}
function fromCache(res) { var h = new Headers(res.headers); h.set('x-from-cache', '1'); return new Response(res.body, { status: res.status, statusText: res.statusText, headers: h }); }
// الشبكة أولاً، وعند الفشل/البطء النسخة المحفوظة؛ والشبكة تُكمل وتحدّث النسخة في الخلفية
function networkFirst(req, key, cacheName, timeout, fallback, event) {
  return caches.open(cacheName).then(function (cache) {
    var net = fetch(req).then(function (res) { if (res && res.ok) { try { cache.put(key, res.clone()); } catch (e) {} } return res; });
    var guarded = net.catch(function () { return null; });
    var timer = new Promise(function (resolve) { setTimeout(function () { resolve('timeout'); }, timeout); });
    return Promise.race([guarded, timer]).then(function (first) {
      if (first && first !== 'timeout') return first;                                   // ردّت الشبكة في الوقت
      return cache.match(key).then(function (hit) {
        if (hit) { if (event && event.waitUntil) event.waitUntil(guarded); return fromCache(hit); }
        return first === 'timeout' ? net.catch(function () { return fallback(); }) : fallback();   // لا نسخة: انتظر الشبكة أو أظهر البديل
      });
    });
  });
}
function staleWhileRevalidate(req, key, event) {
  return caches.open(SHELL).then(function (cache) {
    return cache.match(key).then(function (hit) {
      var net = fetch(req).then(function (res) { if (res && res.ok) cache.put(key, res.clone()); return res; }).catch(function () { return hit; });
      if (hit) { if (event && event.waitUntil) event.waitUntil(net); return hit; }
      return net;
    });
  });
}
if (typeof self !== 'undefined' && self.addEventListener && typeof ServiceWorkerGlobalScope !== 'undefined') {
  self.addEventListener('install', function (e) {
    e.waitUntil(caches.open(SHELL).then(function (c) { return Promise.all(ASSETS.map(function (a) { return c.add(a).catch(function () {}); })); }).then(function () { return self.skipWaiting(); }));
  });
  self.addEventListener('activate', function (e) {
    e.waitUntil(caches.keys().then(function (ks) { return Promise.all(ks.filter(function (k) { return (k.indexOf('ssuae-') === 0) && k !== SHELL && k !== DATA; }).map(function (k) { return caches.delete(k); })); }).then(function () { return self.clients.claim(); }));
  });
  self.addEventListener('message', function (e) { if (e.data === 'SKIP_WAITING') self.skipWaiting(); });
  self.addEventListener('fetch', function (event) {
    var req = event.request, u = new URL(req.url), kind = route(u.pathname, u.search, req.method, req.mode, u.origin === self.location.origin);
    if (kind === 'bypass') return;
    if (kind === 'page') { var key = new Request(u.origin + pageKey(u.pathname)); event.respondWith(networkFirst(req, key, SHELL, PAGE_TIMEOUT, offlineHtml, event)); }
    else if (kind === 'data') event.respondWith(networkFirst(req, new Request(u.origin + u.pathname), DATA, DATA_TIMEOUT, function () { return new Response('{"error":"offline"}', { status: 503, headers: { 'Content-Type': 'application/json' } }); }, event));
    else if (kind === 'asset') event.respondWith(staleWhileRevalidate(req, new Request(u.origin + u.pathname), event));
  });
}
if (typeof module !== 'undefined') module.exports = { route: route, pageKey: pageKey, networkFirst: networkFirst, staleWhileRevalidate: staleWhileRevalidate, offlineHtml: offlineHtml, _v: function () { return { SHELL: SHELL, DATA: DATA }; } };
