/* تثبيت الموقع كتطبيق على الهاتف + تسجيل خدمة العمل بلا إنترنت + شريط «غير متصل».
   لتعطيل الخدمة فوراً: أضف ?nosw=1 إلى أي رابط. */
(function () {
  if (window.__pwa) return; window.__pwa = 1;
  var DAY = 86400000, KD = 'ssuae_pwa_dismiss', KV = 'ssuae_pwa_visits';
  function ls(k, v) { try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) {} return null; }
  function track(e) { try { window.SiteTrack && SiteTrack.event(e); } catch (x) {} }
  var standalone = false; try { standalone = (window.matchMedia && matchMedia('(display-mode: standalone)').matches) || navigator.standalone === true; } catch (e) {}

  // ---- مفتاح الإيقاف الطارئ ----
  if (/[?&]nosw=1/.test(location.search) && 'serviceWorker' in navigator) {
    navigator.serviceWorker.getRegistrations().then(function (rs) { return Promise.all(rs.map(function (r) { return r.unregister(); })); })
      .then(function () { return window.caches ? caches.keys().then(function (ks) { return Promise.all(ks.map(function (k) { return caches.delete(k); })); }) : 0; })
      .then(function () { location.replace(location.pathname); });
    return;
  }
  // ---- التسجيل ----
  if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) {
    window.addEventListener('load', function () { navigator.serviceWorker.register('/sw.js').catch(function () {}); });
  }
  // ---- شريط عدم الاتصال ----
  var bar = null;
  function offline(on) {
    if (on && !bar) { bar = document.createElement('div'); bar.setAttribute('role', 'status');
      bar.style.cssText = 'position:fixed;top:0;inset-inline:0;z-index:80;background:#8A520F;color:#fff;text-align:center;font:600 13px/1.6 inherit;padding:6px 10px;font-family:inherit';
      bar.textContent = '📡 أنت غير متصل — تُعرض آخر نسخة محفوظة على جهازك'; document.body.appendChild(bar); }
    else if (!on && bar) { bar.remove(); bar = null; }
  }
  window.addEventListener('offline', function () { offline(true); }); window.addEventListener('online', function () { offline(false); });
  if (navigator.onLine === false) { if (document.body) offline(true); else window.addEventListener('DOMContentLoaded', function () { offline(true); }); }

  // ---- التثبيت ----
  var deferred = null, card = null, shown = false;
  var ua = navigator.userAgent || '', isIOS = /iphone|ipad|ipod/i.test(ua) && /safari/i.test(ua) && !/crios|fxios|edgios|opios/i.test(ua);
  var visits = (parseInt(ls(KV) || '0', 10) || 0) + 1; ls(KV, String(visits));
  function dismissed() { var d = {}; try { d = JSON.parse(ls(KD) || '{}'); } catch (e) {} return d; }
  function suppressed() { var d = dismissed(); if (!d.t) return false; return Date.now() - d.t < ((d.n || 1) >= 2 ? 90 : 14) * DAY; }
  function build(ios) {
    var css = '.pwa-c{position:fixed;inset-inline:12px;bottom:12px;max-width:480px;margin:0 auto;z-index:60;background:var(--surface,#fff);color:var(--ink,#15302D);border:1px solid var(--line,#DBE4E1);border-radius:18px;box-shadow:0 14px 40px -12px rgba(0,0,0,.4);padding:14px 14px 12px;display:flex;gap:12px;align-items:flex-start;font-family:inherit;animation:pwain .35s ease}' +
      '@keyframes pwain{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}.pwa-c img{width:46px;height:46px;border-radius:12px;flex:0 0 auto}.pwa-c b{display:block;font-size:14.5px;margin-bottom:2px}.pwa-c p{margin:0;font-size:12.5px;line-height:1.7;color:var(--muted,#5E706C)}' +
      '.pwa-a{display:flex;gap:8px;margin-top:10px}.pwa-a button{font-family:inherit;font-size:13px;font-weight:700;border-radius:11px;padding:8px 16px;cursor:pointer;border:1px solid var(--line,#DBE4E1);background:var(--surface,#fff);color:var(--ink,#15302D)}.pwa-a .go{background:var(--primary,#0E5A54);border-color:var(--primary,#0E5A54);color:#fff}';
    if (!document.getElementById('pwa-css')) { var s = document.createElement('style'); s.id = 'pwa-css'; s.textContent = css; document.head.appendChild(s); }
    var c = document.createElement('div'); c.className = 'pwa-c'; c.setAttribute('role', 'dialog'); c.setAttribute('aria-label', 'تثبيت التطبيق');
    c.innerHTML = '<img src="/icon-192.png" alt=""><div style="flex:1;min-width:0"><b>ثبّت الخطة الأسبوعية على هاتفك</b><p>' +
      (ios ? 'اضغط زرّ المشاركة <span aria-hidden="true">⬆️</span> في المتصفح ثم اختر «إضافة إلى الشاشة الرئيسية». تفتحها بنقرة وتظهر لك الخطة حتى بلا إنترنت.' : 'تفتحها بنقرة من الشاشة الرئيسية وتظهر لك آخر خطة حتى بلا إنترنت.') + '</p><div class="pwa-a">' + (ios ? '' : '<button class="go" type="button">تثبيت</button>') + '<button class="later" type="button">' + (ios ? 'حسناً' : 'لاحقاً') + '</button></div></div>';
    return c;
  }
  function hide(remember) { if (card) { card.remove(); card = null; } if (remember) { var d = dismissed(); ls(KD, JSON.stringify({ t: Date.now(), n: (d.n || 0) + 1 })); } }
  function show(ios) {
    if (shown || standalone || suppressed() || visits < 2 || !document.body) return; shown = true; card = build(ios); document.body.appendChild(card); track('install_prompt');
    card.querySelector('.later').addEventListener('click', function () { hide(true); });
    var go = card.querySelector('.go'); if (go) go.addEventListener('click', function () { api.promptInstall(); });
  }
  var api = window.SitePWA = {
    canInstall: function () { return !!deferred; },
    promptInstall: function () { if (!deferred) return Promise.resolve(false); var d = deferred; deferred = null; hide(false); d.prompt(); return (d.userChoice || Promise.resolve({})).then(function (r) { if (r && r.outcome === 'accepted') track('install'); else { var x = dismissed(); ls(KD, JSON.stringify({ t: Date.now(), n: (x.n || 0) + 1 })); } return !!(r && r.outcome === 'accepted'); }); }
  };
  window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); deferred = e; show(false); });
  window.addEventListener('appinstalled', function () { hide(false); track('install'); ls(KD, JSON.stringify({ t: Date.now() + 365 * DAY, n: 9 })); });
  if (isIOS && !standalone) window.addEventListener('load', function () { setTimeout(function () { show(true); }, 2500); });
})();
