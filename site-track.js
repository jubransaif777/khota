/* إحصاءات استخدام مجهولة: عدد الزيارات فقط. بلا ملفات تتبّع ولا هوية ولا عنوان IP، وتحترم «عدم التتبّع». */
(function () {
  if (window.__trk) return; window.__trk = 1;
  var API = '/.netlify/functions/track';
  try { if (navigator.doNotTrack === '1' || localStorage.getItem('ssuae_notrack') === '1') { window.SiteTrack = { event: function () {} }; return; } } catch (e) {}
  var P = (location.pathname.replace(/\/index\.html$/, '/').replace(/\.html$/, '').replace(/(.)\/+$/, '$1')) || '/';
  function send(o) {
    try { var body = JSON.stringify(o);
      if (navigator.sendBeacon && navigator.sendBeacon(API, new Blob([body], { type: 'application/json' }))) return;
      fetch(API, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: body, keepalive: true }).catch(function () {});
    } catch (e) {}
  }
  function ls(k, v) { try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) {} return null; }
  var today = new Date().toLocaleDateString('en-CA', { timeZone: 'Asia/Dubai' });
  var standalone = false; try { standalone = (window.matchMedia && matchMedia('(display-mode: standalone)').matches) || navigator.standalone === true; } catch (e) {}
  window.SiteTrack = { event: function (e) { send({ t: 'ev', e: e, p: P, s: standalone ? 1 : 0 }); } };
  function hit() {
    var first = ls('ssuae_trk_day') !== today, isNew = ls('ssuae_trk_dev') === null;
    if (first) ls('ssuae_trk_day', today); if (isNew) ls('ssuae_trk_dev', '1');
    var g = ''; try { g = (window.SiteClass && SiteClass.group('')) || ''; } catch (e) {}
    send({ t: 'hit', p: P, f: first ? 1 : 0, n: isNew ? 1 : 0, s: standalone ? 1 : 0, g: first ? g : '' });
  }
  if (document.readyState === 'complete') setTimeout(hit, 600); else window.addEventListener('load', function () { setTimeout(hit, 600); });
})();
