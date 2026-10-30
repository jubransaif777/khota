/* إعلانات الإدارة: شريط أعلى كل صفحة عامة + نافذة منبثقة اختيارية (مرة واحدة لكل جهاز). المصدر: لوحة الإدارة ← المحتوى. */
(function () {
  if (window.__scInit) return; window.__scInit = true;
  var API = '/.netlify/functions/content', K_DIS = 'ssuae_sc_dismiss', K_SEEN = 'ssuae_sc_seen';
  var ICON = { info: '📢', important: '⚠️', urgent: '🚨' };
  function load(k) { try { var v = JSON.parse(localStorage.getItem(k) || '[]'); return Array.isArray(v) ? v : []; } catch (e) { return []; } }
  function save(k, v) { try { localStorage.setItem(k, JSON.stringify(v.slice(-60))); } catch (e) {} }
  var css = '.sc-wrap{display:flex;flex-direction:column;gap:10px;margin:12px 0 0;font-family:inherit}' +
    '.sc-b{position:relative;display:flex;gap:11px;align-items:flex-start;padding:13px 14px;border-radius:var(--radius,16px);border:1px solid;color:var(--ink,#15302D);box-shadow:var(--shadow,none);animation:scin .35s ease}' +
    '@keyframes scin{from{opacity:0;transform:translateY(-6px)}to{opacity:1;transform:none}}' +
    '.sc-info{background:var(--primary-soft,#E1EEEC);border-color:rgba(14,90,84,.28)}' +
    '.sc-important{background:var(--alert-soft,#FBEEDA);border-color:#D9A45A}' +
    '.sc-urgent{background:#FDECEA;border-color:#E08A80;color:#6B1A12}' +
    '@media (prefers-color-scheme:dark){.sc-urgent{background:#3A1815;border-color:#8E3A31;color:#FFD9D4}}' +
    '.sc-ic{font-size:22px;line-height:1.25;flex:0 0 auto}.sc-tx{flex:1;min-width:0}' +
    '.sc-t{display:block;font-size:14.5px;font-weight:700;margin-bottom:2px}' +
    '.sc-m{font-size:14.5px;line-height:1.75;white-space:pre-line;overflow-wrap:anywhere}' +
    '.sc-x{border:none;background:none;color:inherit;opacity:.55;font-size:16px;cursor:pointer;padding:2px 4px;line-height:1;flex:0 0 auto}.sc-x:hover{opacity:1}' +
    '.sc-scrim{position:fixed;inset:0;background:rgba(10,20,18,.55);z-index:90;display:flex;align-items:center;justify-content:center;padding:22px;animation:scin .25s ease}' +
    '.sc-modal{max-width:380px;width:100%;background:var(--surface,#fff);color:var(--ink,#15302D);border:1px solid var(--line,#DBE4E1);border-radius:22px;padding:22px 20px 14px;text-align:center;box-shadow:0 24px 60px -20px rgba(0,0,0,.5);font-family:inherit}' +
    '.sc-modal .sc-big{font-size:38px;line-height:1.2}.sc-modal h3{margin:8px 0 6px;font-size:18px}.sc-modal p{margin:0 0 14px;font-size:15px;line-height:1.8;white-space:pre-line;overflow-wrap:anywhere}' +
    '.sc-ok{width:100%;border:none;border-radius:14px;background:var(--primary,#0E5A54);color:#fff;font-family:inherit;font-size:15px;font-weight:700;padding:12px;cursor:pointer}';
  var st = document.createElement('style'); st.id = 'sc-css'; st.textContent = css; document.head.appendChild(st);

  function mk(tag, cls, txt) { var e = document.createElement(tag); if (cls) e.className = cls; if (txt != null) e.textContent = txt; return e; }
  function banners(list) {
    var app = document.querySelector('.app') || document.body, dis = load(K_DIS);
    var show = list.filter(function (a) { return a.level === 'urgent' || dis.indexOf(a.id) < 0; });
    if (!show.length) return;
    var wrap = mk('div', 'sc-wrap'); wrap.id = 'scWrap'; wrap.setAttribute('role', 'region'); wrap.setAttribute('aria-label', 'إعلانات المدرسة');
    show.forEach(function (a) {
      var b = mk('div', 'sc-b sc-' + a.level); b.setAttribute('data-id', a.id);
      b.appendChild(mk('span', 'sc-ic', ICON[a.level] || ICON.info));
      var tx = mk('div', 'sc-tx'); if (a.title) tx.appendChild(mk('b', 'sc-t', a.title)); tx.appendChild(mk('div', 'sc-m', a.text)); b.appendChild(tx);
      if (a.level !== 'urgent') { var x = mk('button', 'sc-x', '✕'); x.type = 'button'; x.setAttribute('aria-label', 'إخفاء الإعلان');
        x.addEventListener('click', function () { var d = load(K_DIS); d.push(a.id); save(K_DIS, d); b.remove(); if (!wrap.children.length) wrap.remove(); }); b.appendChild(x); }
      wrap.appendChild(b);
    });
    var nav = app.querySelector('.topnav'); if (nav && nav.parentNode === app) nav.insertAdjacentElement('afterend', wrap); else app.insertBefore(wrap, app.firstChild);
  }
  function popups(list) {
    var seen = load(K_SEEN), queue = list.filter(function (a) { return a.popup && seen.indexOf(a.id) < 0; });
    function next() {
      var a = queue.shift(); if (!a) return;
      var s = mk('div', 'sc-scrim'), m = mk('div', 'sc-modal'); m.setAttribute('role', 'dialog'); m.setAttribute('aria-modal', 'true');
      m.appendChild(mk('div', 'sc-big', ICON[a.level] || ICON.info)); m.appendChild(mk('h3', null, a.title || 'إعلان')); m.appendChild(mk('p', null, a.text));
      var ok = mk('button', 'sc-ok', 'حسناً'); ok.type = 'button';
      ok.addEventListener('click', function () { var d = load(K_SEEN); d.push(a.id); save(K_SEEN, d); s.remove(); next(); });
      m.appendChild(ok); s.appendChild(m); document.body.appendChild(s); ok.focus();
    }
    next();
  }
  function run() {
    fetch(API).then(function (r) { return r.json(); }).then(function (j) {
      if (!j || !j.ok) return; window.SiteContent = j;
      if (j.announcements && j.announcements.length) { banners(j.announcements); popups(j.announcements); }
      try { document.dispatchEvent(new CustomEvent('sitecontent', { detail: j })); } catch (e) {}
    }).catch(function () {});
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', run); else run();
})();
