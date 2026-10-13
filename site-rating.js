/* بطاقة «قيّم الموقع» — تُضاف تلقائياً في أسفل الصفحات العامة. تقييم مجهول: نجوم 1–5 + ملاحظة اختيارية. */
(function () {
  if (window.__ratingInit) return; window.__ratingInit = true;
  var KEY = 'ssuae_rating_v1', API = '/.netlify/functions/rating', LABELS = ['', 'ضعيف', 'مقبول', 'جيد', 'جيد جداً', 'ممتاز'];
  function load() { try { return JSON.parse(localStorage.getItem(KEY) || '{}'); } catch (e) { return {}; } }
  function save(o) { try { localStorage.setItem(KEY, JSON.stringify(o)); } catch (e) {} }
  function uuid() { try { if (window.crypto && crypto.randomUUID) return crypto.randomUUID(); } catch (e) {}
    var s = ''; for (var i = 0; i < 32; i++) s += Math.floor(Math.random() * 16).toString(16); return s; }
  function pageKey() { var p = location.pathname.replace(/index\.html$/, '').replace(/\.html$/, ''); if (p.length > 1) p = p.replace(/\/$/, ''); return p || '/'; }

  var css = '.rt-card{margin:22px 0 6px;padding:18px 16px 16px;background:var(--surface,#fff);border:1px solid var(--line,#DBE4E1);border-radius:var(--radius,16px);box-shadow:var(--shadow,0 1px 2px rgba(16,45,42,.05));text-align:center;font-family:inherit;color:var(--ink,#15302D)}' +
    '.rt-card *{box-sizing:border-box;font-family:inherit}' +
    '.rt-t{font-size:16px;font-weight:700;margin:0}.rt-s2{font-size:12px;color:var(--muted,#5E706C);margin:3px 0 10px}' +
    '.rt-stars{display:flex;justify-content:center;gap:4px;direction:rtl}' +
    '.rt-star{width:52px;height:52px;border:none;background:none;cursor:pointer;font-size:38px;line-height:1;color:var(--line,#DBE4E1);padding:0;transition:color .15s,transform .15s;-webkit-tap-highlight-color:transparent}' +
    '.rt-star.on,.rt-star.hv{color:#E0A526}.rt-star:active{transform:scale(.88)}.rt-star.pop{animation:rtpop .35s ease}' +
    '@keyframes rtpop{0%{transform:scale(.8)}60%{transform:scale(1.25)}100%{transform:scale(1)}}' +
    '.rt-lab{min-height:22px;font-size:14px;font-weight:600;color:var(--primary-ink,#0B403B);margin-top:2px}' +
    '.rt-form{display:none;margin-top:10px;text-align:right}.rt-form.show{display:block}' +
    '.rt-ta{width:100%;min-height:78px;resize:vertical;border:1px solid var(--line,#DBE4E1);border-radius:12px;padding:10px 12px;font-size:14px;background:var(--bg,#F1F5F4);color:inherit;line-height:1.6}' +
    '.rt-ta:focus{outline:2px solid var(--primary,#0E5A54);outline-offset:1px;background:var(--surface,#fff)}' +
    '.rt-row{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-top:8px}.rt-cnt{font-size:12px;color:var(--muted,#5E706C)}' +
    '.rt-send{border:none;border-radius:12px;background:var(--primary,#0E5A54);color:#fff;font-size:14px;font-weight:700;padding:11px 20px;cursor:pointer}.rt-send:disabled{opacity:.6;cursor:default}' +
    '.rt-err{display:none;margin-top:8px;font-size:12.5px;color:#B03A2E;text-align:center}.rt-err.show{display:block}' +
    '.rt-done{display:none}.rt-done.show{display:block}.rt-done b{display:block;font-size:15px;margin:2px 0 4px}.rt-done .rt-ro{pointer-events:none}' +
    '.rt-edit{border:none;background:none;color:var(--primary,#0E5A54);font-size:13px;font-weight:600;cursor:pointer;text-decoration:underline;margin-top:4px}' +
    '.rt-hp{position:absolute;left:-9999px;width:1px;height:1px;opacity:0}' +
    '@media (prefers-reduced-motion:reduce){.rt-star,.rt-star.pop{transition:none;animation:none}}';
  var st = document.createElement('style'); st.id = 'rating-css'; st.textContent = css; document.head.appendChild(st);

  var card = document.createElement('section'); card.className = 'rt-card'; card.setAttribute('aria-label', 'تقييم الموقع');
  card.innerHTML =
    '<div class="rt-ask"><p class="rt-t">ما رأيك بالموقع؟</p><p class="rt-s2">تقييمك مجهول، ويساعدنا على تطوير الموقع</p>' +
    '<div class="rt-stars" role="group" aria-label="التقييم من 1 إلى 5"></div><div class="rt-lab" aria-live="polite"></div>' +
    '<div class="rt-form"><textarea class="rt-ta" maxlength="300" placeholder="ملاحظتك أو اقتراحك (اختياري)" aria-label="ملاحظتك"></textarea>' +
    '<input class="rt-hp" type="text" name="website" tabindex="-1" autocomplete="off" aria-hidden="true">' +
    '<div class="rt-row"><span class="rt-cnt">0 / 300</span><button class="rt-send" type="button">إرسال التقييم</button></div></div>' +
    '<div class="rt-err" role="alert"></div></div>' +
    '<div class="rt-done"><div class="rt-stars rt-ro" aria-hidden="true"></div><b>شكراً لك! وصل تقييمك 💚</b><button class="rt-edit" type="button">تعديل التقييم</button></div>';

  var $ = function (s) { return card.querySelector(s); };
  var ask = $('.rt-ask'), done = $('.rt-done'), starsBox = $('.rt-ask .rt-stars'), roBox = $('.rt-done .rt-stars'),
      lab = $('.rt-lab'), form = $('.rt-form'), ta = $('.rt-ta'), cnt = $('.rt-cnt'), send = $('.rt-send'), err = $('.rt-err'), hp = $('.rt-hp');
  var state = load(), val = 0, btns = [];

  for (var i = 1; i <= 5; i++) {
    (function (n) {
      var b = document.createElement('button'); b.type = 'button'; b.className = 'rt-star'; b.textContent = '★';
      b.setAttribute('aria-label', n + ' من 5 — ' + LABELS[n]);
      b.addEventListener('click', function () { pick(n, true); });
      b.addEventListener('mouseenter', function () { paint(n, true); });
      b.addEventListener('mouseleave', function () { paint(val, false); });
      starsBox.appendChild(b); btns.push(b);
      var r = document.createElement('span'); r.className = 'rt-star'; r.textContent = '★'; roBox.appendChild(r);
    })(i);
  }
  function paint(n, hover) { btns.forEach(function (b, k) { if (hover) { b.classList.remove('on'); b.classList.toggle('hv', k < n); } else { b.classList.remove('hv'); b.classList.toggle('on', k < n); } }); }
  function paintRo(n) { Array.prototype.forEach.call(roBox.children, function (r, k) { r.classList.toggle('on', k < n); }); }
  function pick(n, user) {
    val = n; btns.forEach(function (b, k) { b.classList.remove('hv'); b.classList.toggle('on', k < n); });
    if (user) { var b = btns[n - 1]; b.classList.remove('pop'); void b.offsetWidth; b.classList.add('pop'); }
    lab.textContent = LABELS[n]; form.classList.add('show'); err.classList.remove('show'); err.textContent = '';
  }
  ta.addEventListener('input', function () { cnt.textContent = ta.value.length + ' / 300'; });
  function showDone() { paintRo(val); ask.style.display = 'none'; done.classList.add('show'); }
  function showAsk() { done.classList.remove('show'); ask.style.display = ''; pick(val, false); ta.value = state.comment || ''; cnt.textContent = ta.value.length + ' / 300'; }
  $('.rt-edit').addEventListener('click', showAsk);

  send.addEventListener('click', function () {
    if (!val) return;
    if (!state.id) { state.id = uuid(); save(state); }
    send.disabled = true; var old = send.textContent; send.textContent = 'جارٍ الإرسال…'; err.classList.remove('show');
    fetch(API, { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id: state.id, stars: val, comment: ta.value, page: pageKey(), website: hp.value }) })
      .then(function (r) { return r.json().then(function (j) { return { s: r.status, j: j }; }); })
      .then(function (x) {
        send.disabled = false; send.textContent = old;
        if (x.j && x.j.ok) { state.stars = val; state.comment = ta.value; save(state); showDone(); }
        else { err.textContent = x.s === 429 ? 'تجاوزت عدد المحاولات، حاول لاحقاً.' : 'تعذّر الإرسال الآن. حاول مرة أخرى بعد قليل.'; err.classList.add('show'); }
      })
      .catch(function () { send.disabled = false; send.textContent = old; err.textContent = 'تعذّر الإرسال الآن. تحقّق من الاتصال وحاول مرة أخرى.'; err.classList.add('show'); });
  });

  function mount() {
    var app = document.querySelector('.app'); var anchor = app && app.querySelector('footer.foot, .foot, .credits');
    if (app && anchor && anchor.parentNode === app) app.insertBefore(card, anchor); else (app || document.body).appendChild(card);
    if (state.stars >= 1 && state.stars <= 5) { val = state.stars; paintRo(val); ask.style.display = 'none'; done.classList.add('show'); }
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mount); else mount();
})();
