/* رسم شهادة الشكر والتقدير (A4 أفقي) وطباعتها. تستعملها صفحة /certs لولي الأمر ولوحة الإدارة للمعاينة والطباعة.
   SiteCert.html(c, r)  c: {type,date,name,section,text}  r: {school,ministry,sigs:[{title,name}]} */
(function () {
  if (window.SiteCert) return;
  var CSS = ".certwrap{container-type:inline-size;width:100%;border-radius:6px;overflow:hidden;box-shadow:0 2px 14px -6px rgba(0,0,0,.4);margin-bottom:10px;background:#FFFDF7}.cert{--g:#0E5A54;--gd:#B8923A;font-size:min(1vw,6px);position:relative;width:100%;aspect-ratio:297/210;background:#FFFDF7;color:#1B2B29;font-family:'Amiri','Readex Pro',serif;text-align:center}@supports (width:1cqw){.cert{font-size:1cqw}}.c-frame{position:absolute;inset:2.2em;border:.9em solid var(--g);padding:.55em}.c-inner{position:relative;width:100%;height:100%;border:.25em solid var(--gd);display:flex;flex-direction:column;align-items:center;padding:2em 5em 1.7em}.c-corner{position:absolute;width:3.2em;height:3.2em;border:.25em solid var(--gd);background:#FFFDF7}.c-corner.tl{top:-1.7em;right:-1.7em}.c-corner.tr{top:-1.7em;left:-1.7em}.c-corner.bl{bottom:-1.7em;right:-1.7em}.c-corner.br{bottom:-1.7em;left:-1.7em}.c-corner::after{content:\"\";position:absolute;inset:.7em;background:var(--g)}.c-head{width:100%;display:flex;align-items:center;justify-content:space-between;gap:2em}.c-head>div{flex:1 1 0;min-width:0;line-height:1.5}.c-min{font-size:2.5em;color:#4B5E5B;text-align:start}.c-school{font-size:2.8em;font-weight:700;color:var(--g);text-align:end}.c-crest{width:9.4em;height:9.4em;object-fit:contain;flex:0 0 auto}.c-title{font-family:'Reem Kufi','Amiri',serif;font-weight:700;font-size:6.2em;color:var(--gd);line-height:1.2;margin:.1em 0 0;letter-spacing:.02em}.c-lead{font-size:2.5em;color:#4B5E5B;line-height:1.4}.c-name{font-weight:700;font-size:5.3em;color:var(--g);line-height:1.3;padding:0 1.5em;border-bottom:.14em solid var(--gd);margin:.05em 0 .05em;min-width:44%;max-width:100%}.c-name.n2{font-size:4.5em}.c-name.n3{font-size:3.8em}.c-sec{font-size:2.6em;color:#4B5E5B;line-height:1.4;min-height:1.4em;margin-bottom:.2em}.c-body{font-size:3.1em;line-height:1.65;max-width:92%;flex:1;min-height:0;display:flex;align-items:center;justify-content:center;white-space:normal;overflow:hidden}.c-body.b2{font-size:2.7em}.c-body.b3{font-size:2.3em;line-height:1.6}.c-body.b4{font-size:2em;line-height:1.55}.c-body.b5{font-size:1.6em;line-height:1.45}.c-body.b6{font-size:1.4em;line-height:1.4}.c-foot{width:100%;display:flex;align-items:flex-start;justify-content:space-between;gap:3em;margin-top:.3em}.c-sig,.c-date{flex:1 1 0;min-width:0;font-size:2.35em;line-height:1.45}.c-sp{height:1em}.c-sig i,.c-date i{display:block;height:0;border-top:1px solid #6C7B78;width:78%;margin:0 auto .3em}.c-date i{border-top-color:transparent}.c-sig b,.c-date b{display:block;color:var(--g);min-height:1.45em}.c-sig span,.c-date span{display:block;color:#4B5E5B;font-size:.92em}";
  var PRINT = '#certPrintSheet{display:none}@media print{@page{size:A4 landscape;margin:0}html,body{background:#fff !important;-webkit-print-color-adjust:exact;print-color-adjust:exact}body.cert-printing > *:not(#certPrintSheet){display:none !important}body.cert-printing #certPrintSheet{display:block;width:297mm;height:210mm;overflow:hidden}body.cert-printing #certPrintSheet .certwrap{box-shadow:none;border-radius:0;margin:0;width:297mm}}';
  var st = document.createElement('style'); st.id = 'siteCertCss'; st.textContent = CSS + PRINT; (document.head || document.documentElement).appendChild(st);
  var MONTHS = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function fdate(s) { var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ''); return m ? (+m[3]) + ' ' + MONTHS[+m[2] - 1] + ' ' + m[1] : ''; }
  function title(t) { t = String(t || '').trim(); return /^شهادة/.test(t) ? t : 'شهادة ' + t; }
  function text(c) { return String(c.body != null ? c.body : c.text).replace(/\{الاسم\}/g, c.name || '').replace(/\{الصف\}/g, c.section || ''); }
  function schoolHtml(t) { var w = String(t || '').trim().split(/\s+/); return w.length > 3 ? esc(w.slice(0, 2).join(' ')) + '<br>' + esc(w.slice(2).join(' ')) : esc(t); }   // سطران: «مدرسة المعيريض» / «للحلقة الثانية بنين»
  function sp() { return '<div class="c-sp"></div>'; }
  function html(c, r) {
    var tx = c.text != null ? c.text : text(c), ss = r.sigs || [];
    var blocks = ss.map(function (s) { return '<div class="c-sig">' + sp() + '<i></i><b>' + (s.name ? esc(s.name) : '&nbsp;') + '</b><span>' + esc(s.title) + '</span></div>'; });
    var dt = '<div class="c-date">' + sp() + '<i></i><b>' + esc(fdate(c.date)) + '</b><span>التاريخ</span></div>';
    blocks.splice(ss.length === 2 ? 1 : ss.length, 0, dt);                       // التاريخ في الوسط بين توقيعين، وإلا في الطرف
    var min = String(r.ministry || '').split(/\s*[—–]\s*/).filter(Boolean).map(esc).join('<br>');
    return '<div class="certwrap"><div class="cert"><div class="c-frame"><div class="c-inner">' +
      '<span class="c-corner tl"></span><span class="c-corner tr"></span><span class="c-corner bl"></span><span class="c-corner br"></span>' +
      '<div class="c-head"><div class="c-min">' + min + '</div><img class="c-crest" alt="" src="/crest.png"><div class="c-school">' + schoolHtml(r.school) + '</div></div>' +
      '<div class="c-title">' + esc(title(c.type)) + '</div><div class="c-lead">تُمنح هذه الشهادة للطالب</div>' +
      '<div class="c-name' + ((c.name || '').length > 30 ? ' n3' : (c.name || '').length > 20 ? ' n2' : '') + '">' + esc(c.name) + '</div><div class="c-sec">' + esc(c.section || '') + '</div>' +
      '<div class="c-body' + (tx.length > 460 ? ' b6' : tx.length > 340 ? ' b5' : tx.length > 260 ? ' b4' : tx.length > 170 ? ' b3' : tx.length > 110 ? ' b2' : '') + '"><div>' + esc(tx).replace(/\n/g, '<br>') + '</div></div>' +
      '<div class="c-foot">' + blocks.join('') + '</div>' +
      '</div></div></div></div>';
  }
  function print(c, r) {
    var box = document.getElementById('certPrintSheet'); if (!box) { box = document.createElement('div'); box.id = 'certPrintSheet'; box.setAttribute('aria-hidden', 'true'); document.body.appendChild(box); }
    box.innerHTML = html(c, r); document.body.classList.add('cert-printing');
    var done = function () { document.body.classList.remove('cert-printing'); box.innerHTML = ''; window.removeEventListener('afterprint', done); };
    window.addEventListener('afterprint', done); setTimeout(function () { window.print(); }, 80);
  }
  window.SiteCert = { html: html, print: print, title: title, fdate: fdate, text: text, esc: esc };
})();
