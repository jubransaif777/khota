/* رسم شهادة الشكر والتقدير (A4 أفقي) وطباعتها. تستعملها صفحة /certs لولي الأمر ولوحة الإدارة للمعاينة والطباعة.
   SiteCert.html(c, r)  c: {type,date,name,section,text}  r: {school,ministry,sigs:[{title,name}]} */
(function () {
  if (window.SiteCert) return;
  var CSS = "  .certwrap{container-type:inline-size;width:100%;border-radius:6px;overflow:hidden;box-shadow:0 2px 14px -6px rgba(0,0,0,.4);margin-bottom:10px;background:#FFFDF7}.cert{--g:#0E5A54;--gd:#B8923A;font-size:min(1vw,6px);position:relative;width:100%;aspect-ratio:297/210;background:#FFFDF7;color:#1B2B29;font-family:'Amiri','Readex Pro',serif;text-align:center}@supports (width:1cqw){.cert{font-size:1cqw}}.c-frame{position:absolute;inset:2.2em;border:.9em solid var(--g);padding:.55em}.c-inner{position:relative;width:100%;height:100%;border:.25em solid var(--gd);display:flex;flex-direction:column;align-items:center;padding:1.9em 6em 1.6em}.c-corner{position:absolute;width:3.2em;height:3.2em;border:.25em solid var(--gd);background:#FFFDF7}.c-corner.tl{top:-1.7em;right:-1.7em}.c-corner.tr{top:-1.7em;left:-1.7em}.c-corner.bl{bottom:-1.7em;right:-1.7em}.c-corner.br{bottom:-1.7em;left:-1.7em}.c-corner::after{content:\"\";position:absolute;inset:.7em;background:var(--g)}.c-crest{width:6.2em;height:6.2em;object-fit:contain;border-radius:1.3em;margin-bottom:.4em}.c-min{font-size:1.9em;color:#4B5E5B;line-height:1.4}.c-school{font-size:2.6em;font-weight:700;color:var(--g);line-height:1.4}.c-title{font-family:'Reem Kufi','Amiri',serif;font-weight:700;font-size:5.6em;color:var(--gd);line-height:1.25;margin:.15em 0 .05em;letter-spacing:.02em}.c-lead{font-size:2.2em;color:#4B5E5B;line-height:1.4}.c-name{font-weight:700;font-size:4.6em;color:var(--g);line-height:1.4;padding:0 3em;border-bottom:.14em solid var(--gd);margin:.05em 0 .05em;min-width:40%}.c-sec{font-size:2.1em;color:#4B5E5B;line-height:1.4;min-height:1.4em;margin-bottom:.3em}.c-body{font-size:2.6em;line-height:1.7;max-width:84%;flex:1;min-height:0;display:flex;align-items:center;justify-content:center;white-space:normal;overflow:hidden}.c-body.b2{font-size:2.2em}.c-body.b3{font-size:1.85em;line-height:1.6}.c-foot{width:100%;display:flex;justify-content:space-around;align-items:flex-end;gap:2em;margin-top:.2em}.c-sig{flex:0 0 34%;font-size:2.1em;line-height:1.45}.c-sig i{display:block;height:0;border-top:1px solid #6C7B78;margin:1.9em 0 .3em}.c-sig b{display:block;color:var(--g)} .c-sig span{color:#4B5E5B}.c-date{font-size:2.1em;color:#4B5E5B;line-height:1.5;min-width:20%}.c-date b{color:var(--g);display:block}";
  var PRINT = '#certPrintSheet{display:none}@media print{@page{size:A4 landscape;margin:0}html,body{background:#fff !important;-webkit-print-color-adjust:exact;print-color-adjust:exact}body.cert-printing > *:not(#certPrintSheet){display:none !important}body.cert-printing #certPrintSheet{display:block;width:297mm;height:210mm;overflow:hidden}body.cert-printing #certPrintSheet .certwrap{box-shadow:none;border-radius:0;margin:0;width:297mm}}';
  var st = document.createElement('style'); st.id = 'siteCertCss'; st.textContent = CSS + PRINT; (document.head || document.documentElement).appendChild(st);
  var MONTHS = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function fdate(s) { var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ''); return m ? (+m[3]) + ' ' + MONTHS[+m[2] - 1] + ' ' + m[1] : ''; }
  function title(t) { t = String(t || '').trim(); return /^شهادة/.test(t) ? t : 'شهادة ' + t; }
  function text(c) { return String(c.body != null ? c.body : c.text).replace(/\{الاسم\}/g, c.name || '').replace(/\{الصف\}/g, c.section || ''); }
  function html(c, r) {
    var tx = c.text != null ? c.text : text(c);
    var sigs = (r.sigs || []).map(function (s) { return '<div class="c-sig"><i></i><span>' + esc(s.title) + '</span>' + (s.name ? '<b>' + esc(s.name) + '</b>' : '') + '</div>'; }).join('');
    return '<div class="certwrap"><div class="cert"><div class="c-frame"><div class="c-inner">' +
      '<span class="c-corner tl"></span><span class="c-corner tr"></span><span class="c-corner bl"></span><span class="c-corner br"></span>' +
      '<img class="c-crest" alt="" src="/crest.png">' +
      (r.ministry ? '<div class="c-min">' + esc(r.ministry) + '</div>' : '') + '<div class="c-school">' + esc(r.school) + '</div>' +
      '<div class="c-title">' + esc(title(c.type)) + '</div><div class="c-lead">تُمنح هذه الشهادة للطالب</div>' +
      '<div class="c-name">' + esc(c.name) + '</div><div class="c-sec">' + esc(c.section || '') + '</div>' +
      '<div class="c-body' + (tx.length > 260 ? ' b3' : tx.length > 120 ? ' b2' : '') + '"><div>' + esc(tx).replace(/\n/g, '<br>') + '</div></div>' +
      '<div class="c-foot"><div class="c-date">التاريخ<b>' + esc(fdate(c.date)) + '</b></div>' + sigs + '</div>' +
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
