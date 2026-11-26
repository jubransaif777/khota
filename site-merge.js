/* دمج الخطة مع مواعيد الاختبارات لثلاث أدوات: رسالة واتساب جاهزة، وملف تقويم (.ics)، ونموذج الملخص.
   يستعملها زرّ المشاركة والتقويم في صفحة الملخص، وتبويب «رسائل الواتساب» في لوحة الإدارة. */
(function () {
  var XS = { IS: ['إسلامية', '🕌'], AR: ['عربي', '📖'], SS: ['دراسات', '🌍'], E: ['إنجليزي', '🔤'], MA: ['رياضيات', '➗'], SC: ['علوم', '🔬'], PE: ['بدنية', '⚽'], DT: ['تصميم', '⚙️'], DR: ['مسرح', '🎭'], VA: ['فنية', '🎨'], AA: ['موسيقى', '🎵'], AI: ['ذكاء', '🤖'] };
  var DAYIDX = { MON: 0, TUE: 1, WED: 2, THU: 3, FRI: 4, SAT: 5, SUN: 6 };
  function codeOf(s) {
    var t = String(s || '').toLowerCase().replace(/[\u064B-\u0652\u0640]/g, '').replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه');
    if (/اسلام|islamic/.test(t)) return 'IS'; if (/عرب|arabic/.test(t)) return 'AR'; if (/انجليز|english/.test(t)) return 'E';
    if (/اجتماع|دراسات|social/.test(t)) return 'SS'; if (/ذكاء|\bai\b|ccdi/.test(t)) return 'AI'; if (/بدن|رياضيه|physical|\bpe\b/.test(t)) return 'PE';
    if (/رياضيات|math/.test(t)) return 'MA'; if (/علوم|science/.test(t)) return 'SC'; if (/سمع|موسيق|music/.test(t)) return 'AA';
    if (/بصر|فنون|فني|\bart\b/.test(t)) return 'VA'; if (/مسرح|drama/.test(t)) return 'DR'; if (/تصميم|تقني|design/.test(t)) return 'DT'; return null;
  }
  function dd(a, b) { return Math.abs((new Date(a + 'T00:00:00Z') - new Date(b + 'T00:00:00Z')) / 86400000); }
  function westify(s) { return String(s == null ? '' : s).replace(/[٠-٩]/g, function (d) { return '٠١٢٣٤٥٦٧٨٩'.indexOf(d); }); }

  // نموذج الأسبوع لمجموعة (6G): الدروس + الاختبارات (pick = '6G1' أو '' للكل)
  function model(o) {
    var plan = o.plan, ex = o.exams, group = o.group, wk = o.weekIndex, pick = o.pick || '', choices = o.choices || {};
    var w = plan.weeks[wk], ents = ((w.sections && w.sections[group]) || []).filter(function (x) { return !x.sec || !pick || pick.indexOf(group) !== 0 || x.sec === pick; });   // v3.4: شعبة محدّدة ← دروسها + المشترك
    var cl = ''; (plan.sectionOrder || []).forEach(function (s) { if (s.code === group) cl = s.label; });
    var m = { group: group, classLabel: cl || group, weekLabel: westify(w.label || ''), range: westify(w.range || ''), notices: w.notices || [], rows: [], items: [], others: [], hasExams: !!ex, subs: [], differs: false, pick: '', examWeek: null };
    var order = [], by = {}; ents.forEach(function (x) { if (!by[x.subject]) { by[x.subject] = []; order.push(x.subject); } by[x.subject].push(x); });
    var xw = null;
    if (ex && ex.weeks) { if (w.start) ex.weeks.forEach(function (x) { if (x.start && dd(x.start, w.start) <= 3 && (!xw || dd(x.start, w.start) < dd(xw.start, w.start))) xw = x; }); if (!xw && w.num != null) ex.weeks.forEach(function (x) { if (x.num === w.num) xw = x; }); }
    m.examWeek = xw;
    if (xw) {
      var subs = (ex.sectionOrder || []).filter(function (s) { return s.indexOf(group) === 0; }), idx = {}, map = {}, list = [], all = [];
      (ex.dayOrder || []).forEach(function (d, i) { idx[d[0]] = i; });
      function collect(only) { map = {}; list = []; subs.forEach(function (sec) { if (only && sec !== only) return; ((xw.sections && xw.sections[sec]) || []).forEach(function (d) { (d.exams || []).forEach(function (e) {
        var k = d.day + '|' + e.code + '|' + JSON.stringify(e.card || null); if (!map[k]) { map[k] = { day: d.day, dayLabel: d.label, code: e.code, subject: e.subject, card: e.card || null, secs: [] }; list.push(map[k]); } map[k].secs.push(sec); }); }); });
        list.sort(function (a, b) { return ((idx[a.day] || 0) - (idx[b.day] || 0)) || (a.code < b.code ? -1 : (a.code > b.code ? 1 : 0)); }); return list; }
      all = collect(''); m.subs = subs; m.differs = subs.length > 1 && all.some(function (i) { return i.secs.length < subs.length; });
      m.pick = (pick && subs.indexOf(pick) >= 0) ? pick : ''; m.items = m.pick ? collect(m.pick) : all;
    }
    var covered = {};
    order.forEach(function (subj) {
      var es = by[subj], ch = choices[subj], shown = es, code = codeOf(subj); if (code) covered[code] = 1;
      if (ch) { var one = es.filter(function (x) { return x.teacher === ch; }); if (one.length) shown = one; }
      m.rows.push({ subject: subj, emoji: es[0].emoji, code: code, shown: shown, exams: m.items.filter(function (i) { return i.code === code; }) });
    });
    m.others = m.items.filter(function (i) { return !covered[i.code]; });
    m.reminders = []; ents.forEach(function (x) { var t = x.reminder || x.note; if (t) m.reminders.push({ subject: x.subject, text: t }); if (x.homework) m.reminders.push({ subject: x.subject, text: 'واجب: ' + x.homework }); });
    return m;
  }
  function tag(m, i) { return (!m.pick && m.subs.length > 1 && i.secs.length < m.subs.length) ? ' (' + i.secs.join('، ') + ')' : ''; }
  function cut(s, n) { s = String(s || '').replace(/\s+/g, ' ').trim(); return s.length > n ? s.slice(0, n - 1) + '…' : s; }

  // ---------- رسالة واتساب ----------
  function whatsapp(m, o) {
    o = o || {}; var L = [], link = o.link || 'https://khotamaereid.net/summary';
    L.push('📋 *ملخص أسبوع ' + m.classLabel + (m.pick ? ' (' + m.pick + ')' : '') + '*'); L.push([m.weekLabel, m.range].filter(Boolean).join(' · '));
    if (o.notices !== false && m.notices.length) { L.push(''); L.push('📢 *إعلان الأسبوع*'); m.notices.forEach(function (t) { L.push(cut(t, 160)); }); }
    if (o.exams !== false) {
      L.push(''); L.push('📝 *الاختبارات*');
      if (!m.hasExams || !m.examWeek) L.push('لم تُنشر مواعيد اختبارات هذا الأسبوع بعد.');
      else if (!m.items.length) L.push('✅ لا اختبارات هذا الأسبوع.');
      else m.items.forEach(function (i) { var c = i.card || {}, bits = [c.title && cut(c.title, 60), c.pages && ('ص ' + c.pages), c.marks && ('الدرجة ' + c.marks)].filter(Boolean);
        L.push('• ' + i.dayLabel + ' — ' + (i.subject || (XS[i.code] || [i.code])[0]) + tag(m, i) + (bits.length ? ': ' + bits.join(' · ') : '')); });
    }
    if (o.reminders !== false && m.reminders.length) { L.push(''); L.push('🔔 *تذكيرات*'); m.reminders.slice(0, 6).forEach(function (r) { L.push('• ' + r.subject + ': ' + cut(r.text, 90)); }); }
    if (o.lessons) { L.push(''); L.push('📚 *الدروس*'); m.rows.forEach(function (r) { L.push('• ' + r.subject + ': ' + r.shown.map(function (x) { return cut(x.lesson, 70); }).join(' / ')); }); }
    L.push(''); L.push('🔗 التفاصيل الكاملة: ' + link);
    return L.join('\n');
  }

  // ---------- ملف التقويم (.ics) ----------
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function ymd(d) { return d.getUTCFullYear() + pad(d.getUTCMonth() + 1) + pad(d.getUTCDate()); }
  function examDate(startIso, dayTok) {                       // الاثنين الذي يحوي بداية الأسبوع + إزاحة اليوم
    var s = new Date(startIso + 'T00:00:00Z'), dow = (s.getUTCDay() + 6) % 7;                      // الاثنين=0 … الأحد=6
    var mon = new Date(s.getTime() + (dow >= 5 ? 7 - dow : -dow) * 86400000);                       // بداية سبت/أحد ← الاثنين القادم، وغير ذلك ← الاثنين الذي قبلها
    return new Date(mon.getTime() + (DAYIDX[dayTok] || 0) * 86400000);
  }
  function esc(s) { return String(s == null ? '' : s).replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n'); }
  function fold(line) {                                        // طيّ الأسطر عند 75 بايت دون شطر حرف
    var out = [], cur = '', bytes = 0, limit = 75;
    for (var ch of line) { var b = unescape(encodeURIComponent(ch)).length; if (bytes + b > limit) { out.push(cur); cur = ' '; bytes = 1; limit = 75; } cur += ch; bytes += b; }
    out.push(cur); return out.join('\r\n');
  }
  function ics(m, o) {
    o = o || {}; if (!m.examWeek || !m.items.length) return '';
    var now = o.now ? new Date(o.now) : new Date(), stamp = now.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
    var L = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//khotamaereid.net//exams//AR', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH', 'X-WR-CALNAME:' + esc('اختبارات ' + m.classLabel), 'X-WR-TIMEZONE:Asia/Dubai'];
    m.items.forEach(function (i) {
      var d = examDate(m.examWeek.start, i.day), d2 = new Date(d.getTime() + 86400000), c = i.card || {};
      var title = 'اختبار ' + (i.subject || (XS[i.code] || [i.code])[0]) + ' — ' + m.classLabel + tag(m, i);
      var desc = [c.title && ('المقرر: ' + c.title), c.pages && ('الصفحات: ' + c.pages), c.type && ('النوع: ' + c.type), c.marks && ('الدرجة: ' + c.marks), c.sources && ('المصادر: ' + c.sources), 'التفاصيل: https://khotamaereid.net/summary'].filter(Boolean).join('\n');
      L.push('BEGIN:VEVENT', 'UID:ssuae-' + (m.examWeek.num || ymd(d)) + '-' + (i.secs.join('') || m.group) + '-' + i.code + '-' + i.day + '@khotamaereid.net', 'DTSTAMP:' + stamp,
        'DTSTART;VALUE=DATE:' + ymd(d), 'DTEND;VALUE=DATE:' + ymd(d2), 'SUMMARY:' + esc(title), 'DESCRIPTION:' + esc(desc), 'TRANSP:TRANSPARENT', 'CATEGORIES:اختبار',
        'BEGIN:VALARM', 'ACTION:DISPLAY', 'DESCRIPTION:' + esc('غداً: ' + title), 'TRIGGER:-PT4H', 'END:VALARM', 'END:VEVENT');   // تذكير الثامنة مساء اليوم السابق
    });
    L.push('END:VCALENDAR'); return L.map(fold).join('\r\n') + '\r\n';
  }
  window.SiteMerge = { model: model, whatsapp: whatsapp, ics: ics, codeOf: codeOf, _examDate: examDate, _fold: fold };
  if (typeof module !== 'undefined') module.exports = window.SiteMerge;
})();
