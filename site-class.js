/* شعبة ولي الأمر الموحّدة: يختارها مرة واحدة فتسري على الخطة والملخص والجدول والاختبارات.
   تُحفظ على جهازه فقط. الخطة والملخص يستعملان «المجموعة» (6G)، والجدول والاختبارات «الشعبة» (6G1). */
(function () {
  var K = 'ssuae_myclass';
  function grpOf(s) { var m = String(s || '').match(/^(\d{1,2}[A-Za-z])/); return m ? m[1].toUpperCase() : ''; }
  function rd() {
    try { var o = JSON.parse(localStorage.getItem(K) || 'null'); if (o && typeof o === 'object') return o; } catch (e) {}
    try { var old = JSON.parse(localStorage.getItem('ssuae_sched') || '{}'); if (old && old.sel) return { grp: grpOf(old.sel), sec: old.sel }; } catch (e) {}   // ترحيل اختيار الجدول القديم
    return {};
  }
  function wr(o) { try { localStorage.setItem(K, JSON.stringify(o)); } catch (e) {} }
  window.SiteClass = {
    group: function (def) { return rd().grp || def || ''; },                                   // 6G
    section: function (def) { var o = rd(); return o.sec || (o.grp ? o.grp + '1' : (def || '')); },  // 6G1 (إن عُرفت المجموعة فقط: الشعبة 1)
    sub: function () { return rd().sec || ''; },                                               // الشعبة الدقيقة المختارة صراحةً أو ''
    setGroup: function (g) { var o = rd(); if (o.grp !== g) { o.grp = g; o.sec = ''; } wr(o); }, // تغيير المجموعة يمسح الشعبة الدقيقة
    setSection: function (s) { wr({ grp: grpOf(s), sec: s }); },
    clearSub: function () { var o = rd(); o.sec = ''; wr(o); }                                   // «كل الشعب» في المجموعة
  };
})();
