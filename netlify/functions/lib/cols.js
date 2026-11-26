// v3.4 — قراءة أعمدة ورقة الأسبوع من عناوينها (لا من أرقام ثابتة).
// يعمل مع الملف القديم (٨ أعمدة بالمجموعة: 5G…8A) ومع الأعمدة الجديدة بالشعبة (5G1, 5G2…).
//   عنوان العمود يحمل رمز الشعبة، مثل: «الخامس العام ١\n5G1». أضف عموداً أو احذفه فيتبعه الموقع.
const XLSX = require('xlsx');
const LEGACY = { groups: ['5G','5A','6G','6A','7G','7A','8G','8A'], col0: 3, note: 11, exam: 12, rem: 13 };
const westify = s => String(s == null ? '' : s).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d));
const CODE = /(?:^|[^A-Za-z0-9])([5-9])\s*([GA])\s*(\d{1,2})?(?![A-Za-z0-9])/i;

// يعيد { headerRow, secCols:[{col,group,sec}], note, exam, rem, legacy:boolean }
function detect(ws, cell) {
  const ref = ws && ws['!ref'];
  const maxC = ref ? XLSX.utils.decode_range(ref).e.c + 1 : 0;
  for (let r = 1; r <= 8; r++) {
    const cols = []; let note = 0, exam = 0, rem = 0, hw = 0;
    for (let c = 3; c <= Math.max(maxC, 3); c++) {
      const t = westify(cell(ws, r, c));
      if (!t) continue;
      const m = t.match(CODE);
      if (m) { const group = m[1] + m[2].toUpperCase(); cols.push({ col: c, group, sec: m[3] ? group + m[3] : '' }); continue; }
      if (/remind|تذكير/i.test(t)) rem = c;
      else if (/homework|واجب|مهام|مهمة|مهمّة/i.test(t)) hw = c;
      else if (/exam|امتحان|اختبار|موعد/i.test(t)) exam = c;
      else if (/note|ملاحظ/i.test(t)) note = c;
    }
    if (cols.length >= 2) {
      // مجموعة بعمود واحد = بلا تمييز شعب
      const per = {}; cols.forEach(x => per[x.group] = (per[x.group] || 0) + 1);
      cols.forEach(x => { if (per[x.group] < 2) x.sec = ''; });
      return { headerRow: r, secCols: cols, note, exam, rem, hw, legacy: false }   // بلا عنوان = العمود غير موجود (0);
    }
  }
  return { headerRow: 4, secCols: LEGACY.groups.map((g, i) => ({ col: LEGACY.col0 + i, group: g, sec: '' })),
           note: LEGACY.note, exam: LEGACY.exam, rem: LEGACY.rem, hw: 0, legacy: true };
}
// الأعمدة الخاصة بمجموعة
const colsOf = (d, group) => d.secCols.filter(x => x.group === group);
module.exports = { detect, colsOf, LEGACY, CODE };
