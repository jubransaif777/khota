// يحدّث النسخة الاحتياطية المضمّنة في index.html (تظهر فقط لو تعذّر جلب الخطة الحيّة) من ملف إكسل الفصل الجديد.
// الاستعمال:  node tools/refresh-fallback.js "مسار-ملف-الخطة.xlsx"
const fs = require('fs'), path = require('path');
const plan = require('../netlify/functions/plan.js');
const file = process.argv[2]; if (!file) { console.error('الاستعمال: node tools/refresh-fallback.js ملف.xlsx'); process.exit(1); }
const data = plan._pivot(fs.readFileSync(file));
const p = path.join(__dirname, '..', 'index.html'); let html = fs.readFileSync(p, 'utf8');
const re = /window\.__FALLBACK__ = \{.*\};/;
if (!re.test(html)) { console.error('لم أجد __FALLBACK__ في index.html'); process.exit(1); }
html = html.replace(re, () => 'window.__FALLBACK__ = ' + JSON.stringify(data) + ';');
fs.writeFileSync(p, html); console.log('تم: ' + data.term + ' — ' + data.weeks.length + ' أسبوعاً');
