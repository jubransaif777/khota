// ============================================================
//  شهادات الشكر والتقدير — يكتبها المدير من لوحة الإدارة، ويجدها وليّ الأمر بالاستعلام برقم الهوية والاسم الأول.
//  • عام (POST {id,name}): يعيد شهادات ذلك الطالب فقط. أي فشل (رقم غير موجود / اسم لا يطابق / لا شهادات) = ردّ «nomatch» واحد.
//    ويشارك حارس ألف والتقرير: 5 محاولات فاشلة لكل رقم هوية في اليوم، و30 لكل عنوان في الساعة، وسجلّ تدقيق بلا أسماء.
//  • الإدارة (POST + pw = ADMIN_PASSWORD): get / issue / update / remove / settings / roster
//  يُحفظ في Netlify Blobs (مخزن site-certs). رقم هوية الطالب لا يُعاد في الردّ العام أبداً.
// ============================================================
const crypto = require('crypto');
const guard = require('./lib/guard');

const J = (code, obj) => ({ statusCode: code, headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }, body: JSON.stringify(obj) });
const MAX_CERTS = 3000, MAX_BATCH = 300;
const dubaiToday = () => new Date().toLocaleDateString('en-CA', { timeZone: 'Asia/Dubai' });
const sha = x => crypto.createHash('sha256').update(String(x)).digest();
const sameSecret = (a, b) => crypto.timingSafeEqual(sha(a), sha(b));
const clean = (s, max) => String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').replace(/\r\n?/g, '\n').replace(/[ \t]+/g, ' ').replace(/ ?\n ?/g, '\n').replace(/\n{3,}/g, '\n\n').trim().slice(0, max || 200);
const latin = s => String(s == null ? '' : s).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d));
const digits = s => latin(s).replace(/\D/g, '');
const isDate = s => /^\d{4}-\d{2}-\d{2}$/.test(s) && new Date(s + 'T00:00:00Z').toISOString().slice(0, 10) === s;
const newId = () => crypto.randomBytes(8).toString('hex');
const okId = s => /^[a-zA-Z0-9-]{6,40}$/.test(String(s || ''));
class Bad extends Error { constructor(path, msg) { super(msg); this.path = path; } }

// تطبيع الاسم العربي: يُسقط التشكيل والتطويل ويوحّد الهمزات والياء والتاء المربوطة ويضمّ المسافات
const compact = s => latin(String(s == null ? '' : s)).toLowerCase().replace(/[ً-ٰٟـ]/g, '').replace(/[أإآٱ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه').replace(/[^\p{L}\p{N}]+/gu, '');
function firstNameOk(fullName, input) {                // الاسم الأول (أو الأولان كـ«عبد الله»)
  const q = compact(input); if (!q) return false;
  const toks = String(fullName || '').trim().split(/\s+/).filter(Boolean);
  return q === compact(toks[0]) || (toks.length > 1 && q === compact(toks[0] + toks[1]));
}

const DEFAULT_SETTINGS = () => ({
  school: 'مدرسة المعيريض للحلقة الثانية بنين',
  ministry: 'وزارة التربية والتعليم — دولة الإمارات العربية المتحدة',
  sigs: [{ title: 'مدير المدرسة', name: '' }, { title: 'وكيل المدرسة للشؤون الأكاديمية', name: 'جبران سيف' }],
});
const emptyDoc = () => ({ rev: 0, updated: 0, settings: DEFAULT_SETTINGS(), certs: [] });
async function readDoc(store) {
  const d = await store.get('certs', { type: 'json' }).catch(() => null);
  const doc = Object.assign(emptyDoc(), d && typeof d === 'object' ? d : {});
  doc.settings = Object.assign(DEFAULT_SETTINGS(), doc.settings || {});
  if (!Array.isArray(doc.certs)) doc.certs = [];
  return doc;
}

function vSettings(x) {
  const o = {};
  o.school = clean(x.school, 100); if (!o.school) throw new Bad('settings.school', 'اسم المدرسة مطلوب');
  o.ministry = clean(x.ministry, 120);
  const sigs = Array.isArray(x.sigs) ? x.sigs : [];
  if (sigs.length > 3) throw new Bad('settings.sigs', 'الحدّ الأقصى ثلاثة توقيعات');
  o.sigs = sigs.map((s, i) => { const title = clean(s && s.title, 60); if (!title) throw new Bad('settings.sigs[' + i + '].title', 'مسمّى الموقّع مطلوب'); return { title, name: clean(s && s.name, 60) }; });
  return o;
}
function vStudent(s, p) {
  const sid = digits(s && s.sid); if (sid.length < 10 || sid.length > 20) throw new Bad(p + '.sid', 'رقم الهوية غير صحيح');
  const name = clean(s && s.name, 80); if (!name) throw new Bad(p + '.name', 'اسم الطالب مطلوب');
  return { sid, name, section: clean(s && s.section, 40) };
}
function vContent(x, p) {
  const type = clean(x.type, 40); if (!type) throw new Bad(p + '.type', 'نوع الشهادة مطلوب');
  const body = clean(x.body, 600); if (!body) throw new Bad(p + '.body', 'نص الشهادة مطلوب');
  const date = x.date ? String(x.date) : dubaiToday(); if (!isDate(date)) throw new Bad(p + '.date', 'التاريخ غير صحيح');
  return { type, body, date };
}

const failedLogin = new Map();
function tooMany(ip) { const now = Date.now(), a = (failedLogin.get(ip) || []).filter(t => now - t < 3600000); failedLogin.set(ip, a); return a.length >= 10; }
function noteFail(ip) { const a = failedLogin.get(ip) || []; a.push(Date.now()); failedLogin.set(ip, a); }
const ipOf = event => { const h = (event && event.headers) || {}; return (h['x-nf-client-connection-ip'] || h['client-ip'] || h['x-forwarded-for'] || '0').split(',')[0].trim(); };
const resolve = (c) => String(c.body).replace(/\{الاسم\}/g, c.name).replace(/\{الصف\}/g, c.section || '');

async function publicLookup(event, q, deps) {
  const ip = ipOf(event), hd = event.headers || {}, ua = /mobile|android|iphone|ipad/i.test(hd['user-agent'] || '') ? 'm' : 'd';
  const idD = digits(q.id), name = clean(q.name, 80);
  if (idD.length < 10 || !name) return J(400, { status: 'invalid' });
  const g = await guard.check(deps.audit, idD, ip);
  const base = { k: 'cert', id: guard.mask(idD), idh: g.idh.slice(0, 6), iph: g.iph.slice(0, 6), m: 'POST', ua };
  if (!g.ok) { await guard.log(deps.audit, Object.assign({ r: g.reason === 'locked' ? 'locked' : 'rate' }, base)); return J(429, { status: g.reason }); }
  const doc = await readDoc(deps.store);
  const mine = doc.certs.filter(c => c.sid === idD);
  const nameOk = mine.some(c => firstNameOk(c.name, name));
  const list = nameOk ? mine.filter(c => !c.hidden) : [];
  if (!list.length) { await guard.fail(deps.audit, g); await guard.log(deps.audit, Object.assign({ r: 'fail' }, base)); return J(404, { status: 'nomatch' }); }
  await guard.log(deps.audit, Object.assign({ r: 'ok' }, base));
  list.sort((a, b) => (b.date || '').localeCompare(a.date || '') || b.created - a.created);
  return J(200, { status: 'ok', student: { name: list[0].name }, school: doc.settings.school, ministry: doc.settings.ministry, sigs: doc.settings.sigs,
    certs: list.map(c => ({ id: c.id, type: c.type, date: c.date, name: c.name, section: c.section, text: resolve(c) })) });
}

async function admin(event, b, deps) {
  const ip = ipOf(event), pw = process.env.ADMIN_PASSWORD;
  if (tooMany(ip)) return J(429, { ok: false, error: 'rate' });
  if (!pw || !b.pw || !sameSecret(b.pw, pw)) { noteFail(ip); return J(401, { ok: false, error: 'auth' }); }
  if (b.action === 'roster') {
    try { const list = deps.roster ? await deps.roster() : null; if (!list) return J(200, { ok: true, roster: [], note: 'nofile' }); return J(200, { ok: true, roster: list }); }
    catch (e) { return J(200, { ok: true, roster: [], note: 'error' }); }
  }
  const doc = await readDoc(deps.store);
  if (b.action === 'get') return J(200, { ok: true, today: dubaiToday(), doc });
  if (!['issue', 'update', 'remove', 'settings'].includes(b.action)) return J(400, { ok: false, error: 'action' });
  if (Number(b.rev) !== doc.rev) return J(409, { ok: false, error: 'conflict', today: dubaiToday(), doc });
  const out = Object.assign({}, doc, { certs: doc.certs.slice() });
  let issued = 0, skipped = 0;
  try {
    if (b.action === 'issue') {
      const list = Array.isArray(b.students) ? b.students : [];
      if (!list.length) throw new Bad('students', 'اختر طالباً واحداً على الأقل');
      if (list.length > MAX_BATCH) throw new Bad('students', 'الحدّ الأقصى ' + MAX_BATCH + ' طالباً في الدفعة');
      if (doc.certs.length + list.length > MAX_CERTS) throw new Bad('students', 'تجاوزتَ الحدّ الأقصى للشهادات (' + MAX_CERTS + ')');
      const content = vContent(b, 'cert');
      const seen = new Set(doc.certs.map(c => [c.sid, c.type, c.date, c.body].join('|')));
      list.forEach((s, i) => {
        const st = vStudent(s, 'students[' + i + ']'), key = [st.sid, content.type, content.date, content.body].join('|');
        if (seen.has(key)) { skipped++; return; } seen.add(key);
        out.certs.push(Object.assign({ id: newId(), hidden: false, created: Date.now() + i }, st, content)); issued++;
      });
    } else if (b.action === 'update') {
      const i = out.certs.findIndex(c => c.id === b.id); if (i < 0) throw new Bad('id', 'الشهادة غير موجودة');
      const cur = out.certs[i], f = b.fields || {};
      const st = vStudent({ sid: f.sid != null ? f.sid : cur.sid, name: f.name != null ? f.name : cur.name, section: f.section != null ? f.section : cur.section }, 'fields');
      const content = vContent({ type: f.type != null ? f.type : cur.type, body: f.body != null ? f.body : cur.body, date: f.date != null ? f.date : cur.date }, 'fields');
      out.certs[i] = Object.assign({}, cur, st, content, { hidden: f.hidden != null ? !!f.hidden : !!cur.hidden });
    } else if (b.action === 'remove') {
      const ids = new Set(Array.isArray(b.ids) ? b.ids : [b.id]);
      const before = out.certs.length; out.certs = out.certs.filter(c => !ids.has(c.id));
      if (out.certs.length === before) throw new Bad('id', 'الشهادة غير موجودة');
    } else if (b.action === 'settings') out.settings = vSettings(b.settings || {});
  } catch (e) { if (e instanceof Bad) return J(400, { ok: false, error: 'invalid', path: e.path, message: e.message }); throw e; }
  out.rev = doc.rev + 1; out.updated = Date.now();
  await deps.store.setJSON('certs', out);
  return J(200, { ok: true, today: dubaiToday(), doc: out, issued, skipped });
}

async function handle(event, deps) {
  const method = event.httpMethod || 'GET';
  if (method === 'GET') {                                               // للمراقبة فقط: عدد الشهادات (بلا أي بيانات)
    if (!(event.queryStringParameters || {}).ping) return J(400, { status: 'invalid' });
    return J(200, { status: 'ok', count: (await readDoc(deps.store)).certs.length });
  }
  if (method !== 'POST') return J(405, { status: 'method' });
  let b = {}; try { b = JSON.parse(event.body || '{}'); } catch (e) { return J(400, { status: 'invalid' }); }
  try { return b.action ? await admin(event, b, deps) : await publicLookup(event, b, deps); }
  catch (e) { return J(500, { status: 'error', ok: false, error: 'server' }); }
}

exports.handler = async (event) => {
  let store = null;
  try { const bl = require('@netlify/blobs'); try { bl.connectLambda(event); } catch (e) {} store = bl.getStore('site-certs'); } catch (e) { return J(500, { status: 'error', ok: false, error: 'storage' }); }
  const roster = async () => { const rp = require('./report'); const buf = await rp._download(); return buf ? rp._roster(buf) : null; };
  return handle(event, { store, audit: guard.open(event), roster });
};
exports._handle = handle; exports._firstNameOk = firstNameOk; exports._compact = compact;
