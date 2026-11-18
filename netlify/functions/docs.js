// ============================================================
//  المستندات (PDF): جداول الاختبارات والأدلة والتعاميم — يرفعها المدير من لوحة الإدارة ويقرؤها الأولياء في /docs
//  • عام (GET): قائمة المستندات الظاهرة، أو ?id=… لفتح الملف (&dl=1 للتنزيل). المخفي والمنتهي لا يُقدَّم.
//  • الإدارة (POST + pw = ADMIN_PASSWORD): get / upload / update / remove
//  • التخزين في Netlify Blobs (المخزن site-docs). الحدّ ٤ ميغابايت للملف (حدّ الدوال ٦ ميغابايت بعد الترميز).
//  ملف مستقل: لا يحتاج أي ملف آخر في مجلد lib.
// ============================================================
const crypto = require('crypto');

const J = (code, obj) => ({ statusCode: code, headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }, body: JSON.stringify(obj) });
const MAX_RAW = 4 * 1024 * 1024, MAX_DOCS = 200, NEW_DAYS = 7;
const dubaiToday = () => new Date().toLocaleDateString('en-CA', { timeZone: 'Asia/Dubai' });
const sha = x => crypto.createHash('sha256').update(String(x)).digest();
const sameSecret = (a, b) => crypto.timingSafeEqual(sha(a), sha(b));
const clean = (s, max) => String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').replace(/\r\n?/g, '\n').replace(/[ \t]+/g, ' ').replace(/ ?\n ?/g, '\n').replace(/\n{3,}/g, '\n\n').trim().slice(0, max || 200);
const isDate = s => { try { return /^\d{4}-\d{2}-\d{2}$/.test(s) && new Date(s + 'T00:00:00Z').toISOString().slice(0, 10) === s; } catch (e) { return false; } };
const newId = () => crypto.randomBytes(6).toString('hex');
const okId = s => /^[a-f0-9]{12}$/.test(String(s || ''));
class Bad extends Error { constructor(path, msg) { super(msg); this.path = path; } }

const emptyDoc = () => ({ rev: 0, updated: 0, docs: [] });
async function readIndex(store) {
  const d = await store.get('index', { type: 'json' }).catch(() => null);
  const doc = Object.assign(emptyDoc(), d && typeof d === 'object' ? d : {});
  if (!Array.isArray(doc.docs)) doc.docs = [];
  return doc;
}
const visible = (d, today) => !d.hidden && (!d.expires || d.expires >= today);

function vMeta(x, p, cur) {
  cur = cur || {};
  const pick = k => (x[k] !== undefined ? x[k] : cur[k]);
  const title = clean(pick('title'), 120); if (title.length < 3) throw new Bad(p + '.title', 'اكتب عنواناً للمستند (٣ أحرف على الأقل)');
  const cat = clean(pick('cat'), 30); if (!cat) throw new Bad(p + '.cat', 'اختر تصنيفاً');
  const desc = clean(pick('desc'), 300);
  const date = pick('date') ? String(pick('date')) : dubaiToday(); if (!isDate(date)) throw new Bad(p + '.date', 'التاريخ غير صحيح');
  const expires = pick('expires') ? String(pick('expires')) : ''; if (expires && (!isDate(expires) || expires < date)) throw new Bad(p + '.expires', 'تاريخ الانتهاء غير صحيح أو قبل تاريخ النشر');
  return { title, cat, desc, date, expires, pinned: !!pick('pinned'), hidden: !!pick('hidden') };
}
function decodePdf(b64) {
  const s = String(b64 || '').replace(/^data:[^,]*,/, '').replace(/\s+/g, '');
  if (!s || !/^[A-Za-z0-9+/]+={0,2}$/.test(s)) throw new Bad('data', 'الملف غير صالح');
  if (s.length > Math.ceil(MAX_RAW * 4 / 3) + 8) throw new Bad('data', 'حجم الملف يتجاوز ٤ ميغابايت. اضغطه ثم أعد المحاولة.');
  const buf = Buffer.from(s, 'base64');
  if (buf.length > MAX_RAW) throw new Bad('data', 'حجم الملف يتجاوز ٤ ميغابايت. اضغطه ثم أعد المحاولة.');
  if (buf.length < 100 || buf.slice(0, 5).toString('latin1') !== '%PDF-') throw new Bad('data', 'الملف ليس PDF صحيحاً');
  return buf;
}
const safeName = t => clean(t, 80).replace(/[\\/:*?"<>|]+/g, ' ').trim() || 'document';

const failedLogin = new Map();
function tooMany(ip) { const now = Date.now(), a = (failedLogin.get(ip) || []).filter(t => now - t < 3600000); failedLogin.set(ip, a); return a.length >= 10; }
function noteFail(ip) { const a = failedLogin.get(ip) || []; a.push(Date.now()); failedLogin.set(ip, a); }
const ipOf = event => { const h = (event && event.headers) || {}; return (h['x-nf-client-connection-ip'] || h['client-ip'] || h['x-forwarded-for'] || '0').split(',')[0].trim(); };

async function bump(store, id) {                       // عدّاد فتح تقريبي (لا يؤثّر على فتح الملف إن فشل)
  try { const n = parseInt(await store.get('v:' + id, { type: 'text' }), 10) || 0; await store.set('v:' + id, String(n + 1)); } catch (e) {}
}

async function publicGet(event, deps) {
  const q = Object.assign({}, event.queryStringParameters || {}), today = dubaiToday();
  if (!q.id) {                                                          // الرابط النظيف: المعرّف من المسار نفسه (/files/<id>.pdf أو /.netlify/functions/docs/<id>.pdf)
    for (const u of [event.path, event.rawUrl, event.rawPath]) {
      const m = /\/(?:files|docs)\/([A-Za-z0-9]{6,40})(?:\.pdf)?\/?(?:[?#].*)?$/i.exec(String(u || ''));
      if (m) { q.id = m[1]; break; }
    }
  }
  if (q.id) q.id = String(q.id).replace(/\.pdf$/i, '');                    // الرابط النظيف: /files/<id>.pdf
  if (q.ping) return J(200, { status: 'ok', count: (await readIndex(deps.store)).docs.length });
  const idx = await readIndex(deps.store);
  if (!q.id) {
    const now = Date.now(), list = idx.docs.filter(d => visible(d, today)).map(d => ({ id: d.id, title: d.title, cat: d.cat, desc: d.desc, date: d.date, size: d.size, pinned: !!d.pinned, isNew: now - d.created < NEW_DAYS * 86400000 }));
    list.sort((a, b) => (b.pinned - a.pinned) || b.date.localeCompare(a.date));
    return J(200, { status: 'ok', docs: list });
  }
  if (!okId(q.id)) return J(404, { status: 'nomatch' });
  const d = idx.docs.find(x => x.id === q.id);
  if (!d || !visible(d, today)) return J(404, { status: 'nomatch' });
  const ab = await deps.store.get('f:' + d.id, { type: 'arrayBuffer' }).catch(() => null);
  if (!ab) return J(404, { status: 'nomatch' });
  const buf = Buffer.from(ab);
  await bump(deps.store, d.id);
  const fname = safeName(d.title) + '.pdf';
  return { statusCode: 200, isBase64Encoded: true, body: buf.toString('base64'), headers: {
    'Content-Type': 'application/pdf', 'Content-Disposition': (q.dl ? 'attachment' : 'inline') + '; filename="document.pdf"; filename*=UTF-8\'\'' + encodeURIComponent(fname),
    'Cache-Control': 'public, max-age=60', 'X-Content-Type-Options': 'nosniff' } };
}

async function admin(event, b, deps) {
  const ip = ipOf(event), pw = process.env.ADMIN_PASSWORD;
  if (tooMany(ip)) return J(429, { ok: false, error: 'rate' });
  if (!pw || !b.pw || !sameSecret(b.pw, pw)) { noteFail(ip); return J(401, { ok: false, error: 'auth' }); }
  const idx = await readIndex(deps.store);
  if (b.action === 'get') {
    const views = {}; for (const d of idx.docs) views[d.id] = parseInt(await deps.store.get('v:' + d.id, { type: 'text' }).catch(() => 0), 10) || 0;
    return J(200, { ok: true, today: dubaiToday(), doc: idx, views });
  }
  if (!['upload', 'update', 'remove'].includes(b.action)) return J(400, { ok: false, error: 'action' });
  if (Number(b.rev) !== idx.rev) return J(409, { ok: false, error: 'conflict', today: dubaiToday(), doc: idx });
  const out = Object.assign({}, idx, { docs: idx.docs.slice() }); let id = null;
  try {
    if (b.action === 'upload') {
      if (idx.docs.length >= MAX_DOCS) throw new Bad('data', 'بلغتَ الحدّ الأقصى (' + MAX_DOCS + ' مستنداً). احذف القديم أولاً.');
      const buf = decodePdf(b.data), m = vMeta(b, 'doc');
      id = newId(); await deps.store.set('f:' + id, buf);
      const now = Date.now(); out.docs.push(Object.assign({ id, size: buf.length, fname: clean(b.filename, 120), created: now, updated: now }, m));
    } else if (b.action === 'update') {
      const i = out.docs.findIndex(d => d.id === b.id); if (i < 0) throw new Bad('id', 'المستند غير موجود');
      const cur = out.docs[i], m = vMeta(b.fields || {}, 'fields', cur); let extra = {};
      if (b.data) { const buf = decodePdf(b.data); await deps.store.set('f:' + cur.id, buf); extra = { size: buf.length, fname: clean(b.filename, 120), updated: Date.now() }; }
      out.docs[i] = Object.assign({}, cur, m, extra); id = cur.id;
    } else if (b.action === 'remove') {
      const i = out.docs.findIndex(d => d.id === b.id); if (i < 0) throw new Bad('id', 'المستند غير موجود');
      id = out.docs[i].id; out.docs.splice(i, 1);
      await deps.store.delete('f:' + id).catch(() => {}); await deps.store.delete('v:' + id).catch(() => {});
    }
  } catch (e) { if (e instanceof Bad) return J(400, { ok: false, error: 'invalid', path: e.path, message: e.message }); throw e; }
  out.rev = idx.rev + 1; out.updated = Date.now();
  await deps.store.setJSON('index', out);
  return J(200, { ok: true, today: dubaiToday(), doc: out, id });
}

async function handle(event, deps) {
  const method = event.httpMethod || 'GET';
  try {
    if (method === 'GET' || method === 'HEAD') return await publicGet(event, deps);
    if (method !== 'POST') return J(405, { status: 'method' });
    let b = {}; try { b = JSON.parse(event.body || '{}'); } catch (e) { return J(400, { ok: false, error: 'invalid' }); }
    return await admin(event, b, deps);
  } catch (e) { return J(500, { status: 'error', ok: false, error: 'server' }); }
}

exports.handler = async (event) => {
  let store = null;
  try { const bl = require('@netlify/blobs'); try { bl.connectLambda(event); } catch (e) {} store = bl.getStore('site-docs'); } catch (e) { return J(500, { status: 'error', ok: false, error: 'storage' }); }
  return handle(event, { store });
};
exports._handle = handle;
