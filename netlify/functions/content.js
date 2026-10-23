// ============================================================
//  مركز المحتوى — إعلانات/أخبار/روابط تُكتب من لوحة الإدارة وتظهر في صفحات الأولياء دون المرور بالإكسل.
//  • عام (GET):  يعيد المحتوى النشط فقط.
//  • الإدارة (POST + pw = ADMIN_PASSWORD): {action:'get'} أو {action:'save', rev, announcements?, news?, links?}
//  يُحفظ في Netlify Blobs (مخزن site-content). يلزم "@netlify/blobs" في package.json (مضاف سابقاً للتقييمات).
// ============================================================
const crypto = require('crypto');
const J = (code, obj, cache) => ({ statusCode: code, headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': cache || 'no-store' }, body: JSON.stringify(obj) });
const LIM = { announcements: 30, news: 200, links: 100 };
const LEVELS = ['info', 'important', 'urgent'];
const dubaiToday = () => new Date().toLocaleDateString('en-CA', { timeZone: 'Asia/Dubai' });   // YYYY-MM-DD بتوقيت الإمارات
const sha = x => crypto.createHash('sha256').update(String(x)).digest();
const sameSecret = (a, b) => crypto.timingSafeEqual(sha(a), sha(b));
const clean = (s, max) => String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').replace(/\r\n?/g, '\n').replace(/[ \t]+/g, ' ').replace(/ ?\n ?/g, '\n').replace(/\n{3,}/g, '\n\n').trim().slice(0, max);
const isDate = s => /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(Date.parse(s + 'T00:00:00Z')) && new Date(s + 'T00:00:00Z').toISOString().slice(0, 10) === s;
const newId = () => crypto.randomBytes(8).toString('hex');
const okId = s => /^[a-zA-Z0-9-]{6,40}$/.test(String(s || ''));
class Bad extends Error { constructor(path, msg) { super(msg); this.path = path; } }

function vAnn(x, p, today) {
  const text = clean(x.text, 600); if (!text) throw new Bad(p + '.text', 'نص الإعلان مطلوب');
  const level = LEVELS.includes(x.level) ? x.level : 'info';
  const start = x.start ? String(x.start) : today, end = x.end ? String(x.end) : '';
  if (!isDate(start)) throw new Bad(p + '.start', 'تاريخ البداية غير صحيح');
  if (end && !isDate(end)) throw new Bad(p + '.end', 'تاريخ النهاية غير صحيح');
  if (end && end < start) throw new Bad(p + '.end', 'تاريخ النهاية قبل البداية');
  return { id: okId(x.id) ? x.id : newId(), title: clean(x.title, 80), text, level, start, end, popup: !!x.popup, created: Number(x.created) || Date.now() };
}
function vNews(x, p, today) {
  const title = clean(x.title, 120); if (!title) throw new Bad(p + '.title', 'عنوان الخبر مطلوب');
  const date = x.date ? String(x.date) : today; if (!isDate(date)) throw new Bad(p + '.date', 'التاريخ غير صحيح');
  return { id: okId(x.id) ? x.id : newId(), title, body: clean(x.body, 1500), date, pinned: !!x.pinned };
}
function vLink(x, p) {
  const title = clean(x.title, 80); if (!title) throw new Bad(p + '.title', 'عنوان الرابط مطلوب');
  let url = clean(x.url, 500).replace(/\s/g, ''); if (!url) throw new Bad(p + '.url', 'الرابط مطلوب');
  if (!/^https?:\/\//i.test(url)) url = 'https://' + url;
  let u; try { u = new URL(url); } catch (e) { throw new Bad(p + '.url', 'الرابط غير صحيح'); }
  if (!/^https?:$/.test(u.protocol) || !u.hostname.includes('.')) throw new Bad(p + '.url', 'الرابط غير صحيح');
  return { id: okId(x.id) ? x.id : newId(), title, url: u.href, category: clean(x.category, 40), desc: clean(x.desc, 120) };
}
const VALID = { announcements: vAnn, news: vNews, links: vLink };

const emptyDoc = () => ({ rev: 0, updated: 0, announcements: [], news: [], links: [] });
async function readDoc(store) { const d = await store.get('content', { type: 'json' }).catch(() => null); return d && typeof d === 'object' ? Object.assign(emptyDoc(), d) : emptyDoc(); }

const failed = new Map();      // محاولات دخول فاشلة: 10 في الساعة لكل عنوان
function tooMany(ip) { const now = Date.now(); const a = (failed.get(ip) || []).filter(t => now - t < 3600000); failed.set(ip, a); return a.length >= 10; }
function noteFail(ip) { const a = failed.get(ip) || []; a.push(Date.now()); failed.set(ip, a); }

function publicView(doc, today) {
  const rank = { urgent: 0, important: 1, info: 2 };
  const announcements = doc.announcements.filter(a => a.start <= today && (!a.end || a.end >= today))
    .sort((a, b) => rank[a.level] - rank[b.level] || b.created - a.created).map(a => ({ id: a.id, title: a.title, text: a.text, level: a.level, popup: a.popup }));
  const news = doc.news.slice().sort((a, b) => (a.pinned !== b.pinned ? (a.pinned ? -1 : 1) : (b.date || '').localeCompare(a.date || ''))).map(n => ({ id: n.id, date: n.date, title: n.title, body: n.body, pinned: n.pinned }));
  const links = doc.links.map(l => ({ id: l.id, title: l.title, url: l.url, category: l.category, desc: l.desc }));
  return { ok: true, today, rev: doc.rev, announcements, news, links };
}

async function handle(event, store, opt) {
  const today = (opt && opt.today) || dubaiToday(); const h = event.headers || {};
  const ip = h['x-nf-client-connection-ip'] || (h['x-forwarded-for'] || '').split(',')[0].trim() || 'unknown';
  if (event.httpMethod === 'GET') return J(200, publicView(await readDoc(store), today), 'public, max-age=20');
  if (event.httpMethod !== 'POST') return J(405, { ok: false, error: 'method' });
  let b = {}; try { b = JSON.parse(event.body || '{}'); } catch (e) { return J(400, { ok: false, error: 'json' }); }
  const pw = process.env.ADMIN_PASSWORD;
  if (tooMany(ip)) return J(429, { ok: false, error: 'rate' });
  if (!pw || !b.pw || !sameSecret(b.pw, pw)) { noteFail(ip); return J(401, { ok: false, error: 'auth' }); }
  const doc = await readDoc(store);
  if (b.action === 'get') return J(200, { ok: true, today, doc });
  if (b.action === 'save') {
    if (Number(b.rev) !== doc.rev) return J(409, { ok: false, error: 'conflict', today, doc });       // غُيِّر المحتوى من جهاز آخر
    const out = Object.assign({}, doc);
    try {
      for (const k of Object.keys(LIM)) {
        if (b[k] === undefined) continue;
        if (!Array.isArray(b[k])) throw new Bad(k, 'قائمة غير صالحة');
        if (b[k].length > LIM[k]) throw new Bad(k, 'تجاوزتَ الحدّ الأقصى (' + LIM[k] + ')');
        const ids = new Set();
        out[k] = b[k].map((x, i) => { const v = VALID[k](x || {}, k + '[' + i + ']', today); if (ids.has(v.id)) v.id = newId(); ids.add(v.id); return v; });
      }
    } catch (e) { if (e instanceof Bad) return J(400, { ok: false, error: 'invalid', path: e.path, message: e.message }); throw e; }
    out.rev = doc.rev + 1; out.updated = Date.now();
    await store.setJSON('content', out);
    return J(200, { ok: true, today, doc: out });
  }
  return J(400, { ok: false, error: 'action' });
}

exports.handler = async (event) => {
  let store;
  try { const bl = require('@netlify/blobs'); try { bl.connectLambda(event); } catch (e) {} store = bl.getStore('site-content'); }
  catch (e) { return event.httpMethod === 'GET' ? J(200, { ok: true, today: dubaiToday(), rev: 0, announcements: [], news: [], links: [], degraded: true }, 'public, max-age=20') : J(500, { ok: false, error: 'storage' }); }
  try { return await handle(event, store); }
  catch (e) { return event.httpMethod === 'GET' ? J(200, { ok: true, today: dubaiToday(), rev: 0, announcements: [], news: [], links: [], degraded: true }, 'public, max-age=20') : J(500, { ok: false, error: 'storage' }); }
};
exports._handle = handle;
