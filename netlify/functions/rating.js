// ============================================================
//  تقييم الموقع — يحفظ تقييمات الزوّار (نجوم 1–5 + ملاحظة اختيارية) في Netlify Blobs
//  • الزائر: POST {id, stars, comment?, page?}   (تقييم واحد لكل جهاز، وإعادة التقييم تعدّله)
//  • الإدارة: POST {action:'list', pw}            (ADMIN_PASSWORD نفسها)
//  لا تُخزَّن أسماء ولا عناوين IP. يلزم إضافة "@netlify/blobs" إلى package.json.
// ============================================================
const crypto = require('crypto');
const PAGES = ['/', '/schedule', '/exams', '/news', '/links'];
const J = (code, obj) => ({ statusCode: code, headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }, body: JSON.stringify(obj) });

// حدّ بسيط للإرسال: 8 مرات في الساعة لكل عنوان (في ذاكرة الدالة)
const hits = new Map();
function limited(ip) {
  const now = Date.now(); const a = (hits.get(ip) || []).filter(t => now - t < 3600000);
  if (a.length >= 8) { hits.set(ip, a); return true; }
  a.push(now); hits.set(ip, a); return false;
}
const sha = x => crypto.createHash('sha256').update(String(x)).digest();
const sameSecret = (a, b) => crypto.timingSafeEqual(sha(a), sha(b));
const cleanText = s => String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();

async function listKeys(store) {
  const keys = [];
  try { for await (const page of store.list({ paginate: true })) for (const b of page.blobs) keys.push(b.key); }
  catch (e) { keys.length = 0; const r = await store.list(); (r.blobs || []).forEach(b => keys.push(b.key)); }
  return keys;
}

async function handle(event, store) {
  if (event.httpMethod !== 'POST') return J(405, { ok: false, error: 'method' });
  let b = {}; try { b = JSON.parse(event.body || '{}'); } catch (e) { return J(400, { ok: false, error: 'json' }); }

  // ---------- لوحة الإدارة ----------
  if (b.action === 'list') {
    const pw = process.env.ADMIN_PASSWORD;
    if (!pw || !b.pw || !sameSecret(b.pw, pw)) return J(401, { ok: false, error: 'auth' });
    const keys = (await listKeys(store)).slice(0, 3000), items = [];
    for (let i = 0; i < keys.length; i += 25) {
      const vals = await Promise.all(keys.slice(i, i + 25).map(k => store.get(k, { type: 'json' }).catch(() => null)));
      vals.forEach(v => { if (v && v.stars >= 1 && v.stars <= 5) items.push(v); });
    }
    const count = items.length, sum = items.reduce((s, v) => s + v.stars, 0);
    const dist = [0, 0, 0, 0, 0]; items.forEach(v => dist[v.stars - 1]++);
    const byPage = {}; items.forEach(v => { const p = byPage[v.page] || (byPage[v.page] = { count: 0, sum: 0 }); p.count++; p.sum += v.stars; });
    Object.keys(byPage).forEach(k => { byPage[k] = { count: byPage[k].count, avg: +(byPage[k].sum / byPage[k].count).toFixed(2) }; });
    const week = Date.now() - 7 * 86400000;
    const recent = items.slice().sort((x, y) => y.ts - x.ts).slice(0, 150).map(v => ({ stars: v.stars, comment: v.comment || '', page: v.page, dev: v.dev, ts: v.ts }));
    return J(200, { ok: true, count, avg: count ? +(sum / count).toFixed(2) : 0, dist, byPage,
      withComment: items.filter(v => v.comment).length, last7: items.filter(v => v.ts >= week).length, recent });
  }

  // ---------- تقييم زائر ----------
  if (b.website) return J(200, { ok: true });                         // فخّ للروبوتات: حقل مخفي لا يملؤه البشر
  const id = String(b.id || ''); if (!/^[a-zA-Z0-9-]{16,64}$/.test(id)) return J(400, { ok: false, error: 'id' });
  const stars = parseInt(b.stars, 10); if (!(stars >= 1 && stars <= 5)) return J(400, { ok: false, error: 'stars' });
  const h = event.headers || {};
  const ip = h['x-nf-client-connection-ip'] || (h['x-forwarded-for'] || '').split(',')[0].trim() || 'unknown';
  if (limited(ip)) return J(429, { ok: false, error: 'rate' });
  const comment = cleanText(b.comment).slice(0, 300);
  const page = PAGES.includes(b.page) ? b.page : 'other';
  const dev = /mobile|android|iphone|ipad/i.test(h['user-agent'] || '') ? 'm' : 'd';
  const prev = await store.get(id, { type: 'json' }).catch(() => null);
  await store.setJSON(id, { stars, comment, page, dev, ts: Date.now(), first: (prev && prev.first) || Date.now(), edits: prev ? (prev.edits || 0) + 1 : 0 });
  return J(200, { ok: true });
}

exports.handler = async (event) => {
  let store;
  try { const bl = require('@netlify/blobs'); try { bl.connectLambda(event); } catch (e) {} store = bl.getStore('site-ratings'); }
  catch (e) { return J(500, { ok: false, error: 'storage' }); }
  try { return await handle(event, store); } catch (e) { return J(500, { ok: false, error: 'storage' }); }
};
exports._handle = handle;
