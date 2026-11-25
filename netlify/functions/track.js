// ============================================================
//  إحصاءات الاستخدام المجهولة + قراءة سجل الأمان للمدير.
//  • عام (POST بلا كلمة سرّ): نبضة واحدة لكل زيارة: {t:'hit'|'ev', p, f, n, s, g, e}. لا تُخزَّن هوية ولا عنوان IP.
//  • المدير (POST + pw): {action:'stats', days} أو {action:'audit'}.
//  الأعداد تقريبية (±): الكتابة بلا ذرّية، لذا تُوزَّع على 8 أجزاء عشوائية.
// ============================================================
const crypto = require('crypto');
const guard = require('./lib/guard');
const J = (code, obj) => ({ statusCode: code, headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }, body: JSON.stringify(obj) });
const PAGES = ['/', '/summary', '/schedule', '/exams', '/news', '/links', '/alef', '/report', '/certs', '/docs'];
const EVENTS = ['share', 'ics', 'print', 'install', 'install_prompt', 'whatsapp', 'cert_view', 'cert_print', 'doc_open', 'doc_download'];
const SHARDS = 8;
let NOW = () => Date.now();
const dayOf = t => new Date(t).toLocaleDateString('en-CA', { timeZone: 'Asia/Dubai' });
const hourOf = t => +new Date(t).toLocaleString('en-GB', { timeZone: 'Asia/Dubai', hour: '2-digit', hour12: false });
const sha = x => crypto.createHash('sha256').update(String(x)).digest();
const sameSecret = (a, b) => crypto.timingSafeEqual(sha(a), sha(b));
const isBot = ua => /bot|crawl|spider|preview|facebookexternalhit|whatsapp|telegram|curl|wget|python|node-fetch|headless/i.test(ua || '');
const isMobile = ua => /mobile|android|iphone|ipad/i.test(ua || '');
const empty = () => ({ t: 0, v: {}, u: 0, nw: 0, m: 0, s: 0, h: new Array(24).fill(0), c: {}, e: {} });

const RATE = new Map();     // ip -> {n, reset}   200 نبضة/ساعة لكل عنوان (يمنع تضخيم الأرقام)
function limited(ip) { const now = NOW(), r = RATE.get(ip); if (!r || now > r.reset) { RATE.set(ip, { n: 1, reset: now + 3600000 }); return false; } r.n++; return r.n > 200; }
const ipOf = e => { const h = (e && e.headers) || {}; return (h['x-nf-client-connection-ip'] || h['client-ip'] || (h['x-forwarded-for'] || '0').split(',')[0]).trim(); };

async function bump(st, d) {
  const key = 'd:' + dayOf(NOW()) + ':' + crypto.randomInt(SHARDS);
  let doc = null; try { doc = await st.get(key, { type: 'json' }); } catch (e) {}
  doc = Object.assign(empty(), doc || {});
  doc.t += d.view ? 1 : 0;
  if (d.view) { doc.v[d.page] = (doc.v[d.page] || 0) + 1; doc.h[d.hour] = (doc.h[d.hour] || 0) + 1; if (d.mobile) doc.m++; if (d.standalone) doc.s++; }
  if (d.uniq) { doc.u++; if (d.group) doc.c[d.group] = (doc.c[d.group] || 0) + 1; }
  if (d.newDev) doc.nw++;
  if (d.event) doc.e[d.event] = (doc.e[d.event] || 0) + 1;
  await st.setJSON(key, doc);
}
function merge(into, doc) {
  into.t += doc.t || 0; into.u += doc.u || 0; into.nw += doc.nw || 0; into.m += doc.m || 0; into.s += doc.s || 0;
  for (const k in doc.v || {}) into.v[k] = (into.v[k] || 0) + doc.v[k];
  for (let i = 0; i < 24; i++) into.h[i] += (doc.h && doc.h[i]) || 0;
  for (const k in doc.c || {}) into.c[k] = (into.c[k] || 0) + doc.c[k];
  for (const k in doc.e || {}) into.e[k] = (into.e[k] || 0) + doc.e[k];
}
async function stats(st, days) {
  days = Math.max(1, Math.min(90, +days || 14));
  const out = []; const today = NOW();
  const dates = []; for (let i = days - 1; i >= 0; i--) dates.push(dayOf(today - i * 86400000));
  const keys = []; for (const d of dates) for (let s = 0; s < SHARDS; s++) keys.push([d, 'd:' + d + ':' + s]);
  const docs = {}; for (let i = 0; i < keys.length; i += 24) {
    const part = keys.slice(i, i + 24); const vals = await Promise.all(part.map(([, k]) => st.get(k, { type: 'json' }).catch(() => null)));
    part.forEach(([d], j) => { if (vals[j]) { docs[d] = docs[d] || empty(); merge(docs[d], vals[j]); } });
  }
  const total = empty(); dates.forEach(d => { const x = docs[d] || empty(); out.push(Object.assign({ date: d }, x)); merge(total, x); });
  return { days: out, total, range: { from: dates[0], to: dates[dates.length - 1], n: days } };
}

async function handle(event, deps) {
  if (event.httpMethod !== 'POST') return J(405, { ok: false });
  let b = {}; try { b = JSON.parse(event.body || '{}'); } catch (e) { return J(400, { ok: false }); }
  const st = deps.store;
  if (b.action) {                                         // ---- المدير ----
    const pw = process.env.ADMIN_PASSWORD;
    { const gt = await guard.gate(event, b.pw, pw, 'track', deps && deps.audit); if (!gt.ok) return J(gt.code, { ok: false, error: gt.error }); }
    if (!st) return J(500, { ok: false, error: 'storage' });
    try {
      if (b.action === 'stats') return J(200, Object.assign({ ok: true }, await stats(st, b.days)));
      if (b.action === 'audit') {
        const ast = deps.audit || guard.open(event); await guard.purge(ast).catch(() => 0);
        const evs = await guard.events(ast, 300, 30);
        const cut = { d1: NOW() - 86400000, d7: NOW() - 7 * 86400000 }, sum = { d1: { ok: 0, fail: 0, locked: 0, rate: 0, admin: 0 }, d7: { ok: 0, fail: 0, locked: 0, rate: 0, admin: 0 } };
        const byIp = {}, lockedIds = {};
        evs.forEach(e => { ['d1', 'd7'].forEach(k => { if (e.t >= cut[k]) sum[k][e.r] = (sum[k][e.r] || 0) + 1; });
          if (e.t >= cut.d7 && (e.r === 'fail' || e.r === 'locked' || e.r === 'rate' || e.r === 'admin')) { byIp[e.iph] = (byIp[e.iph] || 0) + 1; }
          if (e.t >= cut.d1 && e.r === 'locked') lockedIds[e.idh] = e.id; });
        const suspects = Object.entries(byIp).filter(([, n]) => n >= 5).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([iph, n]) => ({ iph, n }));
        return J(200, { ok: true, summary: sum, suspects, lockedToday: Object.values(lockedIds), events: evs.slice(0, 60), getUsed: evs.filter(e => e.m === 'GET' && e.t >= cut.d7).length });
      }
      return J(400, { ok: false, error: 'action' });
    } catch (e) { return J(500, { ok: false, error: 'storage' }); }
  }
  // ---- نبضة عامة ----
  const ua = (event.headers || {})['user-agent'] || '';
  if (isBot(ua) || !st) return J(204, {});
  if (limited(ipOf(event))) return J(204, {});
  const page = PAGES.includes(b.p) ? b.p : 'other', group = /^\d{1,2}[AG]$/.test(b.g || '') ? b.g : '';
  const d = { page, hour: hourOf(NOW()), mobile: isMobile(ua), standalone: !!b.s };
  if (b.t === 'hit') { d.view = true; d.uniq = !!b.f; d.newDev = !!b.n; d.group = group; }
  else if (b.t === 'ev' && EVENTS.includes(b.e)) { d.event = b.e; }
  else return J(400, { ok: false });
  try { await bump(st, d); } catch (e) {}
  return J(204, {});
}
function open(event) { try { const bl = require('@netlify/blobs'); try { bl.connectLambda(event); } catch (e) {} return bl.getStore('site-stats'); } catch (e) { return null; } }
exports.handler = async (event) => handle(event, { store: open(event) });
exports._handle = handle; exports._setNow = f => { NOW = f || (() => Date.now()); guard._setNow(f); }; exports._rate = RATE;
