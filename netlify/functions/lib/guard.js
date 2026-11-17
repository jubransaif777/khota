// ============================================================
//  حارس الاستعلامات الحسّاسة (ألف والتقرير): حدّ محاولات دائم + سجلّ تدقيق بلا بيانات شخصية.
//  • الحدّ: 5 محاولات فاشلة لكل رقم هوية في اليوم (مهما تبدّل العنوان)، و30 محاولة لكل عنوان في الساعة.
//  • السجلّ: النتيجة والوقت وآخر 3 أرقام من الهوية وبصمة مبهَمة للعنوان. لا أسماء ولا كلمات مرور ولا أرقام كاملة.
//  يُخزَّن في Netlify Blobs (مخزن site-audit)، وإن تعذّر يعمل بحدّ في الذاكرة بدلاً من التوقف.
// ============================================================
const crypto = require('crypto');
const LIM = { idPerDay: 5, ipPerHour: 30 };
const KEEP_DAYS = 30;
let NOW = () => Date.now();                       // للاختبار
const mem = new Map();                             // بديل عند تعذّر المخزن

const salt = () => process.env.AUDIT_SALT || process.env.ADMIN_PASSWORD || 'khota-audit';
const hmac = x => crypto.createHmac('sha256', salt()).update(String(x)).digest('hex').slice(0, 16);
const dayKey = () => new Date(NOW()).toLocaleDateString('en-CA', { timeZone: 'Asia/Dubai' });
const hourKey = () => new Date(NOW()).toLocaleString('sv-SE', { timeZone: 'Asia/Dubai' }).slice(0, 13).replace(' ', 'T');
const mask = id => { const s = String(id || '').replace(/\D/g, ''); return s ? '•'.repeat(Math.max(0, s.length - 3)) + s.slice(-3) : ''; };

function open(event) {
  try { const bl = require('@netlify/blobs'); try { bl.connectLambda(event); } catch (e) {} return bl.getStore('site-audit'); }
  catch (e) { return null; }
}
async function getN(st, key) {
  if (!st) return mem.get(key) || 0;
  try { const v = await st.get(key, { type: 'json' }); return (v && v.n) || 0; } catch (e) { return mem.get(key) || 0; }
}
async function bump(st, key) {
  const n = (await getN(st, key)) + 1; mem.set(key, n);
  if (st) { try { await st.setJSON(key, { n, d: dayKey() }); } catch (e) {} }
  return n;
}

// قبل البحث: هل يُسمح بالمحاولة؟ (تُحتسب محاولة العنوان هنا)
async function check(st, idDigits, ip) {
  const idh = hmac('id:' + idDigits), iph = hmac('ip:' + ip);
  const nId = await getN(st, 'f:' + dayKey() + ':' + idh);
  if (nId >= LIM.idPerDay) return { ok: false, reason: 'locked', idh, iph };
  const nIp = await bump(st, 'i:' + hourKey() + ':' + iph);
  if (nIp > LIM.ipPerHour) return { ok: false, reason: 'ratelimited', idh, iph };
  return { ok: true, idh, iph };
}
async function fail(st, ctx) { return bump(st, 'f:' + dayKey() + ':' + ctx.idh); }

async function log(st, ev) {
  const rec = Object.assign({ t: NOW() }, ev);
  mem.set('lastlog', rec);
  if (!st) return;
  try { await st.setJSON('e:' + String(9999999999999 - rec.t).padStart(13, '0') + ':' + crypto.randomBytes(3).toString('hex'), rec); } catch (e) {}
}
async function listKeys(st, prefix) {
  const keys = [];
  try { for await (const p of st.list({ prefix, paginate: true })) for (const b of p.blobs) keys.push(b.key); }
  catch (e) { keys.length = 0; try { const r = await st.list({ prefix }); (r.blobs || []).forEach(b => keys.push(b.key)); } catch (e2) {} }
  return keys;
}
// الأحداث (الأحدث أولاً)
async function events(st, limit, days) {
  if (!st) return [];
  const since = NOW() - (days || KEEP_DAYS) * 86400000, keys = (await listKeys(st, 'e:')).slice(0, limit || 300), out = [];
  for (let i = 0; i < keys.length; i += 25) {
    const vals = await Promise.all(keys.slice(i, i + 25).map(k => st.get(k, { type: 'json' }).catch(() => null)));
    vals.forEach(v => { if (v && v.t >= since) out.push(v); });
  }
  return out.sort((a, b) => b.t - a.t);
}
// حذف ما تجاوز مدة الحفظ (يُستدعى عند قراءة المدير)
async function purge(st) {
  if (!st) return 0; let n = 0; const old = NOW() - KEEP_DAYS * 86400000;
  const keys = await listKeys(st, 'e:'); const oldKeys = keys.slice(-200);        // الأقدم في آخر الترتيب (مفتاح معكوس)
  for (const k of oldKeys) { const v = await st.get(k, { type: 'json' }).catch(() => null); if (v && v.t < old) { try { await st.delete(k); n++; } catch (e) {} } }
  return n;
}
module.exports = { LIM, open, check, fail, log, events, purge, mask, hmac, dayKey,
  _setNow: f => { NOW = f || (() => Date.now()); }, _mem: mem };

