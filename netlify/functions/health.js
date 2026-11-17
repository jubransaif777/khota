// ============================================================
//  مراقبة مصادر الموقع. يعمل بثلاث طرق:
//   1) تلقائياً كل 30 دقيقة (health-cron.js + netlify.toml)  2) «افحص الآن» من لوحة الإدارة  3) اختبار التنبيه.
//  التنبيه (اختياري) عبر تيليجرام: TELEGRAM_BOT_TOKEN + TELEGRAM_CHAT_ID، أو رابط عام: ALERT_WEBHOOK_URL.
//  لا ينبّه إلا بعد فشلين متتاليين، وينبّه مرة واحدة، ثم يخبرك حين يعود المصدر.
// ============================================================
const crypto = require('crypto');
const SITE = () => (process.env.SITE_URL || process.env.URL || 'https://khotamaereid.net').replace(/\/+$/, '');
const J = (code, obj) => ({ statusCode: code, headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }, body: JSON.stringify(obj) });
const sha = x => crypto.createHash('sha256').update(String(x)).digest();
const sameSecret = (a, b) => crypto.timingSafeEqual(sha(a), sha(b));
let NOW = () => Date.now();
const FAILS_TO_ALERT = 2, REMIND_MS = 12 * 3600000, COOLDOWN_MS = 20000, KEEP_HISTORY = 60, SLOW_MS = 8000;

// ---- التقييم: كل فحص يعيد {status:'ok'|'warn'|'fail', detail} ----
const sumLen = o => Object.values(o || {}).reduce((n, a) => n + (Array.isArray(a) ? a.length : 0), 0);
const CHECKS = [
  { id: 'plan', label: 'الخطة الأسبوعية (ملف SharePoint)', path: '/.netlify/functions/plan',
    eval(r) { if (!r.json || !Array.isArray(r.json.weeks) || !r.json.weeks.length) return ['fail', 'لا تُرجع الدالة أسابيع: تأكّد من رابط الملف وصلاحية المشاركة.'];
      const w = r.json.weeks[r.json.currentWeekIndex || 0]; const n = w ? sumLen(w.sections) : 0;
      return n ? ['ok', 'الأسبوع الحالي «' + (w.label || '') + '»، عدد الدروس المنشورة: ' + n] : ['warn', 'الملف يُقرأ، لكن لا دروس منشورة للأسبوع الحالي («' + ((w && w.label) || '') + '»).']; } },
  { id: 'exams', label: 'جدول الاختبارات (ملف SharePoint)', path: '/.netlify/functions/exams',
    eval(r) { if (!r.json || !Array.isArray(r.json.weeks) || !r.json.weeks.length) return ['fail', 'لا تُرجع الدالة أسابيع: تأكّد من رابط الملف.'];
      const w = r.json.weeks[r.json.currentWeekIndex || 0]; let n = 0; if (w && w.sections) Object.values(w.sections).forEach(days => (days || []).forEach(d => { n += (d.exams || []).length; }));
      return n ? ['ok', 'الأسبوع الحالي، عدد مواعيد الاختبارات: ' + n] : ['warn', 'الملف يُقرأ، لكن لا اختبارات مجدولة للأسبوع الحالي.']; } },
  { id: 'schedule', label: 'جدول الحصص (ملف SharePoint)', path: '/.netlify/functions/schedule',
    eval(r) { const s = r.json && r.json.status; if (s === 'ok') return ['ok', 'عدد الشعب: ' + (r.json.sections || []).length]; if (s === 'nourl') return ['warn', 'غير مُعدّ: لم يُضبط SCHEDULE_SHARE_URL.'];
      return ['fail', s === 'badformat' ? 'الملف لا يحوي تبويب schedule بالصيغة المتوقعة.' : 'تعذّر تحميل الملف (' + (s || 'لا رد') + ').']; } },
  { id: 'content', label: 'مركز المحتوى (التخزين)', path: '/.netlify/functions/content',
    eval(r) { if (!r.json || !r.json.ok) return ['fail', 'الدالة content لا تستجيب.']; return r.json.degraded ? ['fail', 'مخزن Blobs غير متاح: تحقّق من @netlify/blobs وأعد النشر.'] : ['ok', 'يعمل.']; } },
  { id: 'alef', label: 'بيانات ألف (ملف SharePoint)', path: '/.netlify/functions/alef?ping=1',
    eval(r) { const j = r.json || {}; if (j.status === 'ok') return j.rows > 0 ? ['ok', 'عدد الطلاب في الملف: ' + j.rows] : ['warn', 'الملف يُقرأ لكنه فارغ.']; return ['fail', j.status === 'badformat' ? 'صيغة الملف غير متوقعة (UAEID/Email/Password).' : 'تعذّر تحميل الملف.']; } },
  { id: 'report', label: 'تقارير المستوى (ملف SharePoint)', path: '/.netlify/functions/report?ping=1',
    eval(r) { const j = r.json || {}; if (j.status === 'ok') return j.rows > 0 ? ['ok', 'عدد أسطر التقارير: ' + j.rows] : ['warn', 'الملف يُقرأ لكنه فارغ.']; return ['fail', 'تعذّر تحميل الملف.']; } },
];

async function getJson(fetchFn, url) {
  const t0 = NOW(); const ctl = typeof AbortController !== 'undefined' ? new AbortController() : null; const timer = ctl ? setTimeout(() => ctl.abort(), 15000) : null;
  try { const res = await fetchFn(url, { signal: ctl && ctl.signal, headers: { 'user-agent': 'khota-health/1.0', 'cache-control': 'no-cache' } }); let json = null; try { json = await res.json(); } catch (e) {} return { ok: res.ok, status: res.status, json, ms: NOW() - t0 }; }
  catch (e) { return { ok: false, status: 0, json: null, ms: NOW() - t0, err: String(e && e.name || e) }; }
  finally { if (timer) clearTimeout(timer); }
}
async function runChecks(fetchFn) {
  const out = await Promise.all(CHECKS.map(async c => {
    const r = await getJson(fetchFn, SITE() + c.path); let st, detail;
    if (r.status === 0) { st = 'fail'; detail = r.err === 'AbortError' ? 'لم يرد خلال 15 ثانية.' : 'تعذّر الاتصال بالدالة.'; }
    else { [st, detail] = c.eval(r); if (st === 'ok' && r.ms > SLOW_MS) { st = 'warn'; detail += ' (بطيء: ' + Math.round(r.ms / 1000) + ' ثوانٍ)'; } }
    return { id: c.id, label: c.label, status: st, detail, ms: r.ms };
  }));
  return out;
}

// ---- التنبيه ----
const alertConfig = () => ({ telegram: !!(process.env.TELEGRAM_BOT_TOKEN && process.env.TELEGRAM_CHAT_ID), webhook: !!process.env.ALERT_WEBHOOK_URL });
async function sendAlert(fetchFn, text, extra) {
  const res = { telegram: null, webhook: null };
  if (process.env.TELEGRAM_BOT_TOKEN && process.env.TELEGRAM_CHAT_ID) {
    try { const r = await fetchFn('https://api.telegram.org/bot' + process.env.TELEGRAM_BOT_TOKEN + '/sendMessage', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ chat_id: process.env.TELEGRAM_CHAT_ID, text }) }); res.telegram = r.ok; } catch (e) { res.telegram = false; }
  }
  if (process.env.ALERT_WEBHOOK_URL) {
    try { const r = await fetchFn(process.env.ALERT_WEBHOOK_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(Object.assign({ text, source: 'khotamaereid.net' }, extra || {})) }); res.webhook = r.ok; } catch (e) { res.webhook = false; }
  }
  return res;
}
const emptyState = () => ({ at: 0, checks: [], fails: {}, alerted: {}, history: [], lastSched: 0 });
async function readState(st) { const s = st ? await st.get('state', { type: 'json' }).catch(() => null) : null; return Object.assign(emptyState(), s || {}); }

// الفحص + تحديث الحالة + التنبيه عند الحاجة
async function runAndAlert(deps, scheduled) {
  const st = deps.store, fetchFn = deps.fetch || fetch, state = await readState(st), t = NOW();
  const checks = await runChecks(fetchFn); const msgs = [];
  checks.forEach(c => {
    const was = (state.checks.find(x => x.id === c.id) || {}).status;
    if (c.status === 'fail') {
      state.fails[c.id] = (state.fails[c.id] || 0) + 1;
      if (state.fails[c.id] >= FAILS_TO_ALERT && !state.alerted[c.id]) { state.alerted[c.id] = t; msgs.push('🔴 ' + c.label + '\nلا يعمل: ' + c.detail); }
      else if (state.alerted[c.id] && t - state.alerted[c.id] >= REMIND_MS) { state.alerted[c.id] = t; msgs.push('🔴 ما زال معطّلاً: ' + c.label + '\n' + c.detail); }
    } else {
      if (state.alerted[c.id]) msgs.push('🟢 عاد للعمل: ' + c.label);
      state.fails[c.id] = 0; delete state.alerted[c.id];
    }
    if (was && was !== c.status && (was === 'fail' || c.status === 'fail')) state.history.unshift({ t, id: c.id, label: c.label, from: was, to: c.status, detail: c.detail });
  });
  state.history = state.history.slice(0, KEEP_HISTORY); state.at = t; state.checks = checks; if (scheduled) state.lastSched = t;
  let delivery = null; if (msgs.length) delivery = await sendAlert(fetchFn, 'مراقبة موقع الخطة الأسبوعية\n\n' + msgs.join('\n\n') + '\n\n' + SITE() + '/admin', { checks });
  if (delivery) state.history.unshift({ t, id: '_alert', label: 'إرسال تنبيه', from: '', to: '', detail: JSON.stringify(delivery) });
  state.history = state.history.slice(0, KEEP_HISTORY);
  if (st) { try { await st.setJSON('state', state); } catch (e) {} }
  return { state, sent: msgs.length, delivery };
}

async function handle(event, deps) {
  if (event.httpMethod !== 'POST') return J(405, { ok: false });
  let b = {}; try { b = JSON.parse(event.body || '{}'); } catch (e) { return J(400, { ok: false }); }
  const pw = process.env.ADMIN_PASSWORD;
  if (!pw || !b.pw || !sameSecret(b.pw, pw)) return J(401, { ok: false, error: 'auth' });
  const cfg = Object.assign(alertConfig(), { site: SITE() });
  try {
    if (b.action === 'status') { const s = await readState(deps.store); return J(200, { ok: true, state: s, config: cfg, now: NOW(), storage: !!deps.store }); }
    if (b.action === 'run') {
      const s0 = await readState(deps.store); if (NOW() - s0.at < COOLDOWN_MS) return J(200, { ok: true, state: s0, config: cfg, now: NOW(), cooldown: true });
      const r = await runAndAlert(deps, false); return J(200, { ok: true, state: r.state, config: cfg, now: NOW(), sent: r.sent });
    }
    if (b.action === 'test_alert') {
      if (!cfg.telegram && !cfg.webhook) return J(200, { ok: true, configured: false });
      const d = await sendAlert(deps.fetch || fetch, '✅ تجربة تنبيه: مراقبة موقع الخطة الأسبوعية تعمل.\n' + SITE() + '/admin'); return J(200, { ok: true, configured: true, delivery: d });
    }
    return J(400, { ok: false, error: 'action' });
  } catch (e) { return J(500, { ok: false, error: 'error' }); }
}
function open(event) { try { const bl = require('@netlify/blobs'); try { bl.connectLambda(event); } catch (e) {} return bl.getStore('site-health'); } catch (e) { return null; } }
exports.handler = async (event) => handle(event, { store: open(event) });
exports._handle = handle; exports._run = runAndAlert; exports._open = open; exports._setNow = f => { NOW = f || (() => Date.now()); }; exports._CHECKS = CHECKS;
