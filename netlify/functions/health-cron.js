// الفحص التلقائي المجدول (كل 30 دقيقة، يُضبط الجدول في netlify.toml). لا يُستدعى من الإنترنت.
const health = require('./health.js');
exports.handler = async (event) => {
  try { const r = await health._run({ store: health._open(event) }, true); return { statusCode: 200, body: JSON.stringify({ ok: true, sent: r.sent }) }; }
  catch (e) { return { statusCode: 500, body: 'error' }; }
};
