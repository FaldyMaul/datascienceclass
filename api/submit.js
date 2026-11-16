// POST /api/submit — a student page sends one result (mission answer or BI challenge dashboard).
const store = require('../lib/store');

// field → maximum length for text values
const ALLOWED = { kind: 12, name: 40, mission: 40, answer: 40, correct: 40, track: 40, url: 200, insight: 200 };
const KINDS = ['mission', 'case'];
const DASHBOARD = /^https:\/\/(datastudio|lookerstudio)\.google\.com\/[^\s"'<>]+$/;

function clean(body) {
  const out = {};
  for (const key of Object.keys(ALLOWED)) {
    const v = body[key];
    if (typeof v === 'boolean' || (typeof v === 'number' && Number.isFinite(v))) out[key] = v;
    else if (typeof v === 'string') out[key] = v.trim().slice(0, ALLOWED[key]);
  }
  return out;
}

module.exports = async (req, res) => {
  store.cors(res);
  if (req.method === 'OPTIONS') return res.status(204).end();
  if (req.method !== 'POST') return res.status(405).json({ ok: false });
  if (!store.ready()) return res.status(503).json({ ok: false, error: 'results store not connected' });
  let body = req.body;
  try { if (typeof body === 'string') body = JSON.parse(body); } catch (e) { body = null; }
  if (!body || typeof body !== 'object' || JSON.stringify(body).length > 2048) return res.status(400).json({ ok: false });
  const entry = clean(body);
  if (!KINDS.includes(entry.kind) || !entry.name) return res.status(400).json({ ok: false });
  if (entry.kind === 'case' && (typeof entry.url !== 'string' || !DASHBOARD.test(entry.url) || !entry.insight)) return res.status(400).json({ ok: false, error: 'dashboard link or recommendation missing' });
  if (entry.kind !== 'case') { delete entry.url; delete entry.insight; delete entry.track; }
  entry.t = Math.floor(Date.now() / 1000);
  try { await store.add(entry); } catch (e) { return res.status(502).json({ ok: false }); }
  res.status(200).json({ ok: true });
};
