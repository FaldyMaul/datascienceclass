// POST /api/submit — a student page sends one result (game score or mission answer).
const store = require('../lib/store');

const ALLOWED = ['kind', 'name', 'score', 'cutoff', 'net', 'found', 'mission', 'answer', 'correct', 'title'];

function clean(body) {
  const out = {};
  for (const key of ALLOWED) {
    const v = body[key];
    if (typeof v === 'boolean' || (typeof v === 'number' && Number.isFinite(v))) out[key] = v;
    else if (typeof v === 'string') out[key] = v.trim().slice(0, 40);
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
  if (!['game', 'mission'].includes(entry.kind) || !entry.name) return res.status(400).json({ ok: false });
  entry.t = Math.floor(Date.now() / 1000);
  try { await store.add(entry); } catch (e) { return res.status(502).json({ ok: false }); }
  res.status(200).json({ ok: true });
};
