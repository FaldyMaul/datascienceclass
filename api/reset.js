// POST /api/reset — clears all results. Needs the ADMIN_KEY set in the Vercel project settings.
const store = require('../lib/store');

module.exports = async (req, res) => {
  store.cors(res);
  if (req.method === 'OPTIONS') return res.status(204).end();
  if (req.method !== 'POST') return res.status(405).json({ ok: false });
  const key = process.env.ADMIN_KEY || '';
  if (!key || req.headers['x-admin-key'] !== key) return res.status(403).json({ ok: false, error: 'wrong or missing key' });
  if (!store.ready()) return res.status(503).json({ ok: false });
  try { await store.clear(); } catch (e) { return res.status(502).json({ ok: false }); }
  res.status(200).json({ ok: true });
};
