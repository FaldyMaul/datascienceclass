// GET /api/board — every result so far. Read by the lecturer's deck and by teacher.html.
const store = require('../lib/store');

module.exports = async (req, res) => {
  store.cors(res);
  if (req.method === 'OPTIONS') return res.status(204).end();
  if (!store.ready()) return res.status(200).json({ entries: [], store: false });
  try { res.status(200).json({ entries: await store.all(), store: true }); }
  catch (e) { res.status(502).json({ entries: [], store: false }); }
};
