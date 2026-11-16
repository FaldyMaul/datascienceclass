// GET /api/info — is the results store connected, and how many results are saved.
const store = require('../lib/store');

module.exports = async (req, res) => {
  store.cors(res);
  let count = 0, ok = store.ready();
  if (ok) { try { count = await store.count(); } catch (e) { ok = false; } }
  res.status(200).json({ store: ok, count, resetEnabled: Boolean(process.env.ADMIN_KEY) });
};
