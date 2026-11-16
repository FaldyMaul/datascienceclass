// Tiny Redis-over-HTTP client (Upstash REST API). No dependencies.
// Vercel's Upstash Redis integration provides KV_REST_API_URL / KV_REST_API_TOKEN;
// a database created directly at Upstash provides UPSTASH_REDIS_REST_URL / UPSTASH_REDIS_REST_TOKEN.
const BASE = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL || '';
const TOKEN = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN || '';
const KEY = 'class:entries';
const MAX_ENTRIES = 5000;

const ready = () => Boolean(BASE && TOKEN);

async function redis(command) {
  const r = await fetch(BASE, {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + TOKEN, 'Content-Type': 'application/json' },
    body: JSON.stringify(command)
  });
  if (!r.ok) throw new Error('store ' + r.status);
  return (await r.json()).result;
}

async function add(entry) {
  const n = await redis(['RPUSH', KEY, JSON.stringify(entry)]);
  if (n > MAX_ENTRIES) await redis(['LTRIM', KEY, String(-MAX_ENTRIES), '-1']);
}

async function all() {
  const rows = (await redis(['LRANGE', KEY, '0', '-1'])) || [];
  const out = [];
  for (const row of rows) { try { out.push(JSON.parse(row)); } catch (e) { /* skip a damaged row */ } }
  return out;
}

const clear = () => redis(['DEL', KEY]);
const count = () => redis(['LLEN', KEY]);

function cors(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Admin-Key');
  res.setHeader('Cache-Control', 'no-store');
}

module.exports = { ready, add, all, clear, count, cors };
