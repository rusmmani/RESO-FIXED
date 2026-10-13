/**
 * RESO Haircut API proxy -> Supabase Edge Function.
 * Required Vercel env: SUPABASE_URL, SUPABASE_ANON_KEY.
 */
const SUPABASE_URL = String(process.env.SUPABASE_URL || '').replace(/\/$/, '');
const SUPABASE_KEY = String(process.env.SUPABASE_ANON_KEY || '').trim();

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  if (!['GET', 'POST', 'OPTIONS'].includes(req.method)) {
    res.setHeader('Allow', 'GET, POST, OPTIONS');
    return res.status(405).json({ ok: false, error: 'Method tidak diizinkan.' });
  }
  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-admin-pin');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    return res.status(204).end();
  }
  if (!SUPABASE_URL || !SUPABASE_KEY) {
    return res.status(500).json({ ok: false, error: 'Konfigurasi Supabase belum lengkap di Vercel.' });
  }
  const endpoint = SUPABASE_URL + '/functions/v1/reso-api';
  const url = new URL(endpoint);
  if (req.method === 'GET') {
    for (const [k, v] of Object.entries(req.query || {})) {
      if (v !== undefined && v !== null) url.searchParams.set(k, Array.isArray(v) ? String(v[0]) : String(v));
    }
  }
  try {
    const headers = { apikey: SUPABASE_KEY, Authorization: 'Bearer ' + SUPABASE_KEY, Accept: 'application/json' };
    if (req.headers['x-admin-pin']) headers['x-admin-pin'] = String(req.headers['x-admin-pin']);
    let body;
    if (req.method === 'POST') {
      headers['Content-Type'] = 'application/json';
      body = typeof req.body === 'string' ? req.body : JSON.stringify(req.body || {});
    }
    const upstream = await fetch(url.toString(), { method: req.method, headers, body, cache: 'no-store' });
    const text = await upstream.text();
    let data;
    try { data = JSON.parse(text); } catch (_) {
      return res.status(502).json({ ok: false, error: 'Supabase mengembalikan respons non-JSON.', detail: text.slice(0, 400) });
    }
    return res.status(upstream.status).json(data);
  } catch (e) {
    return res.status(502).json({ ok: false, error: 'Gagal menghubungi Supabase.', detail: e && e.message ? e.message : String(e) });
  }
};
