/**
 * RESO Haircut - Vercel API proxy -> Google Apps Script
 *
 * IMPORTANT:
 * Vercel Environment Variable:
 * GAS_URL=https://script.google.com/macros/s/DEPLOYMENT_ID/exec
 */

const GAS_URL = String(process.env.GAS_URL || 'https://script.google.com/macros/s/AKfycbyzX8z6NF0uhZjlvbi9eGDGWp2vfbn1Yt36TnqF817/exec').trim();

function sendJson(res, status, payload) {
  res.status(status);
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  return res.end(JSON.stringify(payload));
}

function parseJsonText(text) {
  const clean = String(text || '').trim();
  if (!clean) return null;

  try {
    return JSON.parse(clean);
  } catch (_) {}

  // Some Google/Apps Script error pages can contain a JSON object inside
  // surrounding HTML/text. Try to recover it for a useful diagnostic.
  const first = clean.indexOf('{');
  const last = clean.lastIndexOf('}');
  if (first >= 0 && last > first) {
    try {
      return JSON.parse(clean.slice(first, last + 1));
    } catch (_) {}
  }

  return null;
}

function looksLikeGoogleLoginOrError(text) {
  const s = String(text || '').toLowerCase();
  return s.includes('accounts.google.com') ||
    s.includes('sign in') ||
    s.includes('google apps script') ||
    s.includes('script.google.com') ||
    s.includes('<html');
}

async function handler(req, res) {
  if (!GAS_URL) {
    return sendJson(res, 500, {
      ok: false,
      error: 'GAS_URL belum diset di Vercel.',
      hint: 'Tambahkan Environment Variable GAS_URL dengan URL Web App Google Apps Script yang berakhiran /exec.'
    });
  }

  let url;
  try {
    url = new URL(GAS_URL);
  } catch (_) {
    return sendJson(res, 500, {
      ok: false,
      error: 'GAS_URL tidak valid.',
      hint: 'Gunakan URL Web App Apps Script, contoh: https://script.google.com/macros/s/DEPLOYMENT_ID/exec'
    });
  }

  // GET /api/reso?api=getClientBookingData&date=...&barber=...
  if (req.method === 'GET') {
    for (const [key, value] of Object.entries(req.query || {})) {
      if (value !== undefined && value !== null) {
        url.searchParams.set(key, Array.isArray(value) ? String(value[0]) : String(value));
      }
    }
  }

  try {
    let upstream;

    if (req.method === 'GET') {
      upstream = await fetch(url.toString(), {
        method: 'GET',
        redirect: 'follow',
        headers: { 'Accept': 'application/json, text/plain, */*' },
        cache: 'no-store'
      });
    } else if (req.method === 'POST') {
      let body = req.body;
      if (body === undefined || body === null) body = {};
      if (typeof body !== 'string') body = JSON.stringify(body);

      // text/plain is intentional: Apps Script Web App handles this reliably
      // without browser CORS/preflight issues.
      upstream = await fetch(url.toString(), {
        method: 'POST',
        redirect: 'follow',
        headers: {
          'Content-Type': 'text/plain;charset=UTF-8',
          'Accept': 'application/json, text/plain, */*'
        },
        body,
        cache: 'no-store'
      });
    } else {
      res.setHeader('Allow', 'GET, POST');
      return sendJson(res, 405, { ok: false, error: 'Method tidak diizinkan.' });
    }

    const text = await upstream.text();
    const data = parseJsonText(text);

    if (data === null) {
      const detail = String(text || '').replace(/\s+/g, ' ').slice(0, 800);
      const googleHint = looksLikeGoogleLoginOrError(text)
        ? 'Google Apps Script mengembalikan halaman HTML. Pastikan deployment Apps Script adalah Web app, Execute as: Me, dan Who has access: Anyone. Gunakan URL deployment yang berakhiran /exec.'
        : 'Pastikan GAS_URL menunjuk ke Web App Apps Script yang berakhiran /exec dan deployment dapat diakses publik.';

      return sendJson(res, 502, {
        ok: false,
        error: 'Response Google Apps Script bukan JSON.',
        upstreamStatus: upstream.status,
        upstreamContentType: upstream.headers.get('content-type') || '',
        upstreamUrl: upstream.url || url.toString(),
        hint: googleHint,
        detail
      });
    }

    if (data && data.ok === true && Object.prototype.hasOwnProperty.call(data, 'data')) {
      return sendJson(res, 200, data.data);
    }

    return sendJson(res, upstream.ok ? 200 : upstream.status, data);
  } catch (err) {
    return sendJson(res, 502, {
      ok: false,
      error: 'Gagal menghubungi Google Apps Script.',
      detail: err && err.message ? err.message : String(err)
    });
  }
}

module.exports = handler;
