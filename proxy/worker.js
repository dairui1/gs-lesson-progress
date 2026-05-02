/**
 * Cloudflare Worker: Zoom webhook proxy → Apps Script
 *
 * Why this exists:
 *   Apps Script Web App responds to POST via a 302 redirect to
 *   script.googleusercontent.com, which fails Zoom's URL validation
 *   security scan. This worker terminates the Zoom request directly
 *   (200 OK, no redirect), handles URL validation locally with HMAC,
 *   and forwards real events to Apps Script via background fetch.
 *
 * Required secrets (Worker → Settings → Variables and Secrets):
 *   APPS_SCRIPT_URL       e.g. https://script.google.com/macros/s/AKfyc.../exec
 *   URL_TOKEN             matches Apps Script Property ZOOM_URL_TOKEN
 *   ZOOM_WEBHOOK_SECRET   matches Zoom Event Subscription Secret Token
 */

export default {
  async fetch(req, env, ctx) {
    if (req.method === 'GET') {
      return json({ ok: true, name: 'zoom-proxy' });
    }
    if (req.method !== 'POST') {
      return new Response('Method Not Allowed', { status: 405 });
    }

    const raw = await req.text();
    let body = {};
    try { body = JSON.parse(raw); } catch (_) {}

    if (body.event === 'endpoint.url_validation') {
      const plainToken = (body.payload && body.payload.plainToken) || '';
      const encryptedToken = await hmacHex(env.ZOOM_WEBHOOK_SECRET, plainToken);
      return json({ plainToken, encryptedToken });
    }

    // Forward in background so Zoom gets 200 immediately (avoids 3s timeout).
    ctx.waitUntil(
      fetch(`${env.APPS_SCRIPT_URL}?token=${env.URL_TOKEN}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: raw
      }).catch(err => console.error('forward failed:', err && err.message))
    );

    return json({ ok: true });
  }
};

async function hmacHex(secret, message) {
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey(
    'raw', enc.encode(secret),
    { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']
  );
  const sig = await crypto.subtle.sign('HMAC', key, enc.encode(message));
  return [...new Uint8Array(sig)].map(b => b.toString(16).padStart(2, '0')).join('');
}

function json(obj) {
  return new Response(JSON.stringify(obj), {
    headers: { 'Content-Type': 'application/json' }
  });
}
