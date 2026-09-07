/**
 * MSI Quiz — Cloudflare Worker Proxy
 *
 * Este Worker recibe las solicitudes del quiz y las reenvía a Anthropic,
 * manteniendo la API key segura del lado del servidor.
 *
 * Variables de entorno necesarias (configurar en Cloudflare Dashboard → Workers → Settings → Variables):
 *   ANTHROPIC_API_KEY  →  tu API key de Anthropic (marcar como "Encrypt")
 *
 * Pasos de deploy: ver README en esta misma carpeta.
 */

// Orígenes autorizados a usar este proxy.
// Reemplazá con el dominio real donde está publicado el quiz.
// Ejemplos:
//   'https://TU-USUARIO.github.io'
//   'https://tu-dominio.com'
// Dejá 'http://localhost' y 'http://127.0.0.1' para pruebas locales.
const ALLOWED_ORIGINS = [
  'https://rodrigomarquezmiranda.github.io',
  'http://localhost',
  'http://127.0.0.1',
];

export default {
  async fetch(request, env) {
    const origin = request.headers.get('Origin') || '';
    const allowed = ALLOWED_ORIGINS.some(o => origin === o || origin.startsWith(o + '/'));

    const cors = {
      'Access-Control-Allow-Origin':  allowed ? origin : 'null',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Vary': 'Origin',
    };

    // Preflight CORS
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: cors });
    }

    // Bloquear orígenes no autorizados
    if (!allowed) {
      return new Response(JSON.stringify({ error: { message: 'Origen no autorizado.' } }), {
        status: 403,
        headers: { 'Content-Type': 'application/json' },
      });
    }

    if (request.method !== 'POST') {
      return new Response('Method Not Allowed', { status: 405, headers: cors });
    }

    // Validar que la API key esté configurada
    if (!env.ANTHROPIC_API_KEY) {
      return new Response(JSON.stringify({ error: { message: 'Worker sin configurar: falta ANTHROPIC_API_KEY.' } }), {
        status: 500,
        headers: { 'Content-Type': 'application/json', ...cors },
      });
    }

    let body;
    try {
      body = await request.json();
    } catch {
      return new Response(JSON.stringify({ error: { message: 'Body inválido.' } }), {
        status: 400,
        headers: { 'Content-Type': 'application/json', ...cors },
      });
    }

    // Reenviar a Anthropic
    try {
      const res = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type':      'application/json',
          'x-api-key':         env.ANTHROPIC_API_KEY,
          'anthropic-version': '2023-06-01',
        },
        body: JSON.stringify(body),
      });

      const data = await res.json();

      return new Response(JSON.stringify(data), {
        status: res.status,
        headers: { 'Content-Type': 'application/json', ...cors },
      });
    } catch (err) {
      return new Response(JSON.stringify({ error: { message: 'Error al contactar Anthropic: ' + err.message } }), {
        status: 502,
        headers: { 'Content-Type': 'application/json', ...cors },
      });
    }
  },
};
