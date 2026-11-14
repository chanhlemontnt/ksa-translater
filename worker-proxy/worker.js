/**
 * Cloudflare Worker Reverse Proxy for Unified Single Domain: https://dich.tnt.name.vn
 *
 * Routing:
 * - /api/*     -> Google Cloud Run Backend
 * - /*         -> Cloudflare Pages Frontend
 */

export default {
  async fetch(request, env, ctx) {
    try {
      const incomingUrl = new URL(request.url);
      const isApiRoute = incomingUrl.pathname === '/api' || incomingUrl.pathname.startsWith('/api/');

      if (isApiRoute) {
        return await proxyToCloudRun(request, incomingUrl, env);
      }
      return await proxyToCloudflarePages(request, incomingUrl, env);
    } catch (err) {
      return new Response(`Worker Error: ${err.message}\n${err.stack}`, {
        status: 500,
        headers: { 'Content-Type': 'text/plain; charset=utf-8' }
      });
    }
  },
};

async function proxyToCloudRun(request, incomingUrl, env) {
  if (!env.CLOUD_RUN_ORIGIN) {
    return Response.json(
      { detail: 'Worker misconfigured: missing CLOUD_RUN_ORIGIN in Settings -> Variables.' },
      { status: 502 }
    );
  }

  let backendBase = env.CLOUD_RUN_ORIGIN.trim().replace(/\/+$/, '');
  if (!backendBase.startsWith('http://') && !backendBase.startsWith('https://')) {
    backendBase = 'https://' + backendBase;
  }

  const targetUrl = backendBase + incomingUrl.pathname + incomingUrl.search;
  const headers = new Headers(request.headers);
  headers.delete('host');

  if (env.WORKER_SHARED_SECRET) {
    headers.set('X-Worker-Shared-Secret', env.WORKER_SHARED_SECRET);
  }
  headers.set('X-Forwarded-Host', incomingUrl.host);
  headers.set('X-Forwarded-Proto', 'https');

  const cfJwt = request.headers.get('Cf-Access-Jwt-Assertion');
  if (cfJwt) {
    headers.set('Cf-Access-Jwt-Assertion', cfJwt);
  }

  const upstreamResp = await fetch(targetUrl, {
    method: request.method,
    headers: headers,
    body: ['GET', 'HEAD'].includes(request.method) ? null : request.body,
    redirect: 'manual',
  });

  return new Response(upstreamResp.body, {
    status: upstreamResp.status,
    statusText: upstreamResp.statusText,
    headers: upstreamResp.headers,
  });
}

async function proxyToCloudflarePages(request, incomingUrl, env) {
  if (!env.PAGES_ORIGIN) {
    return new Response('Lỗi: Chưa cấu hình biến PAGES_ORIGIN trong Worker Settings -> Variables.', {
      status: 502,
      headers: { 'Content-Type': 'text/plain; charset=utf-8' }
    });
  }

  let pagesBase = env.PAGES_ORIGIN.trim().replace(/\/+$/, '');
  if (!pagesBase.startsWith('http://') && !pagesBase.startsWith('https://')) {
    pagesBase = 'https://' + pagesBase;
  }

  const targetUrl = pagesBase + incomingUrl.pathname + incomingUrl.search;
  const headers = new Headers(request.headers);
  headers.delete('host');

  const pagesResp = await fetch(targetUrl, {
    method: request.method,
    headers: headers,
    redirect: 'follow',
  });

  return new Response(pagesResp.body, {
    status: pagesResp.status,
    statusText: pagesResp.statusText,
    headers: pagesResp.headers,
  });
}
