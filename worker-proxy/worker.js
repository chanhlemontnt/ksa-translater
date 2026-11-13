/**
 * Cloudflare Worker Reverse Proxy for Unified Single Domain: https://dich.tnt.name.vn
 *
 * Routing:
 * - /api/*     -> Google Cloud Run Backend (attaches X-Worker-Shared-Secret & forwards Cf-Access-Jwt-Assertion)
 * - /*         -> Cloudflare Pages Frontend (Pages origin URL or static passthrough)
 *
 * Required Worker Environment Variables / Secrets (set via wrangler secret put / Dashboard):
 * - CLOUD_RUN_ORIGIN     : e.g. https://dich-hop-dong-api-xxxxx-uc.a.run.app
 * - PAGES_ORIGIN         : e.g. https://dich-hop-dong-fe.pages.dev
 * - WORKER_SHARED_SECRET : High-entropy random secret shared with Cloud Run
 */

export default {
  async fetch(request, env, ctx) {
    const incomingUrl = new URL(request.url);
    const isApiRoute = incomingUrl.pathname === '/api' || incomingUrl.pathname.startsWith('/api/');

    if (isApiRoute) {
      return proxyToCloudRun(request, incomingUrl, env);
    }
    return proxyToCloudflarePages(request, incomingUrl, env);
  },
};

async function proxyToCloudRun(request, incomingUrl, env) {
  if (!env.CLOUD_RUN_ORIGIN || !env.WORKER_SHARED_SECRET) {
    return Response.json(
      { detail: 'Worker misconfigured: missing CLOUD_RUN_ORIGIN or WORKER_SHARED_SECRET.' },
      { status: 502 }
    );
  }

  const backendBase = env.CLOUD_RUN_ORIGIN.replace(/\/+$/, '');
  const targetUrl = backendBase + incomingUrl.pathname + incomingUrl.search;

  const headers = new Headers(request.headers);
  // Inject shared secret so Cloud Run verifies traffic came exclusively through this Worker
  headers.set('X-Worker-Shared-Secret', env.WORKER_SHARED_SECRET);
  headers.set('X-Forwarded-Host', incomingUrl.host);
  headers.set('X-Forwarded-Proto', 'https');

  // Preserve Cf-Access-Jwt-Assertion injected by Cloudflare Access Zero Trust
  const cfJwt = request.headers.get('Cf-Access-Jwt-Assertion');
  if (cfJwt) {
    headers.set('Cf-Access-Jwt-Assertion', cfJwt);
  }

  // Set Host header expected by Cloud Run HTTPS ingress
  const cloudRunHost = new URL(backendBase).host;
  headers.set('Host', cloudRunHost);

  const upstreamReq = new Request(targetUrl, {
    method: request.method,
    headers,
    body: ['GET', 'HEAD'].includes(request.method) ? null : request.body,
    redirect: 'manual',
  });

  const upstreamResp = await fetch(upstreamReq);
  const respHeaders = new Headers(upstreamResp.headers);
  respHeaders.set('X-Content-Type-Options', 'nosniff');
  respHeaders.set('Referrer-Policy', 'strict-origin-when-cross-origin');

  return new Response(upstreamResp.body, {
    status: upstreamResp.status,
    statusText: upstreamResp.statusText,
    headers: respHeaders,
  });
}

async function proxyToCloudflarePages(request, incomingUrl, env) {
  // If Worker is bound only to dich.tnt.name.vn/api/* route, non-API requests hit Pages directly.
  // If Worker handles the entire dich.tnt.name.vn/* zone, forward UI requests to PAGES_ORIGIN.
  if (!env.PAGES_ORIGIN) {
    return fetch(request);
  }

  const pagesBase = env.PAGES_ORIGIN.replace(/\/+$/, '');
  const targetUrl = pagesBase + incomingUrl.pathname + incomingUrl.search;
  const headers = new Headers(request.headers);
  headers.set('Host', new URL(pagesBase).host);

  const pagesResp = await fetch(
    new Request(targetUrl, {
      method: request.method,
      headers,
      redirect: 'follow',
    })
  );

  const outHeaders = new Headers(pagesResp.headers);
  outHeaders.set('X-Frame-Options', 'DENY');
  outHeaders.set('X-Content-Type-Options', 'nosniff');
  return new Response(pagesResp.body, {
    status: pagesResp.status,
    statusText: pagesResp.statusText,
    headers: outHeaders,
  });
}
