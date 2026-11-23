// Firepit runs the app in an opaque frame. Its authenticated parent owns the
// SharePoint connection and form digest; do not attempt direct cookie access.
export function getFirepitBridge(win = typeof window === 'undefined' ? undefined : window) {
  const bridge = win?.firepit?.sharepoint || win?.firepitSharePoint || win?.__firepitSharePointBridge;
  return bridge && typeof bridge.fetch === 'function' && typeof bridge.siteUrl === 'string' ? bridge : null;
}

const canonicalUrl = value => {
  const url = new URL(value);
  if (/\.(?:sharepoint\.com|sharepoint\.us|sharepoint-mil\.us|sharepoint\.de|sharepoint\.cn)\.mcas-gov\.us$/i.test(url.hostname)) {
    url.hostname = url.hostname.slice(0, -'.mcas-gov.us'.length);
  }
  return url;
};
const normalizeWeb = value => {
  const url = canonicalUrl(value);
  return `${url.origin}${url.pathname.replace(/\/+$/, '')}`.toLowerCase();
};
const transports = new WeakSet();

export function withFirepitTransport(webUrl, fallback) {
  if (transports.has(fallback)) return fallback;
  const transport = async (input, init = {}) => {
    const bridge = getFirepitBridge();
    if (!bridge) return fallback(input, init);
    if (normalizeWeb(webUrl) !== normalizeWeb(bridge.siteUrl)) {
      throw new Error(`The Tracker is configured for ${webUrl}, but Firepit is hosted on ${bridge.siteUrl}. Open the Tracker in Firepit on its configured site. No request was sent and the data destination was not changed.`);
    }
    const url = canonicalUrl(new URL(typeof input === 'string' || input instanceof URL ? String(input) : input.url, `${webUrl}/`).href);
    const web = canonicalUrl(webUrl);
    if (url.protocol !== 'https:' || url.origin !== web.origin || url.username || url.password || url.hash ||
        !url.pathname.toLowerCase().startsWith(`${web.pathname.replace(/\/$/, '').toLowerCase()}/_api/`)) {
      throw new Error('This request is outside the configured Tracker REST service. No request was sent. Keep the Tracker and its data in the hosting SharePoint web.');
    }
    const headers = new Headers(init.headers || (input instanceof Request ? input.headers : undefined));
    const method = String(headers.get('X-HTTP-Method') || headers.get('X-HTTP-Method-Override') || init.method || (input instanceof Request ? input.method : 'GET')).toUpperCase();
    if (!['GET', 'HEAD'].includes(method)) {
      // Firepit permits group membership writes through the canonical selector.
      url.pathname = url.pathname.replace(/\/_api\/web\/sitegroups\((\d+)\)(?=\/|$)/i, '/_api/web/sitegroups/getbyid($1)');
      if (/\/_api\/web\/roledefinitions(?:\/|\(|$)/i.test(url.pathname)) {
        throw new Error('Firepit cannot create or edit SharePoint permission levels. A site owner must create the documented Tracker permission levels in SharePoint first, then retry setup. Existing permissions were not broadened.');
      }
    }
    // These values belong to the trusted host, including for multipart writes.
    headers.delete('X-RequestDigest'); headers.delete('Authorization'); headers.delete('Content-Length');
    let request = url.href;
    if (input instanceof Request) {
      request = new Request(url.href, {
        method: input.method, headers, credentials: input.credentials,
        signal: input.signal,
        ...(!['GET', 'HEAD'].includes(input.method) ? { body: await input.clone().arrayBuffer() } : {}),
      });
    }
    const controller = new AbortController();
    const caller = init.signal || (input instanceof Request ? input.signal : undefined);
    const cancel = () => controller.abort(caller.reason);
    if (caller?.aborted) cancel();
    else caller?.addEventListener('abort', cancel, { once: true });
    let timer, onAbort;
    try {
      const deadline = new Promise((_, reject) => {
        onAbort = () => reject(controller.signal.reason || new DOMException('Request cancelled', 'AbortError'));
        if (controller.signal.aborted) { onAbort(); return; }
        controller.signal.addEventListener('abort', onAbort, { once: true });
        timer = setTimeout(() => controller.abort(new Error('SharePoint request timed out. The operation may have completed; refresh before retrying.')), 60000);
      });
      if (controller.signal.aborted) return await deadline;
      return await Promise.race([bridge.fetch.call(bridge, request, { ...init, headers, signal: controller.signal }), deadline]);
    } finally {
      clearTimeout(timer);
      caller?.removeEventListener('abort', cancel);
      controller.signal.removeEventListener('abort', onAbort);
    }
  };
  transports.add(transport);
  return transport;
}

export function fileDownloadUrl(webUrl, fileUrl) {
  if (!getFirepitBridge()) return fileUrl;
  const url = canonicalUrl(new URL(fileUrl, `${webUrl}/`).href), web = canonicalUrl(webUrl);
  if (url.origin !== web.origin || url.username || url.password || url.hash ||
      !url.pathname.toLowerCase().startsWith(`${web.pathname.replace(/\/$/, '').toLowerCase()}/`)) {
    throw new Error('The document is outside the configured Tracker site. No download request was sent.');
  }
  const path = encodeURIComponent(decodeURIComponent(url.pathname).replace(/'/g, "''"));
  return `${webUrl.replace(/\/$/, '')}/_api/web/GetFileByServerRelativePath(decodedurl='${path}')/$value`;
}

export function requireNativeInvitations() {
  if (getFirepitBridge()) throw new Error('Firepit does not support SharePoint invitation emails. Use the SharePoint page’s Share dialog to send the invitation. Manage Tracker roles separately in Users and managers; this action made no membership changes and sent no email.');
}
