import { afterEach, expect, it, vi } from 'vitest';
import { fileDownloadUrl, withFirepitTransport } from './firepit';
import { resolveWebUrl, resetWebUrlCache, spGet, spPost } from './spContext';
import { SharePointStore } from './spStore';
import { previewMigration } from './siteMigration';

const web = 'https://tenant.sharepoint-mil.us/sites/Tracker/Subweb';
const enable = (response = {}) => {
  const bridge = { siteUrl: web, fetch: vi.fn(async () => new Response(JSON.stringify(response), { headers: { 'Content-Type': 'application/json' } })) };
  window.firepit = { sharepoint: bridge };
  return bridge;
};
afterEach(() => { delete window.firepit; resetWebUrlCache(); });

it('discovers the exact hosting subweb in an opaque frame without parent access', () => {
  const bridge = enable();
  const win = { location: { origin: 'null', pathname: 'blank', search: '', hash: '' }, firepit: { sharepoint: bridge },
    get parent() { throw new Error('Opaque origin'); } };
  expect(resolveWebUrl(win)).toBe(web);
});

it('keeps an explicit forwarded webUrl override instead of silently retargeting it', async () => {
  const bridge = enable();
  const other = 'https://tenant.sharepoint-mil.us/sites/Other';
  expect(resolveWebUrl({ firepit: { sharepoint: bridge }, location: { search: '', hash: '#webUrl=' + encodeURIComponent(other) } })).toBe(other);
  await expect(spGet(other, '/_api/web')).rejects.toThrow('data destination was not changed');
  expect(bridge.fetch).not.toHaveBeenCalled();
});

it('sends REST reads and writes through the parent without requesting or forwarding digests', async () => {
  const bridge = enable({ Id: 3 }), direct = vi.fn();
  await spGet(web, '/_api/web', direct);
  await spPost(web, '/_api/web/lists/getbytitle(\'Tasks\')/items', { body: { Title: 'Saved' }, headers: { 'X-RequestDigest': 'stale', Authorization: 'unused' } }, direct);
  expect(bridge.fetch).toHaveBeenCalledTimes(2);
  expect(direct).not.toHaveBeenCalled();
  const [url, options] = bridge.fetch.mock.calls[1];
  expect(url).toBe(web + "/_api/web/lists/getbytitle('Tasks')/items");
  expect(options.headers.has('X-RequestDigest')).toBe(false);
  expect(options.headers.has('Authorization')).toBe(false);
  expect(JSON.parse(options.body)).toEqual({ Title: 'Saved' });
});

it('retains HTTP errors and bridge policy rejections without native retries', async () => {
  const bridge = enable(), direct = vi.fn();
  bridge.fetch.mockRejectedValueOnce(new Error('SharePoint requests are disabled while Lock Down Mode is on.'));
  await expect(spGet(web, '/_api/web', direct)).rejects.toThrow('Lock Down Mode');
  bridge.fetch.mockResolvedValueOnce(new Response(JSON.stringify({ error: { message: 'Firepit session is read-only.' } }), { status: 403 }));
  await expect(spGet(web, '/_api/web', direct)).rejects.toThrow('hosting policy');
  expect(direct).not.toHaveBeenCalled();
});

it('bounds a stalled bridge and preserves caller cancellation without retrying writes', async () => {
  vi.useFakeTimers();
  try {
    const bridge = enable(), direct = vi.fn();
    bridge.fetch.mockImplementation(() => new Promise(() => {}));
    const transport = withFirepitTransport(web, direct);
    const pending = transport(web + '/_api/web/lists', { method: 'POST', body: '{}' });
    const failure = expect(pending).rejects.toThrow('operation may have completed');
    await vi.advanceTimersByTimeAsync(60000); await failure;
    expect(bridge.fetch.mock.calls[0][1].signal.aborted).toBe(true);
    const controller = new AbortController();
    controller.abort(new DOMException('User cancelled', 'AbortError'));
    await expect(transport(web + '/_api/web', { signal: controller.signal })).rejects.toThrow('User cancelled');
    expect(bridge.fetch).toHaveBeenCalledTimes(1); expect(direct).not.toHaveBeenCalled();
  } finally { vi.useRealTimers(); }
});

it('surfaces a rejected write body without a native retry or changed payload', async () => {
  const bridge = enable(), direct = vi.fn();
  bridge.fetch.mockResolvedValueOnce(new Response(JSON.stringify({ error: { message: { value: 'Firepit rejected request body: bNewDocumentUpdate must be false.' } } }), { status: 400 }));
  await expect(spPost(web, '/_api/web/lists/getbytitle(\'Tasks\')/items(1)/ValidateUpdateListItem', { body: { bNewDocumentUpdate: true, formValues: [] } }, direct)).rejects.toThrow('bNewDocumentUpdate must be false');
  expect(bridge.fetch).toHaveBeenCalledTimes(1); expect(direct).not.toHaveBeenCalled();
  expect(JSON.parse(bridge.fetch.mock.calls[0][1].body)).toEqual({ bNewDocumentUpdate: true, formValues: [] });
});

it('does not send an absolute request to a different REST service', async () => {
  const bridge = enable(), transport = withFirepitTransport(web, vi.fn());
  await expect(transport('https://tenant.sharepoint-mil.us/sites/Other/_api/web')).rejects.toThrow('outside the configured');
  await expect(transport('https://other.sharepoint-mil.us/sites/Tracker/Subweb/_api/web')).rejects.toThrow('outside the configured');
  expect(bridge.fetch).not.toHaveBeenCalled();
});

it('canonicalizes the exact MCAS alias while preserving Request bytes and group writes', async () => {
  const bridge = enable(), transport = withFirepitTransport(web, vi.fn());
  const request = new Request(web.replace('.us/', '.us.mcas-gov.us/') + '/_api/web/sitegroups(7)/users', { method: 'POST', body: '{"LoginName":"person"}' });
  await transport(request);
  const [sent] = bridge.fetch.mock.calls[0];
  expect(sent.url).toBe(web + '/_api/web/sitegroups/getbyid(7)/users');
  expect(await sent.text()).toBe('{"LoginName":"person"}');
  expect(request.bodyUsed).toBe(false);
});

it('downloads bytes through the file REST endpoint without losing punctuation', async () => {
  const bridge = enable();
  bridge.fetch.mockResolvedValueOnce(new Response(new Uint8Array([0, 1, 255])));
  const store = new SharePointStore({ webUrl: web });
  store.listTaskAttachments = vi.fn(async () => [{ name: "O'Brien #1.pdf", url: web + "/Docs/O'Brien%20%231.pdf" }]);
  const blob = await store.downloadTaskAttachment({ spId: 1 }, "O'Brien #1.pdf", { readBytes: true });
  const buffer = await new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(reader.result); reader.onerror = reject; reader.readAsArrayBuffer(blob); });
  expect([...new Uint8Array(buffer)]).toEqual([0, 1, 255]);
  const target = new URL(bridge.fetch.mock.calls[0][0]);
  expect(target.pathname).toBe('/sites/Tracker/Subweb/_api/web/GetFileByServerRelativePath(decodedurl=@file)/$value');
  expect(target.searchParams.get('@file')).toBe("'/sites/Tracker/Subweb/Docs/O''Brien #1.pdf'");
  expect(() => fileDownloadUrl(web, 'https://outside.test/file.pdf')).toThrow('outside');
  const proxyWeb = web.replace('.us/', '.us.mcas-gov.us/');
  expect(fileDownloadUrl(proxyWeb, web + '/Docs/file.pdf')).toContain('/_api/web/GetFileByServerRelativePath');
  expect(fileDownloadUrl(web, proxyWeb + '/Docs/file.pdf')).toContain('/_api/web/GetFileByServerRelativePath');
});

it('keeps form field writes within Firepit versioning policy', async () => {
  enable();
  const store = new SharePointStore({ webUrl: web });
  store.post = vi.fn(async () => ({ value: [] }));
  await store.update('tasks', 2, { TaskTitle: 'Changed' });
  expect(store.post.mock.calls[0][1].body).toMatchObject({ bNewDocumentUpdate: false });
});

it('blocks unavailable invitations and live migration before access changes or foreign reads', async () => {
  const bridge = enable(), store = new SharePointStore({ webUrl: web });
  store.get = vi.fn(); store.post = vi.fn();
  await expect(store.shareSiteAccess({}, 'Viewer', '')).rejects.toThrow('Share dialog');
  await expect(previewMigration({}, {})).rejects.toThrow('Export a tracker backup');
  expect(store.get).not.toHaveBeenCalled(); expect(store.post).not.toHaveBeenCalled();
  expect(bridge.fetch).not.toHaveBeenCalled();
});
