import { describe, expect, it, vi } from 'vitest';
import { CONTAINERS, SharePointStore } from './spStore';
import { spGet } from './spContext';

const webUrl = 'https://tenant.sharepoint-mil.us/sites/metsoft';
const json = (body, status = 200) => new Response(JSON.stringify(body), { status });
const startStorage = async store => {
  const readiness = await store.readiness();
  if (!readiness.ready) await store.provision({ allowCompatibleRead: true });
};

describe('startup schema permissions', () => {
  it.each([['read', '1'], ['site Edit', '2049'], ['list-manager only', '33554432'], ['missing', undefined], ['malformed', 'bad']])(
    'never writes when a list is invisible to an account with %s site permissions', async (_, low) => {
      const fetchImpl = vi.fn(async (url, options) => {
        expect(options.method).toBe('GET');
        if (url.endsWith('/EffectiveBasePermissions')) return json({ d: { EffectiveBasePermissions: { Low: low } } });
        if (url.includes('/lists?')) return json({ value: [] }); // security-trimmed catalog
        throw new Error(`Unexpected request ${url}`);
      });
      const store = new SharePointStore({ webUrl, fetchImpl, scopedAccess: true });
      await expect(startStorage(store)).rejects.toThrow('No setup changes were made');
      expect(fetchImpl.mock.calls.every(([, options]) => options.method === 'GET')).toBe(true);
    },
  );

  it('also protects direct setup calls and missing-column repairs', async () => {
    const store = new SharePointStore({ webUrl });
    store.listExists = vi.fn(async () => true);
    store.get = vi.fn(async path => path.endsWith('/EffectiveBasePermissions') ? { Low: '1' } : { value: [] });
    store.post = vi.fn();
    await expect(startStorage(store)).rejects.toThrow('Modernization-Tracker - Updates: missing fields');
    await expect(store.provision()).rejects.toThrow('No setup changes were made');
    expect(store.post).not.toHaveBeenCalled();
  });

  it('does not require setup privileges when the schema is already readable and complete', async () => {
    const store = new SharePointStore({ webUrl });
    store.listExists = vi.fn(async () => true);
    store.get = vi.fn(async () => ({ value: [...new Set(CONTAINERS.flatMap(c => c.fields.map(f => f.name)))].map(InternalName => ({ InternalName })) }));
    store.provision = vi.fn();
    await startStorage(store);
    expect(store.provision).not.toHaveBeenCalled();
    expect(store.get.mock.calls.every(([path]) => !path.includes('EffectiveBasePermissions'))).toBe(true);
  });

  it.each([403, 404])('does not write if the setup permission check fails with HTTP %s', async status => {
    const fetchImpl = vi.fn(async () => json({ error: { message: 'Access denied' } }, status));
    const store = new SharePointStore({ webUrl, fetchImpl });
    await expect(store.provision()).rejects.toMatchObject({ status });
    expect(fetchImpl).toHaveBeenCalledTimes(1);
    expect(fetchImpl.mock.calls[0][1].method).toBe('GET');
  });

  it.each([403, 404])('keeps a useful HTTP %s diagnostic without suggesting broader permissions or a removed panel', async status => {
    const request = spGet(webUrl, '/_api/web/lists', async () => json({ error: { message: { value: 'Tenant detail' } } }, status));
    await expect(request).rejects.toMatchObject({ status, message: expect.stringContaining('Tenant detail') });
    await expect(request).rejects.not.toThrow('need Edit');
    await expect(request).rejects.not.toThrow('Storage panel');
  });
});

const auditFields = ['EditedAt', 'EditorName', 'EditorKey'];
function upgradeFixture({ missing = auditFields, low = '1', missingList } = {}) {
  const store = new SharePointStore({ webUrl });
  const schemas = Object.fromEntries(CONTAINERS.map(container => [container.key, new Set(container.fields.map(field => field.name).filter(name => container.key !== 'updates' || !missing.includes(name)))]));
  store.listExists = vi.fn(async key => key !== missingList);
  store.get = vi.fn(async path => {
    if (path.endsWith('/EffectiveBasePermissions')) return { Low: low };
    const container = CONTAINERS.find(item => path.includes(`Modernization${item.suffix}')`));
    if (!container) throw new Error(`Unexpected read: ${path}`);
    if (path.includes('/fields?')) return { value: [...schemas[container.key]].map(InternalName => ({ InternalName })) };
    if (path.includes('EnableAttachments')) return { EnableAttachments: true };
    if (path.includes('/items?')) {
      const select = new URL(path, webUrl).searchParams.get('$select').split(',');
      const unknown = select.filter(name => !['Id', 'Title', 'FileSystemObjectType'].includes(name) && !schemas[container.key].has(name));
      if (unknown.length) throw new Error(`Unknown columns: ${unknown.join(', ')}`);
      return { value: container.key === 'updates' ? [{ Id: 3, RecordId: 'u', Summary: 'Existing history', AuthorName: 'Original author', 'odata.etag': '"1"' }] : [] };
    }
    throw new Error(`Unexpected read: ${path}`);
  });
  store.post = vi.fn(async (path, options) => {
    if (path.includes('/fields/createfieldasxml')) {
      schemas.updates.add(options.body.parameters.SchemaXml.match(/\bName="([^"]+)"/)[1]);
    }
    return {};
  });
  return { store, schemas };
}
it('keeps older projects readable for Viewers without selecting or writing the new team column', async () => {
  const { store, schemas } = upgradeFixture({ missing: [] });
  schemas.projects.delete('EngineersJson');
  await startStorage(store);
  expect(store.schemaReadiness).toMatchObject({ ready: false, readable: true });
  await store.load();
  expect(store.post).not.toHaveBeenCalled();
  expect(store.get.mock.calls.filter(([path]) => path.includes('ModernizationProjects') && path.includes('/items?')).every(([path]) => !path.includes('EngineersJson'))).toBe(true);
});
it.each([auditFields, ['EditedAt'], ['EditorName', 'EditorKey']].map(missing => ({ missing })))('loads the pre-editing schema for a Viewer without setup writes ($missing)', async ({ missing }) => {
  const { store } = upgradeFixture({ missing });
  await startStorage(store);
  const data = await store.load();
  expect(data.updates[0]).toMatchObject({ summary: 'Existing history', authorName: 'Original author', editorName: '', editedAt: '' });
  expect(store.updateEditingAvailable).toBe(false);
  expect(store.post).not.toHaveBeenCalled();
  for (const [path] of store.get.mock.calls.filter(([path]) => path.includes('ModernizationUpdates') && path.includes('/items?'))) {
    const select = new URL(path, webUrl).searchParams.get('$select').split(',');
    expect(select).toContain('Archived'); expect(select).toContain('AuthorKey');
    expect(missing.every(name => !select.includes(name))).toBe(true);
  }
  // Direct setup still needs a site owner, independently of compatible reads.
  await expect(store.provision()).rejects.toThrow('No setup changes were made');
  expect(store.post).not.toHaveBeenCalled();
});
it('does not block compatible reads if optional setup-permission discovery is unavailable', async () => {
  const { store } = upgradeFixture();
  const get = store.get;
  store.get = vi.fn(path => { if (path.endsWith('/EffectiveBasePermissions')) throw new Error('Permission discovery unavailable'); return get(path); });
  await startStorage(store);
  expect((await store.load()).updates[0].summary).toBe('Existing history');
  expect(store.post).not.toHaveBeenCalled();
});
it('names missing required fields and missing lists instead of blaming all Viewer permissions', async () => {
  const core = upgradeFixture({ missing: [...auditFields, 'Archived'] }).store;
  await expect(startStorage(core)).rejects.toThrow('Modernization-Tracker - Updates: missing fields');
  expect(core.schemaReadiness.readable).toBe(false);
  expect(core.post).not.toHaveBeenCalled();
  const invisible = upgradeFixture({ missingList: 'tasks' }).store;
  await expect(startStorage(invisible)).rejects.toThrow('Modernization-Tracker - Tasks: list is missing or not visible');
  expect(invisible.post).not.toHaveBeenCalled();
});
it('still propagates denied item reads instead of presenting an empty portfolio', async () => {
  const { store } = upgradeFixture();
  await startStorage(store);
  store.get.mockRejectedValue(Object.assign(new Error('Read denied'), { status: 403 }));
  await expect(store.load()).rejects.toMatchObject({ status: 403 });
  expect(store.post).not.toHaveBeenCalled();
});
it('lets an actual site owner add only the missing edit fields and use them immediately', async () => {
  const { store } = upgradeFixture({ low: String(2048 | 33554432) });
  await startStorage(store);
  expect(store.post).toHaveBeenCalledTimes(3);
  expect(store.post.mock.calls.every(([path]) => path.includes('ModernizationUpdates') && path.endsWith('/fields/createfieldasxml'))).toBe(true);
  expect(store.updateEditingAvailable).toBe(true);
  await store.load();
  const query = store.get.mock.calls.find(([path]) => path.includes('ModernizationUpdates') && path.includes('/items?'))[0];
  expect(auditFields.every(name => new URL(query, webUrl).searchParams.get('$select').split(',').includes(name))).toBe(true);
});
it('keeps posting available but requires audit fields before editing existing history', async () => {
  const { store } = upgradeFixture();
  await startStorage(store);
  store.create = vi.fn(async () => 7);
  await store.saveUpdate({ id: 'new', projectKey: 'p', summary: 'New update', authorName: 'Engineer' });
  const fields = store.create.mock.calls[0][1];
  expect(fields).toMatchObject({ Summary: 'New update', AuthorName: 'Engineer' });
  expect(auditFields.some(name => Object.hasOwn(fields, name))).toBe(false);
  await expect(store.editUpdate({ id: 'u', spId: 3, version: '"1"', summary: 'Changed' })).rejects.toThrow('enable update editing');
  await expect(store.saveUpdate({ id: 'u', editedAt: '2026-11-03T19:00:00Z' })).rejects.toThrow('before saving edit history');
  expect(store.post).not.toHaveBeenCalled();
  await store.deleteUpdate({ id: 'u', spId: 3, version: '"1"' });
  expect(store.post).toHaveBeenCalledExactlyOnceWith(expect.stringContaining('/items(3)'), expect.objectContaining({ body: { Archived: true } }));
});
