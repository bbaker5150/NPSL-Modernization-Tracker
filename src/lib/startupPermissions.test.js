import { describe, expect, it, vi } from 'vitest';
import { CONTAINERS, SharePointStore } from './spStore';
import { spGet } from './spContext';

const webUrl = 'https://tenant.sharepoint-mil.us/sites/metsoft';
const json = (body, status = 200) => new Response(JSON.stringify(body), { status });
const startStorage = async store => {
  const readiness = await store.readiness();
  if (!readiness.ready) await store.provision();
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
    await expect(startStorage(store)).rejects.toThrow('required fields are unavailable');
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
