import { describe, expect, it } from 'vitest';
import { CONTAINERS, SharePointStore } from './spStore';

describe('SharePoint store', () => {
  it('provisions immutable user identity keys for work ownership', () => {
    const projects = CONTAINERS.find((container) => container.key === 'projects');
    const tasks = CONTAINERS.find((container) => container.key === 'tasks');
    const updates = CONTAINERS.find((container) => container.key === 'updates');
    expect(projects.fields.map((field) => field.name)).toContain('OwnerKey');
    expect(tasks.fields.map((field) => field.name)).toContain('OwnerKey');
    expect(updates.fields.map((field) => field.name)).toContain('AuthorKey');
  });

  it('follows SharePoint pagination links so exports include every item', async () => {
    const calls = [];
    const fetchImpl = async (url) => {
      calls.push(url);
      const second = url.includes('$skiptoken');
      return new Response(JSON.stringify(second
        ? { value: [{ Id: 2, Title: 'Second' }] }
        : { value: [{ Id: 1, Title: 'First' }], '@odata.nextLink': 'https://tenant.sharepoint.com/sites/mod/_api/list?$skiptoken=2' }),
      { status: 200, headers: { 'Content-Type': 'application/json' } });
    };
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod', fetchImpl });
    const rows = await store.listItems('projects', [], (item) => item.Title);
    expect(rows).toEqual(['First', 'Second']);
    expect(calls).toHaveLength(2);
    expect(calls[1]).toContain('/sites/mod/_api/list?$skiptoken=2');
  });

  it('routes existing lists with missing identity fields through additive setup', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.listExists = async () => true;
    store.get = async () => ({ value: [{ InternalName: 'RecordId' }, { InternalName: 'ProjectKey' }] });
    const readiness = await store.readiness();
    expect(readiness.ready).toBe(false);
    expect(readiness.checks.every((check) => check.exists && check.missingFields.length > 0)).toBe(true);
  });
});
