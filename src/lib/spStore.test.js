import { describe, expect, it, vi } from 'vitest';
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

  it('hides backing lists from Site Contents while preserving REST access', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.listExists = async () => true;
    store.get = vi.fn(async (path) => path.includes('/fields')
      ? { value: CONTAINERS[0].fields.map((field) => ({ InternalName: field.name })) }
      : { Hidden: false });
    store.post = vi.fn(async () => ({}));

    await store.provision();

    const visibilityWrites = store.post.mock.calls.filter(([path, options]) => !path.includes('/fields') && options?.body?.Hidden === true);
    expect(visibilityWrites).toHaveLength(CONTAINERS.length);
    expect(visibilityWrites.every(([, options]) => options.headers['X-HTTP-Method'] === 'MERGE')).toBe(true);
  });

  it('updates date fields through REST MERGE using locale-independent ISO values', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.post = vi.fn(async () => ({}));
    const task = {
      spId: 42, id: 'task-42', projectKey: 'project', wbs: '1.1', title: 'Review task', phaseKey: 'need-scope',
      order: 1, status: 'In Progress', startDate: '2026-09-01', dueDate: '2026-09-30', finishDate: '',
      ownerName: 'Engineer', ownerEmail: 'engineer@example.invalid', ownerKey: '', notes: '', blockedReason: '', sourceStartLabel: '', dataIssue: '',
    };
    await store.saveTask(task);
    expect(store.post).toHaveBeenCalledWith(expect.stringContaining('/items(42)'), expect.objectContaining({
      headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
      body: expect.objectContaining({
        StartDate: '2026-09-01T12:00:00Z',
        DueDate: '2026-09-30T12:00:00Z',
        FinishDate: null,
      }),
    }));
    expect(store.post.mock.calls[0][0]).not.toContain('ValidateUpdateListItem');
  });
});
