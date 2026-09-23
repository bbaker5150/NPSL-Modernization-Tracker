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
    expect(CONTAINERS.find((container) => container.key === 'acronyms').fields.map((field) => field.name)).toEqual(['RecordId', 'Acronym', 'FullTerm', 'Definition', 'SeedVersion']);
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

  it('creates backing lists hidden without follow-up MERGE operations', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.listExists = async () => false;
    store.get = vi.fn(async () => ({ value: CONTAINERS[0].fields.map((field) => ({ InternalName: field.name })) }));
    store.post = vi.fn(async () => ({}));

    await store.provision();

    const listCreates = store.post.mock.calls.filter(([path]) => path === '/_api/web/lists');
    expect(listCreates).toHaveLength(CONTAINERS.length);
    expect(listCreates.every(([, options]) => options.body.Hidden === true)).toBe(true);
    expect(JSON.stringify(store.post.mock.calls)).not.toContain('X-HTTP-Method');
  });

  it('updates date fields through prompt-free validation posts using SharePoint form values', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.post = vi.fn(async () => ({ value: [] }));
    const task = {
      spId: 42, id: 'task-42', projectKey: 'project', wbs: '1.1', title: 'Review task', phaseKey: 'requirement',
      order: 1, status: 'In Progress', startDate: '2026-09-01', dueDate: '2026-09-30', finishDate: '',
      ownerName: 'Engineer', ownerEmail: 'engineer@example.invalid', ownerKey: '', notes: '', blockedReason: '', sourceStartLabel: '', dataIssue: '',
    };
    await store.saveTask(task);
    expect(store.post.mock.calls[0][0]).toContain('/items(42)/validateupdatelistitem');
    const values = Object.fromEntries(store.post.mock.calls[0][1].body.formValues.map((field) => [field.FieldName, field.FieldValue]));
    expect(values).toMatchObject({
      StartDate: '9/1/2026',
      DueDate: '9/30/2026',
      FinishDate: '',
    });
    expect(JSON.stringify(store.post.mock.calls[0])).not.toContain('X-HTTP-Method');
  });

  it('creates acronyms in the shared SharePoint glossary list', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.create = vi.fn(async () => 81);

    const saved = await store.saveAcronym({ id: 'acronym-css', acronym: 'CSS', term: 'Calibration Standard Specification', definition: 'Technical requirements.' });

    expect(store.create).toHaveBeenCalledWith('acronyms', {
      Title: 'CSS', RecordId: 'acronym-css', Acronym: 'CSS', FullTerm: 'Calibration Standard Specification', Definition: 'Technical requirements.', SeedVersion: '',
    });
    expect(saved.spId).toBe(81);
  });

  it('serializes every full-phase Not Required update with SharePoint form dates', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.post = vi.fn(async () => ({ value: [] }));
    const rows = [1, 2].map((spId) => ({
      spId, id: `task-${spId}`, projectKey: 'project', wbs: `RP-0${spId}`, title: 'Requirement task', phaseKey: 'requirement',
      order: spId, status: 'Not Required', startDate: '', dueDate: '', finishDate: '2026-09-01',
      ownerName: 'Engineer', ownerEmail: 'engineer@example.invalid', ownerKey: '', notes: '', blockedReason: '', sourceStartLabel: '', dataIssue: '',
    }));

    await store.saveTasks(rows);

    expect(store.post).toHaveBeenCalledTimes(2);
    for (const [, options] of store.post.mock.calls) {
      const values = Object.fromEntries(options.body.formValues.map((field) => [field.FieldName, field.FieldValue]));
      expect(values).toMatchObject({ TaskStatus: 'Not Required', StartDate: '', DueDate: '', FinishDate: '9/1/2026' });
      expect(JSON.stringify(values)).not.toMatch(/T12:00:00Z/);
    }
  });
});

describe('tasking schema round trip', () => {
  it('persists deferrals and exceptions with original deadlines and loads directory roles', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.post = vi.fn(async () => ({ value: [] }));
    await store.saveTask({ spId: 9, id: 't9', title: 'Review', projectKey: 'p', phaseKey: 'requirement', status: 'Not Required', dueDate: '2026-09-01', deferredDate: '2026-10-01', deferredJustification: 'Parts delayed', notRequiredJustification: 'Scope changed' });
    const values = Object.fromEntries(store.post.mock.calls[0][1].body.formValues.map((row) => [row.FieldName, row.FieldValue]));
    expect(values).toMatchObject({ Title: 'Review', DueDate: '9/1/2026', DeferredDate: '10/1/2026', DeferredJustification: 'Parts delayed', NotRequiredJustification: 'Scope changed' });
    expect(values).not.toHaveProperty('WBS');
    store.get = vi.fn(async (path) => ({ value: path.includes('ModernizationTasks') ? [{ Id: 9, RecordId: 't9', TaskTitle: 'Review', TaskStatus: 'Not Required', DueDate: '2026-09-01T12:00:00Z', DeferredDate: '2026-10-01T12:00:00Z', DeferredJustification: 'Parts delayed', NotRequiredJustification: 'Scope changed' }] : path.includes('ModernizationUsers') ? [{ Id: 1, RecordId: 'u1', Title: 'Manager', LoginKey: 'manager@example.test', AppRole: 'Manager' }] : [] }));
    const loaded = await store.load();
    expect(loaded.tasks[0]).toMatchObject({ dueDate: '2026-09-01', deferredDate: '2026-10-01', deferredJustification: 'Parts delayed', notRequiredJustification: 'Scope changed' });
    expect(loaded.users[0]).toMatchObject({ role: 'Manager', loginName: 'manager@example.test' });
    expect(store.get.mock.calls.every(([path]) => !path.includes('WBS'))).toBe(true);
  });
});
