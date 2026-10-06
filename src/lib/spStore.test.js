import { describe, expect, it, vi } from 'vitest';
import { CONTAINERS, SharePointStore } from './spStore';

describe('SharePoint store', () => {
  it('provisions immutable user identity keys for work ownership', () => {
    const projects = CONTAINERS.find((container) => container.key === 'projects');
    const tasks = CONTAINERS.find((container) => container.key === 'tasks');
    const updates = CONTAINERS.find((container) => container.key === 'updates');
    expect(projects.fields.map((field) => field.name)).toContain('OwnerKey');
    expect(projects.fields.map((field) => field.name)).toEqual(expect.arrayContaining(['Organization', 'OwnerCanEdit']));
    expect(tasks.fields.map((field) => field.name)).toContain('OwnerKey');
    expect(updates.fields.map((field) => field.name)).toContain('AuthorKey');
    expect(CONTAINERS.find((container) => container.key === 'acronyms').fields.map((field) => field.name)).toEqual(['RecordId', 'Acronym', 'FullTerm', 'Definition', 'SeedVersion', 'Archived']);
  });

  it('follows SharePoint pagination links so exports include every item', async () => {
    const calls = [];
    const fetchImpl = async (url) => {
      calls.push(url);
      if (url.includes("/lists?")) return new Response(JSON.stringify({ value: [{ Id: "11111111-1111-1111-1111-111111111111", Title: "Modernization-Tracker - Projects", BaseTemplate: 100, Hidden: false }] }));
      const second = url.includes('$skiptoken');
      return new Response(JSON.stringify(second
        ? { value: [{ Id: 2, Title: 'Second' }] }
        : { value: [{ Id: 1, Title: 'First' }], '@odata.nextLink': 'https://tenant.sharepoint.com/sites/mod/_api/list?$skiptoken=2' }),
      { status: 200, headers: { 'Content-Type': 'application/json' } });
    };
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod', fetchImpl });
    const rows = await store.listItems('projects', [], (item) => item.Title);
    expect(rows).toEqual(['First', 'Second']);
    expect(calls).toHaveLength(3);
    expect(calls[2]).toContain('/sites/mod/_api/list?$skiptoken=2');
  });

  it('routes existing lists with missing identity fields through additive setup', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.listExists = async () => true;
    store.get = async () => ({ value: [{ InternalName: 'RecordId' }, { InternalName: 'ProjectKey' }] });
    const readiness = await store.readiness();
    expect(readiness.ready).toBe(false);
    expect(readiness.checks.every((check) => check.exists && check.missingFields.length > 0)).toBe(true);
  });

  it('creates visible, app-prefixed lists without follow-up MERGE operations', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.listExists = async () => false;
    store.get = vi.fn(async () => ({ value: CONTAINERS[0].fields.map((field) => ({ InternalName: field.name })) }));
    store.post = vi.fn(async () => ({}));

    await store.provision();

    const listCreates = store.post.mock.calls.filter(([path]) => path === '/_api/web/lists');
    expect(listCreates).toHaveLength(CONTAINERS.length);
    expect(listCreates.every(([, options]) => options.body.Hidden === false)).toBe(true);
    expect(listCreates.map(([, { body }]) => body.Title)).toEqual(['Modernization-Tracker - Projects', 'Modernization-Tracker - Tasks', 'Modernization-Tracker - Updates', 'Modernization-Tracker - Risks', 'Modernization-Tracker - Acronyms', 'Modernization-Tracker - Users', 'Modernization-Tracker - Reference Documents']);
    expect(JSON.stringify(store.post.mock.calls)).not.toContain('X-HTTP-Method');
  });

  it('updates date fields through prompt-free validation posts using SharePoint form values', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.post = vi.fn(async () => ({ value: [] }));
    const task = {
      spId: 42, id: 'task-42', projectKey: 'project', wbs: '1.1', title: 'Review task', phaseKey: 'requirement',
      order: 1, status: 'In Progress', estimatedHours: 2.5, assignedDate: '2026-08-25', startDate: '2026-09-01', dueDate: '2026-09-30', finishDate: '',
      ownerName: 'Engineer', ownerEmail: 'engineer@example.invalid', ownerKey: '', notes: '', blockedReason: '', sourceStartLabel: '', dataIssue: '',
    };
    await store.saveTask(task);
    expect(store.post.mock.calls[0][0]).toContain('/items(42)/validateupdatelistitem');
    const values = Object.fromEntries(store.post.mock.calls[0][1].body.formValues.map((field) => [field.FieldName, field.FieldValue]));
    expect(values).toMatchObject({
      EstimatedHours: '2.5', AssignedDate: '8/25/2026',
      StartDate: '9/1/2026',
      DueDate: '9/30/2026',
      FinishDate: '',
    });
    expect(JSON.stringify(store.post.mock.calls[0])).not.toContain('X-HTTP-Method');
  });

  it('persists project organization and manager-controlled owner editing', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.post = vi.fn(async () => ({ value: [] }));
    await store.saveProject({ spId: 4, id: 'p4', projectKey: 'p4', title: 'Modernize', organization: 'ISE', ownerCanEdit: true });
    const values = Object.fromEntries(store.post.mock.calls[0][1].body.formValues.map((field) => [field.FieldName, field.FieldValue]));
    expect(values).toMatchObject({ Organization: 'NPSL Metrology Engineering', OwnerCanEdit: '1' });
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

describe('SharePoint user role verification', () => {
  const row = { spId: 7, id: 'u7', title: 'Engineer', loginName: 'i:0#.f|membership|engineer@example.test', email: 'engineer@example.test', role: 'Manager' };
  it.each(['Manager', 'Viewer'])('reads the saved SharePoint role back before confirming %s', async (role) => {
    const requested = { ...row, role };
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.post = vi.fn(async () => ({ value: [{ FieldName: 'AppRole', HasException: false, ErrorMessage: null }] }));
    store.get = vi.fn(async () => ({ Id: 7, RecordId: row.id, Title: row.title, LoginKey: row.loginName, Email: row.email, AppRole: role }));
    expect(await store.saveUser(requested)).toEqual(requested);
    expect(store.post.mock.calls[0][1].body.formValues).toContainEqual({ FieldName: 'AppRole', FieldValue: role });
    expect(store.get.mock.calls[0][0]).toContain('/items(7)?$select=');
  });
  it('rejects a successful HTTP response when the stored role is still User', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.post = vi.fn(async () => ({ value: [] }));
    store.get = vi.fn(async () => ({ Id: 7, RecordId: row.id, LoginKey: row.loginName, AppRole: 'Viewer' }));
    await expect(store.saveUser(row)).rejects.toThrow('Current stored role: Viewer');
  });
  it('surfaces SharePoint field validation errors instead of claiming success', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.post = vi.fn(async () => ({ ValidateUpdateListItem: [{ FieldName: 'AppRole', HasException: true, ErrorMessage: 'Access denied' }] }));
    await expect(store.saveUser(row)).rejects.toThrow('Access denied');
  });
});

describe('native SharePoint task attachments', () => {
  it('uploads raw bytes to the persisted task and reads back its attachment list', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.archivedDocuments = vi.fn(async () => []);
    store.get = vi.fn().mockResolvedValueOnce({ value: [] }).mockResolvedValueOnce({ EnableAttachments: true }).mockResolvedValueOnce({ value: [{ FileName: "Engineer’s report.pdf", ServerRelativeUrl: '/sites/mod/Lists/ModernizationTasks/Attachments/42/report.pdf' }] });
    store.post = vi.fn(async () => ({}));
    const bytes = new Uint8Array([0, 255, 42]).buffer;
    const file = { name: "Engineer's report.pdf", size: 3, arrayBuffer: async () => bytes };
    const files = await store.addTaskAttachment({ spId: 42 }, file);
    expect(store.post.mock.calls[0][0]).toContain("/items(42)/AttachmentFiles/add(FileName='Engineer''s%20report.pdf')");
    expect(store.post.mock.calls[0][1]).toMatchObject({ raw: true, body: bytes, headers: { 'Content-Type': 'application/octet-stream' } });
    expect(files[0].url).toBe('https://tenant.sharepoint.com/sites/mod/Lists/ModernizationTasks/Attachments/42/report.pdf');
  });
  it('rejects duplicate names, invalid files, and unsaved tasks before posting', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.get = vi.fn(async () => ({ value: [{ FileName: 'REPORT.pdf', ServerRelativeUrl: '/report.pdf' }] }));
    store.post = vi.fn();
    await expect(store.addTaskAttachment({ spId: 42 }, { name: 'report.pdf', size: 1 })).rejects.toThrow('already attached');
    await expect(store.addTaskAttachment({ spId: 42 }, { name: 'large.pdf', size: 51 * 1024 * 1024 })).rejects.toThrow('50 MB');
    await expect(store.listTaskAttachments({})).rejects.toThrow('Save the task');
    expect(store.post).not.toHaveBeenCalled();
  });
});


describe('prompt-free SharePoint removal and downloads', () => {
  it('renames an attachment by copying its bytes and archiving the previous name', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    const blob = new Blob(['report'], { type: 'application/pdf' });
    store.listTaskAttachments = vi.fn(async () => [{ name: 'old.pdf' }]);
    store.downloadTaskAttachment = vi.fn(async () => blob);
    store.addTaskAttachment = vi.fn(async () => []);
    store.deleteTaskAttachment = vi.fn(async () => undefined);
    await store.renameTaskAttachment({ spId: 9 }, 'old.pdf', 'new.pdf');
    expect(store.addTaskAttachment).toHaveBeenCalledWith({ spId: 9 }, expect.objectContaining({ name: 'new.pdf', size: blob.size }));
    expect(store.deleteTaskAttachment).toHaveBeenCalledWith({ spId: 9 }, 'old.pdf');
  });
  it('archives an existing attachment through a verified metadata update', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.listTaskAttachments = vi.fn(async () => [{ name: 'report.pdf' }]);
    store.archivedDocuments = vi.fn().mockResolvedValueOnce(['old.pdf']).mockResolvedValueOnce(['old.pdf', 'report.pdf']);
    store.update = vi.fn();
    await store.deleteTaskAttachment({ spId: 9 }, 'report.pdf');
    expect(store.update).toHaveBeenCalledWith('tasks', 9, { ArchivedDocuments: '["old.pdf","report.pdf"]' });
    store.archivedDocuments.mockResolvedValue([]);
    await expect(store.deleteTaskAttachment({ spId: 9 }, 'report.pdf')).rejects.toThrow('did not confirm');
    await expect(store.deleteTaskAttachment({ spId: 9 }, 'missing.pdf')).rejects.toThrow('no longer exists');
  });
  it('filters archived records and documents without delete or recycle requests', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.get = vi.fn(async () => ({ value: [{ Id: 1, Archived: true }, { Id: 2, Archived: false }] }));
    expect(await store.listItems('tasks', ['Archived'], (row) => row.Id)).toEqual([2]);
    store.post = vi.fn(async () => ({ value: [] }));
    await store.recycle('tasks', 1);
    expect(store.post.mock.calls[0][0]).toContain('/validateupdatelistitem');
    expect(store.post.mock.calls[0][1].body.formValues).toContainEqual({ FieldName: 'Archived', FieldValue: '1' });
    store.get = vi.fn(async () => ({ value: [{ FileName: 'deleted.pdf', ServerRelativeUrl: '/deleted.pdf' }, { FileName: 'active.pdf', ServerRelativeUrl: '/active.pdf' }] }));
    store.archivedDocuments = vi.fn(async () => ['deleted.pdf']);
    expect(await store.listTaskAttachments({ spId: 1 })).toEqual([{ name: 'active.pdf', url: 'https://tenant.sharepoint.com/active.pdf' }]);
  });
  it('downloads only a listed file as binary data', async () => {
    const blob = new Blob(['contents']);
    const fetchImpl = vi.fn(async () => ({ ok: true, blob: async () => blob }));
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod', fetchImpl });
    store.listTaskAttachments = vi.fn(async () => [{ name: 'report.pdf', url: 'https://tenant.sharepoint.com/report.pdf' }]);
    expect(await store.downloadTaskAttachment({ spId: 1 }, 'report.pdf')).toBe(blob);
    expect(fetchImpl).toHaveBeenCalledWith('https://tenant.sharepoint.com/report.pdf', { credentials: 'include', headers: { Accept: '*/*' } });
    await expect(store.downloadTaskAttachment({ spId: 1 }, 'missing.pdf')).rejects.toThrow('no longer exists');
  });
});


describe('SharePoint people picker', () => {
  it('searches resolved individual users and resolves the selected identity before saving', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/mod' });
    store.post = vi.fn().mockResolvedValueOnce({ d: { ClientPeoplePickerSearchUser: JSON.stringify([
      { Key: 'claims|person', DisplayText: 'Person', IsResolved: true, EntityType: 'User', EntityData: { Email: 'person@example.test' } },
      { Key: 'group', IsResolved: true, EntityType: 'SecGroup' },
      { Key: 'unresolved', IsResolved: false, EntityType: 'User' },
    ]) } }).mockResolvedValueOnce({ d: { Id: 3, Title: 'Person', Email: 'person@example.test', LoginName: 'claims|person' } });
    expect(await store.searchPeople('Person')).toEqual([{ title: 'Person', email: 'person@example.test', loginName: 'claims|person', detail: '' }]);
    expect(store.post.mock.calls[0][1].body.queryParams).toMatchObject({ PrincipalType: 1, QueryString: 'Person', MaximumEntitySuggestions: 15 });
    expect(await store.resolvePerson('claims|person')).toEqual({ title: 'Person', email: 'person@example.test', loginName: 'claims|person' });
    expect(store.post.mock.calls[1]).toEqual(['/_api/web/ensureuser', { body: { logonName: 'claims|person' } }]);
  });
});


describe('page-targeted invitations', () => {
  const site = 'https://tenant.sharepoint.com/sites/mod';
  const pageUrl = `${site}/SitePages/Tracker.aspx`;
  function setup(low = 7, pageLow = 33, level = 1) {
    const store = new SharePointStore({ webUrl: site });
    let member = false;
    store.get = vi.fn(async (path) => {
      if (path.includes('associatedmembergroup')) return { Id: 9, Title: 'ISEA METENG Members' };
      if (path.includes('sitegroups(9)/users')) return { value: member ? [{ LoginName: 'claims|person' }] : [] };
      if (path.includes('getusereffectivepermissions')) return { GetUserEffectivePermissions: { Low: String(path.includes('ListItemAllFields') ? pageLow : low) } };
      if (path.includes('roledefinitions')) return { value: [{ Id: 124, RoleTypeKind: 2 }] };
      return { Exists: true, Level: level };
    });
    store.post = vi.fn(async (path) => {
      if (path.includes('sitegroups(9)/users')) member = true;
      return { StatusCode: 0 };
    });
    return store;
  }
  it.each(['Viewer', 'Project Engineer', 'Manager'])('adds %s to Members, verifies access, then makes exactly one invitation request', async (role) => {
    const store = setup();
    expect(await store.shareSiteAccess({ loginName: 'claims|person', email: 'person@example.test' }, role, pageUrl)).toMatchObject({ access: 'ISEA METENG Members' });
    expect(store.post.mock.calls[0]).toEqual(['/_api/web/sitegroups(9)/users', { body: { LoginName: 'claims|person' } }]);
    expect(store.post.mock.calls.some(([path]) => path.includes('Utility.SendEmail'))).toBe(false);
    const requests = store.post.mock.calls.filter(([path]) => path.includes('ShareObject'));
    expect(requests).toHaveLength(1);
    expect(requests[0]).toEqual(['/_api/SP.Web.ShareObject', { body: expect.objectContaining({
      url: pageUrl,
      peoplePickerInput: JSON.stringify([{ Key: 'claims|person' }]),
      roleValue: 'role:124',
      groupId: 0, propagateAcl: false, sendEmail: true,
      includeAnonymousLinkInEmail: false,
      emailBody: `You have been added to the NPSL Modernization Tracker. Assigned role: ${role.toUpperCase()}.`,
    }) }]);
    expect(requests[0][1].body.emailBody).not.toMatch(/<[^>]+>|https?:\/\//);
    expect(store.get.mock.calls.some(([path]) => path.includes('ListItemAllFields/getusereffectivepermissions'))).toBe(true);
  });
  it('does not add existing members again on retry', async () => {
    const store = setup();
    await store.shareSiteAccess({ loginName: 'claims|person', email: 'person@example.test' }, 'Viewer', pageUrl);
    await store.shareSiteAccess({ loginName: 'claims|person', email: 'person@example.test' }, 'Viewer', pageUrl);
    expect(store.post.mock.calls.filter(([path]) => path.includes('sitegroups'))).toHaveLength(1);
  });
  it('blocks email when membership is missing or cannot be confirmed', async () => {
    const store = setup();
    store.post.mockResolvedValue({ StatusCode: 0 });
    await expect(store.shareSiteAccess({ loginName: 'claims|person' }, 'Viewer', pageUrl)).rejects.toThrow('did not confirm group membership');
    expect(store.post.mock.calls.some(([path]) => path.includes('ShareObject'))).toBe(false);
    const get = store.get;
    store.get = vi.fn((path) => path.includes('associatedmembergroup') ? {} : get(path));
    await expect(store.shareSiteAccess({ loginName: 'claims|person' }, 'Viewer', pageUrl)).rejects.toThrow('Members group is unavailable');
  });
  it('does not share or email an unpublished page or invalid library link', async () => {
    const store = setup(7, 1, 2);
    await expect(store.shareSiteAccess({ loginName: 'claims|person' }, 'Viewer', pageUrl)).rejects.toThrow('Publish or republish');
    await expect(store.shareSiteAccess({ loginName: 'person' }, 'Viewer', `${site}/SitePages/Forms/ByAuthor.aspx`)).rejects.toThrow('direct published tracker page');
    expect(store.post).not.toHaveBeenCalled();
  });
  it('never sends an invitation when page access verification fails and reports invitation failures', async () => {
    const store = setup(7, 0);
    await expect(store.shareSiteAccess({ loginName: 'claims|person' }, 'Viewer', pageUrl)).rejects.toThrow('tracker page access could not be verified');
    expect(store.post.mock.calls.some(([path]) => path.includes('ShareObject'))).toBe(false);
    const emailFailure = setup();
    const post = emailFailure.post.getMockImplementation();
    emailFailure.post.mockImplementation(async (path) => {
      if (!path.includes('ShareObject')) return post(path);
      throw new Error('Email denied');
    });
    await expect(emailFailure.shareSiteAccess({ loginName: 'claims|person', email: 'person@example.test' }, 'Viewer', pageUrl)).rejects.toThrow('Email denied');
  });
});
