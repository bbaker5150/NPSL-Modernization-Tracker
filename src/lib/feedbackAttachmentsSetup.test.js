import { describe, it, expect, vi } from 'vitest';
import { feedbackAttachmentTitle, feedbackAttachmentsSetup } from './feedbackAttachmentsSetup';

const sourcePath = "/_api/web/lists(guid'11111111-1111-1111-1111-111111111111')";
const libraryId = '22222222-2222-2222-2222-222222222222';
const libraryPath = `/_api/web/lists(guid'${libraryId}')`;
const missing = () => Object.assign(new Error('Missing'), { status: 404 });
const clone = value => structuredClone(value);
const expectedLibraryAcl = () => new Map([[10, [1]], [20, [2]], [21, [2]], [22, [2, 4]]]);

function fixture() {
  const roles = [
    { Id: 1, RoleTypeKind: 5, BasePermissions: { Low: '4294967295', High: '2147483647' } },
    { Id: 2, BasePermissions: { Low: '200706', High: '32' } },
    { Id: 3, BasePermissions: { Low: '200711', High: '32' } },
  ];
  const state = {
    list: null, field: null, sourceAcl: new Map([[10, [1]], [20, [2]], [21, [2]], [22, [2, 3]]]), libraryAcl: new Map(),
    roles, failField: false, failCreatorRemoval: false, low: '4294967295',
  };
  const rows = grants => [...grants].map(([PrincipalId, ids]) => ({ PrincipalId, RoleDefinitionBindings: ids.map(id => clone(roles.find(role => role.Id === id))) }));
  const store = {
    webUrl: 'https://example.test/sites/tracker', currentUser: vi.fn(async () => ({ id: 8 })),
    get: vi.fn(async path => {
      if (path === '/_api/web/EffectiveBasePermissions') return { Low: state.low };
      if (path.includes('/associatedownergroup?')) return { Id: 10, Title: 'Site Owners' };
      if (path.includes('/sitegroups?')) return { value: [{ Id: 22, Title: 'Tracker Managers' }] };
      if (path.includes('/roledefinitions?')) return { value: clone(roles) };
      if (path.startsWith(`${sourcePath}/roleassignments?`)) return { value: rows(state.sourceAcl) };
      if (path.startsWith(`${sourcePath}/fields?`)) return { value: state.field ? [clone(state.field)] : [] };
      if (path.startsWith(`${libraryPath}/roleassignments?`)) return { value: rows(state.libraryAcl) };
      if (path.startsWith('/_api/web/lists/getbytitle(')) {
        const title = decodeURIComponent(path.match(/getbytitle\('([^']+)'\)/)[1]);
        const list = [state.list, state.duplicateList].find(row => row?.Title === title);
        if (!list) throw missing();
        return clone(list);
      }
      throw new Error(`Unexpected read: ${path}`);
    }),
    post: vi.fn(async (path, options = {}) => {
      if (path === '/_api/web/lists') state.list = { Id: libraryId, ...options.body, ItemCount: 0, HasUniqueRoleAssignments: false };
      else if (path === '/_api/web/roledefinitions') roles.push({ Id: roles.length + 1, ...options.body });
      else if (path === `${libraryPath}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`) {
        if (!state.list.HasUniqueRoleAssignments) state.libraryAcl = new Map([[8, [1]]]);
        state.list.HasUniqueRoleAssignments = true;
      } else if (path === `${sourcePath}/fields/createfieldasxml`) {
        if (state.failField) throw new Error('Field creation interrupted');
        state.field = { InternalName: 'FeedbackAttachments', TypeAsString: 'Note', RichText: false, AppendOnly: false };
      } else if (path === libraryPath && options.headers?.['X-HTTP-Method'] === 'MERGE') Object.assign(state.list, options.body);
      else if (path.startsWith(`${libraryPath}/roleassignments/`)) {
        const [, principalText, roleText] = path.match(/principalid=(\d+),roledefid=(\d+)/), principal = Number(principalText), role = Number(roleText);
        if (path.includes('addroleassignment')) state.libraryAcl.set(principal, [...new Set([...(state.libraryAcl.get(principal) || []), role])]);
        else if (path.includes('removeroleassignment')) {
          if (principal === 8 && state.failCreatorRemoval) throw new Error('Creator removal interrupted');
          const remaining = (state.libraryAcl.get(principal) || []).filter(value => value !== role);
          if (remaining.length) state.libraryAcl.set(principal, remaining); else state.libraryAcl.delete(principal);
        } else throw new Error(`Unexpected ACL write: ${path}`);
      } else throw new Error(`Unexpected write: ${path}`);
    }),
  };
  const isOwner = vi.fn(async () => true), verifyFeedback = vi.fn(async () => ({ ready: true }));
  const setup = feedbackAttachmentsSetup(store, isOwner, async () => sourcePath, verifyFeedback);
  return { state, store, setup, isOwner, verifyFeedback };
}

describe('private feedback attachments setup', () => {
  it('uses the tracker naming convention and discovers missing optional storage without writes', async () => {
    const { setup, store } = fixture();
    expect(feedbackAttachmentTitle()).toBe('Modernization-Tracker - Feedback Attachments');
    expect(feedbackAttachmentTitle('Custom App')).toBe('CustomApp Tracker - Feedback Attachments');
    expect(await setup.status()).toEqual({ ready: false, title: feedbackAttachmentTitle() });
    expect(store.post).not.toHaveBeenCalled();
  });

  it('clones all approved source bindings, removes creator access, and verifies the plain manifest field', async () => {
    const { setup, state, store } = fixture();
    expect(await setup.provision()).toEqual({ ready: true, title: feedbackAttachmentTitle() });
    expect(state.list).toMatchObject({ BaseTemplate: 101, NoCrawl: true, HasUniqueRoleAssignments: true, ItemCount: 0 });
    expect(state.libraryAcl).toEqual(expectedLibraryAcl());
    expect(state.roles[3]).toMatchObject({ Name: 'ModernizationFeedback Attachment Review', BasePermissions: { Low: '200743', High: '32' } });
    expect(state.roles[2].BasePermissions).toEqual({ Low: '200711', High: '32' });
    expect(state.roles[1].BasePermissions).toEqual({ Low: '200706', High: '32' });
    expect(state.sourceAcl.get(22)).toEqual([2, 3]);
    expect(state.libraryAcl.has(8)).toBe(false);
    expect(state.field).toMatchObject({ TypeAsString: 'Note', RichText: false, AppendOnly: false });
    const calls = store.post.mock.calls.map(([path]) => path);
    expect(calls.findIndex(path => path.includes('principalid=20'))).toBeGreaterThan(calls.findIndex(path => path.includes('/fields/createfieldasxml')));
    expect(calls.findIndex(path => path.includes('principalid=10'))).toBeLessThan(calls.findIndex(path => path.includes('removeroleassignment(principalid=8')));
    store.post.mockClear();
    state.list.ItemCount = 7;
    expect(await setup.provision()).toMatchObject({ ready: true });
    expect(store.post).not.toHaveBeenCalled();
  });

  it('requires a site owner and schema permissions before making any writes', async () => {
    const { setup, store, state, isOwner } = fixture();
    isOwner.mockResolvedValue(false);
    await expect(setup.status()).rejects.toThrow('Only site owners');
    await expect(setup.provision()).rejects.toThrow('Only site owners');
    isOwner.mockResolvedValue(true);
    state.low = '2048';
    await expect(setup.provision()).rejects.toThrow('Manage Lists and Manage Permissions');
    expect(store.post).not.toHaveBeenCalled();
  });

  it('requires existing verified feedback and its associated Owners Full Control grant', async () => {
    const { setup, store, state, verifyFeedback } = fixture();
    verifyFeedback.mockResolvedValue({ ready: false });
    expect(await setup.status()).toMatchObject({ ready: false });
    await expect(setup.provision()).rejects.toThrow('Finish setting up private feedback');
    verifyFeedback.mockResolvedValue({ ready: true });
    state.sourceAcl.set(10, [3]);
    await expect(setup.provision()).rejects.toThrow('associated Owners group Full Control');
    expect(store.post).not.toHaveBeenCalled();
  });

  it('does not treat an access failure as a missing library', async () => {
    const { setup, store } = fixture();
    const get = store.get.getMockImplementation();
    store.get.mockImplementation(path => path.includes('/getbytitle(') ? Promise.reject(Object.assign(new Error('Access denied'), { status: 403 })) : get(path));
    await expect(setup.status()).rejects.toThrow('Access denied');
    await expect(setup.provision()).rejects.toThrow('Access denied');
    expect(store.post).not.toHaveBeenCalled();
  });

  it('never replaces an unrelated or populated interrupted library', async () => {
    const { setup, store, state } = fixture();
    state.list = { Id: libraryId, Title: feedbackAttachmentTitle(), Description: 'Unrelated library', BaseTemplate: 101, ItemCount: 0 };
    await expect(setup.provision()).rejects.toThrow('will not overwrite');
    expect(store.post).not.toHaveBeenCalled();
    state.list = null;
    state.failField = true;
    await expect(setup.provision()).rejects.toThrow('Field creation interrupted');
    state.list.ItemCount = 1;
    store.post.mockClear();
    await expect(setup.provision()).rejects.toThrow('will not overwrite');
    expect(store.post).not.toHaveBeenCalled();
  });

  it('resumes interrupted empty app storage without creating a second library', async () => {
    const { setup, store, state } = fixture();
    state.failField = true;
    await expect(setup.provision()).rejects.toThrow('Field creation interrupted');
    expect([...state.libraryAcl.keys()].every(id => [8, 10].includes(id))).toBe(true);
    expect(await setup.status()).toMatchObject({ ready: false, partial: true });
    state.failField = false;
    expect(await setup.provision()).toMatchObject({ ready: true });
    expect(store.post.mock.calls.filter(([path]) => path === '/_api/web/lists')).toHaveLength(1);
    expect(state.libraryAcl).toEqual(expectedLibraryAcl());
  });

  it('leaves storage pending until the automatic creator grant has been removed', async () => {
    const { setup, state } = fixture();
    state.failCreatorRemoval = true;
    await expect(setup.provision()).rejects.toThrow('Creator removal interrupted');
    expect(await setup.status()).toMatchObject({ ready: false, partial: true });
    state.failCreatorRemoval = false;
    expect(await setup.provision()).toMatchObject({ ready: true });
    expect(state.libraryAcl.has(8)).toBe(false);
  });

  it('detects extra readers, missing grants, and changed permission masks without resetting existing ACLs', async () => {
    const { setup, state, store } = fixture();
    await setup.provision();
    state.list.ItemCount = 4;
    store.post.mockClear();
    state.libraryAcl.set(99, [1]);
    await expect(setup.status()).rejects.toThrow('do not exactly match');
    await expect(setup.provision()).rejects.toThrow('do not exactly match');
    state.libraryAcl.delete(99);
    state.libraryAcl.delete(20);
    await expect(setup.status()).rejects.toThrow('do not exactly match');
    state.libraryAcl.set(20, [2]);
    const get = store.get.getMockImplementation();
    store.get.mockImplementation(async path => {
      const body = await get(path);
      if (path.startsWith(`${libraryPath}/roleassignments?`)) body.value.find(row => row.PrincipalId === 20).RoleDefinitionBindings[0].BasePermissions.Low = '200707';
      return body;
    });
    await expect(setup.status()).rejects.toThrow('do not exactly match');
    expect(store.post).not.toHaveBeenCalled();
  });

  it('rejects inherited permissions, unmarked storage, and search-visible storage', async () => {
    const { setup, state, store } = fixture();
    await setup.provision();
    store.post.mockClear();
    const original = clone(state.list);
    for (const change of [{ HasUniqueRoleAssignments: false }, { Description: 'Other files' }, { NoCrawl: false }, { BaseTemplate: 100 }]) {
      Object.assign(state.list, original, change);
      await expect(setup.status()).rejects.toThrow('not verified private storage');
    }
    expect(store.post).not.toHaveBeenCalled();
  });

  it('rejects a conflicting manifest field before creating storage and repairs only a missing field on verified storage', async () => {
    const { setup, state, store } = fixture();
    state.field = { InternalName: 'FeedbackAttachments', TypeAsString: 'Text' };
    await expect(setup.provision()).rejects.toThrow('plain multiple-lines-of-text');
    expect(store.post).not.toHaveBeenCalled();
    state.field = null;
    await setup.provision();
    state.field.RichText = true;
    await expect(setup.status()).rejects.toThrow('plain multiple-lines-of-text');
    state.field = null;
    state.list.ItemCount = 3;
    store.post.mockClear();
    expect(await setup.status()).toMatchObject({ ready: false });
    expect(await setup.provision()).toMatchObject({ ready: true });
    expect(store.post).toHaveBeenCalledTimes(1);
    expect(store.post.mock.calls[0][0]).toBe(`${sourcePath}/fields/createfieldasxml`);
  });

  it('rejects unexpected permission pagination addresses without requesting them', async () => {
    const { setup, store } = fixture();
    const get = store.get.getMockImplementation();
    store.get.mockImplementation(async path => path.startsWith(`${sourcePath}/roleassignments?`) ? { value: [], '@odata.nextLink': 'https://other.example/_api/roles' } : get(path));
    await expect(setup.provision()).rejects.toThrow('Unexpected feedback setup pagination address');
    expect(store.get.mock.calls.some(([path]) => path.includes('other.example'))).toBe(false);
    expect(store.post).not.toHaveBeenCalled();
  });

  it('distinguishes an absent manifest column from malformed or failed field queries', async () => {
    const { setup, store } = fixture();
    const get = store.get.getMockImplementation();
    store.get.mockImplementation(path => path.startsWith(`${sourcePath}/fields?`) ? { unexpected: [] } : get(path));
    await expect(setup.provision()).rejects.toThrow('Could not verify the feedback attachments column');
    store.get.mockImplementation(path => path.startsWith(`${sourcePath}/fields?`) ? Promise.reject(Object.assign(new Error('Field request rejected'), { status: 400 })) : get(path));
    await expect(setup.provision()).rejects.toThrow('Field request rejected');
    expect(store.post).not.toHaveBeenCalled();
    expect(store.get.mock.calls.some(([path]) => path.includes('/fields/getbyinternalnameortitle'))).toBe(false);
  });

  it('supports local demo storage without SharePoint reads or writes', async () => {
    const { store, isOwner } = fixture();
    store.webUrl = '';
    const setup = feedbackAttachmentsSetup(store, isOwner, vi.fn(), vi.fn());
    expect(await setup.status()).toMatchObject({ ready: true });
    expect(await setup.provision()).toMatchObject({ ready: true });
    expect(store.get).not.toHaveBeenCalled();
    expect(store.post).not.toHaveBeenCalled();
  });

  it('publishes the upload title only after schema, private ACLs, and creator cleanup are verified', async () => {
    const { setup, store, state } = fixture();
    const post = store.post.getMockImplementation();
    store.post.mockImplementation(async (path, options) => {
      if (path === '/_api/web/lists') expect(options.body.Title).toBe(`${feedbackAttachmentTitle()} - Setup`);
      if (options?.body?.Title === feedbackAttachmentTitle()) {
        expect(state.list.Title).toBe(`${feedbackAttachmentTitle()} - Setup`);
        expect(state.list.HasUniqueRoleAssignments).toBe(true);
        expect(state.libraryAcl).toEqual(expectedLibraryAcl());
        expect(state.field?.TypeAsString).toBe('Note');
        expect(state.libraryAcl.has(8)).toBe(false);
      }
      await post(path, options);
      if (state.list?.Title === feedbackAttachmentTitle()) {
        expect(state.list.HasUniqueRoleAssignments).toBe(true);
        expect(state.libraryAcl).toEqual(expectedLibraryAcl());
      }
    });
    await setup.provision();
    expect(store.post.mock.calls.at(-1)).toEqual([libraryPath, expect.objectContaining({ body: { Title: feedbackAttachmentTitle(), Description: 'Private feedback files. Managed by NPSL Tracker.', Hidden: false } })]);
  });

  it('does not expose the upload title when initial privacy setup is interrupted', async () => {
    const { setup, store, state } = fixture();
    const post = store.post.getMockImplementation();
    store.post.mockImplementation((path, options) => {
      if (path.includes('/breakroleinheritance')) throw new Error('Privacy setup interrupted');
      return post(path, options);
    });
    await expect(setup.provision()).rejects.toThrow('Privacy setup interrupted');
    expect(state.list).toMatchObject({ Title: `${feedbackAttachmentTitle()} - Setup`, HasUniqueRoleAssignments: false });
    expect(await setup.status()).toMatchObject({ ready: false, partial: true });
  });

  it('requires the manager history upgrade before creating roles or storage', async () => {
    const { setup, verifyFeedback, store } = fixture();
    verifyFeedback.mockResolvedValue({ ready: true, managerReviewNeeded: true });
    await expect(setup.provision()).rejects.toThrow('Enable manager history access before');
    expect(store.post).not.toHaveBeenCalled();
  });

  it('never clones ordinary read, edit, delete, download, or administrative privileges', async () => {
    const { setup, state, store } = fixture();
    for (const extra of [1n, 4n, 8n, 32n, 33554432n]) {
      state.roles[1].BasePermissions.Low = String(200706n | extra);
      await expect(setup.provision()).rejects.toThrow('not add-only');
    }
    expect(store.post).not.toHaveBeenCalled();
  });

  it('reports a missing attachment review role without writes and rejects an unsafe replacement', async () => {
    const { setup, state, store } = fixture();
    await setup.provision();
    const reviewer = state.roles.pop();
    store.post.mockClear();
    expect(await setup.status()).toMatchObject({ ready: false });
    expect(store.post).not.toHaveBeenCalled();
    state.roles.push({ ...reviewer, BasePermissions: { Low: '200751', High: '32' } });
    await expect(setup.status()).rejects.toThrow('review permission level could not be verified');
    await expect(setup.provision()).rejects.toThrow('review permission level could not be verified');
    expect(store.post).not.toHaveBeenCalled();
  });

  it('refuses conflicting libraries and legacy pending libraries already under the upload title', async () => {
    const { setup, state, store } = fixture();
    state.failField = true;
    await expect(setup.provision()).rejects.toThrow('Field creation interrupted');
    state.list.Title = feedbackAttachmentTitle();
    store.post.mockClear();
    await expect(setup.provision()).rejects.toThrow('will not overwrite');
    expect(store.post).not.toHaveBeenCalled();
    state.duplicateList = { ...state.list, Id: '33333333-3333-3333-3333-333333333333', Title: `${feedbackAttachmentTitle()} - Setup` };
    await expect(setup.status()).rejects.toThrow('Both published and pending');
    await expect(setup.provision()).rejects.toThrow('Both published and pending');
    expect(store.post).not.toHaveBeenCalled();
  });
});
