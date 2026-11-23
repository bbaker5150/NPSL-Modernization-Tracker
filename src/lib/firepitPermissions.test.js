import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { TrackerPermissions } from './trackerPermissions';

const webUrl = 'https://tenant.sharepoint-mil.us/sites/tracker';
beforeEach(() => { window.firepit = { sharepoint: { siteUrl: webUrl, fetch: vi.fn() } }; });
afterEach(() => { delete window.firepit; });

function folderFixture({ exists = false, projectKey = null } = {}) {
  const root = '/sites/tracker/Lists/ModernizationTasks', path = `${root}/tracker-project-5`;
  const row = { Id: 22, FileSystemObjectType: 1, FileRef: path, ProjectKey: projectKey, 'odata.etag': '"2"' };
  const folderInfo = { ServerRelativeUrl: path, ItemCount: 0 };
  const store = { webUrl, prefix: 'Modernization', formValues: vi.fn() };
  store.get = vi.fn(async url => {
    if (url.includes('/items?')) return { value: exists ? [{ Id: 22 }] : [] };
    if (url.includes('GetFolderByServerRelativePath')) return { ...folderInfo };
    if (url.includes("getbytitle('ModernizationProjects')/items(5)")) return { Id: 5, ProjectKey: 'p' };
    if (url.includes('/items(22)?')) return { ...row };
    throw new Error(`Unexpected GET: ${url}`);
  });
  store.post = vi.fn(async (url, options) => {
    if (url.endsWith("/RootFolder/Folders/add(url='tracker-project-5')")) { exists = true; return {}; }
    if (url.endsWith('/items(22)')) {
      expect(options.headers).toEqual({ 'IF-MATCH': '"2"', 'X-HTTP-Method': 'MERGE' });
      Object.assign(row, options.body); return {};
    }
    throw new Error(`Unexpected POST: ${url}`);
  });
  const permissions = new TrackerPermissions(store);
  permissions.metadata = vi.fn().mockResolvedValue({ EnableFolderCreation: true, RootFolder: { ServerRelativeUrl: root } });
  return { store, permissions, row, folderInfo, path, project: { spId: 5, projectKey: 'p' } };
}

describe('Firepit project folder creation', () => {
  it('creates a folder through the supported collection route and verifies identity before setting metadata', async () => {
    const { store, permissions, project, row, path } = folderFixture();
    await expect(permissions.folder('tasks', project, true)).resolves.toEqual({ path, scope: "/_api/web/lists/getbytitle('ModernizationTasks')/items(22)" });
    expect(store.post.mock.calls.map(([url]) => url)).toEqual([
      "/_api/web/lists/getbytitle('ModernizationTasks')/RootFolder/Folders/add(url='tracker-project-5')",
      "/_api/web/lists/getbytitle('ModernizationTasks')/items(22)",
    ]);
    expect(row).toMatchObject({ Title: 'tracker-project-5', ProjectKey: 'p' });
    expect(store.get).toHaveBeenCalledWith(expect.stringContaining("GetFolderByServerRelativePath(decodedurl='%2Fsites%2Ftracker%2FLists%2FModernizationTasks%2Ftracker-project-5')?$select=ServerRelativeUrl,ItemCount"), undefined);
    store.post.mockClear();
    await permissions.folder('tasks', project, true);
    expect(store.post).not.toHaveBeenCalled();
  });

  it.each([{ ProjectKey: 'another-project' }, { FileRef: '/sites/tracker/another-folder' }, { FileSystemObjectType: 0 }, { Id: 99 }])('preserves conflicting folders without metadata or permission writes: %j', async changes => {
    const { store, permissions, project, row } = folderFixture({ exists: true });
    Object.assign(row, changes);
    await expect(permissions.folder('tasks', project, true)).rejects.toThrow('Project folder identity');
    expect(store.post).not.toHaveBeenCalled();
  });

  it.each([{ ItemCount: 1 }, { ItemCount: undefined }, { ServerRelativeUrl: '/wrong' }])('does not adopt an existing blank-key folder with unverified emptiness: %j', async changes => {
    const { store, permissions, project, folderInfo } = folderFixture({ exists: true });
    Object.assign(folderInfo, changes);
    await expect(permissions.folder('tasks', project, true)).rejects.toThrow('Project folder identity');
    expect(store.post).not.toHaveBeenCalled();
  });

  it('requires an exact ETag and never retries a conflicting folder update', async () => {
    const { store, permissions, project } = folderFixture({ exists: true });
    store.post.mockRejectedValue(Object.assign(new Error('Concurrent update'), { status: 412 }));
    await expect(permissions.folder('tasks', project, true)).rejects.toMatchObject({ status: 412 });
    expect(store.post).toHaveBeenCalledTimes(1);
    expect(store.post.mock.calls[0][1].headers['IF-MATCH']).toBe('"2"');
  });

  it('does not create a replacement after an unavailable lookup or replay an uncertain create', async () => {
    const { store, permissions, project } = folderFixture();
    store.get.mockRejectedValueOnce(Object.assign(new Error('Denied'), { status: 403 }));
    await expect(permissions.folder('tasks', project, true)).rejects.toMatchObject({ status: 403 });
    expect(store.post).not.toHaveBeenCalled();
    store.post.mockResolvedValue({});
    await expect(permissions.folder('tasks', project, true)).rejects.toThrow('may have been created');
    expect(store.post).toHaveBeenCalledTimes(1);
  });

  it('does not send folder creation through the ordinary list row form endpoint', async () => {
    const { store, permissions } = folderFixture();
    await expect(permissions.addInFolder('tasks', '/folder', {}, 'tracker-project-5')).rejects.toThrow('Use Apply tracker permissions');
    expect(store.post).not.toHaveBeenCalled();
  });
});

function glossaryFixture(role) {
  const scope = "/_api/web/lists/getbytitle('ModernizationAcronyms')";
  const roles = [{ Id: 3, RoleTypeKind: 3, BasePermissions: { Low: '200719', High: '32' } }, ...(role ? [role] : [])];
  const store = { webUrl, prefix: 'Modernization', currentUser: vi.fn().mockResolvedValue({}), listApi: vi.fn().mockResolvedValue(scope) };
  store.get = vi.fn(async path => {
    if (path.includes('/roledefinitions?')) return { value: roles };
    if (path.includes('HasUniqueRoleAssignments')) return { HasUniqueRoleAssignments: false, ReadSecurity: 1, WriteSecurity: 1 };
    throw new Error(`Unexpected GET: ${path}`);
  });
  store.post = vi.fn().mockResolvedValue({});
  const permissions = new TrackerPermissions(store);
  permissions.currentRole = vi.fn().mockResolvedValue({ siteOwner: true });
  permissions.groupByName = vi.fn().mockResolvedValue({ Id: 20 });
  permissions.group = vi.fn(async name => ({ Id: name === 'Manager' ? 22 : 21 }));
  const rows = [];
  permissions.assignments = vi.fn(async () => rows);
  store.post.mockImplementation(async path => {
    const grant = /addroleassignment\(principalid=(\d+),roledefid=(\d+)\)/.exec(path);
    if (grant) rows.push({ PrincipalId: Number(grant[1]), Member: { PrincipalType: 8 }, RoleDefinitionBindings: [{ Id: Number(grant[2]) }] });
    return {};
  });
  return { store, permissions };
}

describe('Firepit glossary permission level preflight', () => {
  it('stops before changing ACLs if an owner must preprovision the custom permission level', async () => {
    const { store, permissions } = glossaryFixture();
    await expect(permissions.prepareGlossaryAccess()).rejects.toThrow('site owner must create the ModernizationAcronyms Add Only permission level');
    expect(store.post).not.toHaveBeenCalled();
  });
  it('reuses only the exact existing read-and-add level', async () => {
    const { store, permissions } = glossaryFixture({ Id: 26, Name: 'ModernizationAcronyms Add Only', BasePermissions: { Low: '200707', High: '32' } });
    await permissions.prepareGlossaryAccess();
    expect(store.post.mock.calls.some(([url]) => url.includes('/roledefinitions'))).toBe(false);
    expect(store.post.mock.calls.filter(([url]) => url.includes('roledefid=26'))).toHaveLength(2);
  });
  it('does not replace or use a same-name permission level with broader permissions', async () => {
    const { store, permissions } = glossaryFixture({ Id: 26, Name: 'ModernizationAcronyms Add Only', BasePermissions: { Low: '200719', High: '32' } });
    await expect(permissions.prepareGlossaryAccess()).rejects.toThrow('did not verify its read and add permissions');
    expect(store.post).not.toHaveBeenCalled();
  });
});
