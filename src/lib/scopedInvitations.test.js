import { describe, expect, it, vi } from 'vitest';
import { CONTAINERS, SharePointStore } from './spStore';

const webUrl = 'https://tenant.sharepoint-mil.us/sites/metsoft';
const page = `${webUrl}/SitePages/Modernization-Tracker.aspx`;
const asset = `${webUrl}/SiteAssets/Tracker/tracker.html`;
const person = { loginName: 'claims|viewer' };
function setup({ deny = '', low = '0', assets = [asset], member = true, ignoreAdd = false } = {}) {
  const store = new SharePointStore({ webUrl, scopedAccess: true, invitationAssetUrls: assets });
  store.permissions.syncGroups = vi.fn(async () => null);
  store.listApi = vi.fn(async key => `/_api/web/lists/getbytitle('${key}')`);
  store.get = vi.fn(async path => {
    if (path.startsWith('/_api/web/getusereffectivepermissions')) throw new Error('Site-wide access must not be required');
    if (path.includes("getbyname('Tracker%20Users')")) return { Id: 99, Title: 'Tracker Users' };
    if (path.includes('sitegroups(99)/users')) return { value: member ? [{ LoginName: person.loginName }] : [] };
    if (path.includes('sitegroups(2)/users')) return { value: [{ LoginName: person.loginName }] };
    if (path.includes('roledefinitions')) return { value: [{ Id: 124, RoleTypeKind: 2 }] };
    if (path.includes('getusereffectivepermissions')) return { d: { GetUserEffectivePermissions: { Low: deny && path.includes(deny) ? low : '33' } } };
    if (path.includes('GetFileByServerRelativePath')) return { Exists: true, Level: 1 };
    throw new Error(`Unexpected request ${path}`);
  });
  store.post = vi.fn(async path => {
    if (path === '/_api/web/sitegroups(99)/users' && !ignoreAdd) member = true;
    return { StatusCode: 0 };
  });
  return store;
}

describe('scoped invitations without site-wide Read', () => {
  it.each(['Viewer', 'Project Engineer', 'Manager'])('enrolls a new %s in Tracker Users, verifies access, then emails once', async role => {
    const store = setup({ member: false });
    await expect(store.shareSiteAccess(person, role, page)).resolves.toEqual({ access: 'Tracker Users', emailRequested: true });
    expect(store.permissions.syncGroups).toHaveBeenCalledWith(person, role);
    expect(store.post.mock.calls.map(([path]) => path)).toEqual(['/_api/web/sitegroups(99)/users', '/_api/SP.Web.ShareObject']);
    expect(store.post.mock.calls[0][1]).toEqual({ body: { LoginName: person.loginName } });
    const membershipReads = store.get.mock.calls.flatMap(([path], i) => path.includes('sitegroups(99)/users') ? [store.get.mock.invocationCallOrder[i]] : []);
    expect(membershipReads).toHaveLength(2);
    expect(membershipReads[0]).toBeLessThan(store.post.mock.invocationCallOrder[0]);
    expect(membershipReads[1]).toBeGreaterThan(store.post.mock.invocationCallOrder[0]);
    expect(membershipReads[1]).toBeLessThan(store.post.mock.invocationCallOrder[1]);
    const accessReads = store.get.mock.calls.flatMap(([path], i) => path.includes('getusereffectivepermissions') ? [store.get.mock.invocationCallOrder[i]] : []);
    expect(accessReads.every(order => order > membershipReads[1] && order < store.post.mock.invocationCallOrder[1])).toBe(true);
  });
  it('reuses membership on an explicit retry after an email failure', async () => {
    const store = setup({ member: false });
    const post = store.post.getMockImplementation();
    let failed = false;
    store.post.mockImplementation(async (path, options) => {
      if (path.includes('ShareObject') && !failed) { failed = true; throw new Error('Email denied'); }
      return post(path, options);
    });
    await expect(store.shareSiteAccess(person, 'Viewer', page)).rejects.toThrow('invitation email was not confirmed');
    await expect(store.shareSiteAccess(person, 'Viewer', page)).resolves.toMatchObject({ emailRequested: true });
    expect(store.post.mock.calls.filter(([path]) => path.includes('sitegroups(99)'))).toHaveLength(1);
    expect(store.post.mock.calls.filter(([path]) => path.includes('ShareObject'))).toHaveLength(2);
  });
  it('stops before email if an added membership is not confirmed', async () => {
    const store = setup({ member: false, ignoreAdd: true });
    await expect(store.shareSiteAccess(person, 'Viewer', page)).rejects.toThrow('SharePoint did not confirm group membership');
    expect(store.post.mock.calls.map(([path]) => path)).toEqual(['/_api/web/sitegroups(99)/users']);
  });
  it.each(['lookup', 'membership', 'malformed membership', 'add'])('reports a baseline %s failure without emailing', async failure => {
    const store = setup({ member: false });
    const get = store.get.getMockImplementation();
    store.get.mockImplementation(path => {
      if (failure === 'lookup' && path.includes('getbyname')) return { Id: 0 };
      if (failure === 'membership' && path.includes('sitegroups(99)')) throw new Error('403 denied');
      if (failure === 'malformed membership' && path.includes('sitegroups(99)')) return {};
      return get(path);
    });
    if (failure === 'add') store.post.mockRejectedValue(new Error('403 denied'));
    await expect(store.shareSiteAccess(person, 'Viewer', page)).rejects.toThrow('allow Tracker Managers to view and manage its membership');
    expect(store.post.mock.calls.some(([path]) => path.includes('ShareObject'))).toBe(false);
    if (failure !== 'add') expect(store.post).not.toHaveBeenCalled();
  });
  it('keeps verified membership but stops before email if a resource still denies Read', async () => {
    const store = setup({ member: false, deny: 'SiteAssets' });
    await expect(store.shareSiteAccess(person, 'Viewer', page)).rejects.toThrow('No invitation email was requested');
    expect(store.post.mock.calls.map(([path]) => path)).toEqual(['/_api/web/sitegroups(99)/users']);
  });
  it('identifies the Updates list grant and retries without reenrolling after list access is restored', async () => {
    const store = setup({ member: false });
    const get = store.get.getMockImplementation();
    let updatesReadable = false;
    store.get.mockImplementation(path => {
      if (!updatesReadable && path.includes("getbytitle('updates')/getusereffectivepermissions")) {
        return { GetUserEffectivePermissions: { Low: '32' } };
      }
      return get(path);
    });
    await expect(store.shareSiteAccess(person, 'Viewer', page)).rejects.toThrow(
      /Modernization-Tracker - Updates.*verified Tracker Users membership.*No invitation email was requested.*lacks the required Read access.*Permissions for this list.*grant Tracker Users Read on this list if missing.*retry the invitation/,
    );
    expect(store.post.mock.calls.map(([path]) => path)).toEqual(['/_api/web/sitegroups(99)/users']);

    updatesReadable = true;
    await expect(store.shareSiteAccess(person, 'Viewer', page)).resolves.toMatchObject({ emailRequested: true });
    expect(store.post.mock.calls.map(([path]) => path)).toEqual(['/_api/web/sitegroups(99)/users', '/_api/SP.Web.ShareObject']);
  });
  it('verifies the page, all seven lists, and configured assets before one explicit email', async () => {
    const store = setup();
    await expect(store.shareSiteAccess(person, 'Viewer', page)).resolves.toMatchObject({ emailRequested: true });
    expect(store.listApi.mock.calls.map(([key]) => key)).toEqual(CONTAINERS.map(c => c.key));
    expect(store.get.mock.calls.filter(([p]) => p.includes('getusereffectivepermissions'))).toHaveLength(9);
    expect(store.get.mock.calls.some(([p]) => p.startsWith('/_api/web/getusereffectivepermissions'))).toBe(false);
    expect(store.post).toHaveBeenCalledExactlyOnceWith('/_api/SP.Web.ShareObject', { body: expect.objectContaining({ url: page, sendEmail: true, propagateAcl: false }) });
  });
  it.each(['SitePages', 'SiteAssets', ...CONTAINERS.map(c => `getbytitle('${c.key}')`)])('does not email when access is missing for %s', async deny => {
    const store = setup({ deny });
    await expect(store.shareSiteAccess(person, 'Viewer', page)).rejects.toThrow('No invitation email was requested');
    expect(store.post).not.toHaveBeenCalled();
  });
  it.each([null, {}, ['33'], 'bad', -1, 1.5])('rejects malformed permission results: %j', async low => {
    const store = setup({ deny: "getbytitle('updates')", low });
    await expect(store.shareSiteAccess(person, 'Viewer', page)).rejects.toThrow(
      /Modernization-Tracker - Updates.*check the selected account’s effective permissions.*invalid effective permission result/,
    );
    expect(store.post).not.toHaveBeenCalled();
  });
  it('stops before email when a resource permission check is denied', async () => {
    const store = setup();
    const get = store.get.getMockImplementation();
    store.get.mockImplementation(path => path.includes("getbytitle('users')") ? Promise.reject(new Error('403 Access denied')) : get(path));
    await expect(store.shareSiteAccess(person, 'Viewer', page)).rejects.toThrow(
      /Modernization-Tracker - Users.*verified Tracker Users membership.*check the selected account’s effective permissions.*403 Access denied/,
    );
    expect(store.post).not.toHaveBeenCalled();
  });
  it('checks the configured folder without granting access or changing inheritance', async () => {
    const store = setup({ assets: [] });
    store.invitationAssetFolders = ['/sites/metsoft/SiteAssets/Modernization Tracker'];
    await store.shareSiteAccess(person, 'Viewer', page);
    expect(store.get.mock.calls.some(([p]) => decodeURIComponent(p).includes("GetFolderByServerRelativePath(decodedurl='/sites/metsoft/SiteAssets/Modernization Tracker')/ListItemAllFields/getusereffectivepermissions"))).toBe(true);
    const denied = setup({ assets: [], deny: 'GetFolderByServerRelativePath' });
    denied.invitationAssetFolders = store.invitationAssetFolders;
    await expect(denied.shareSiteAccess(person, 'Viewer', page)).rejects.toThrow('Tracker assets folder');
    expect(denied.post).not.toHaveBeenCalled();
  });
  it('validates the configured assets without inventing paths for embedded deployments', async () => {
    const store = setup({ assets: [] });
    await store.shareSiteAccess(person, 'Viewer', page);
    expect(store.get.mock.calls.filter(([p]) => p.includes('getusereffectivepermissions'))).toHaveLength(8);
    const foreign = setup({ assets: ['https://outside.example/asset.html'] });
    await expect(foreign.shareSiteAccess(person, 'Viewer', page)).rejects.toThrow('on this SharePoint site');
    expect(foreign.post).not.toHaveBeenCalled();
  });
});

it.each(['missing', 'denied', 'wrong group'])('never falls back to the suite group when Tracker Users is %s', async failure => {
  const store = setup({ member: false });
  const get = store.get.getMockImplementation();
  store.get.mockImplementation(path => {
    if (path.includes("getbyname('Tracker%20Users')")) {
      if (failure === 'wrong group') return { Id: 100, Title: 'Metrology App User' };
      throw Object.assign(new Error(failure === 'denied' ? 'Access denied' : 'Group not found'), { status: failure === 'denied' ? 403 : 404 });
    }
    if (path.includes('Metrology%20App%20User')) return { Id: 100, Title: 'Metrology App User' };
    return get(path);
  });
  await expect(store.shareSiteAccess(person, 'Viewer', page)).rejects.toThrow('Tracker Users');
  expect(store.get.mock.calls.some(([path]) => path.includes('Metrology%20App%20User'))).toBe(false);
  expect(store.post).not.toHaveBeenCalled();
});
it.each(['Viewer', 'Project Engineer', 'Manager'])('never reads or writes suite membership while inviting %s', async role => {
  const store = setup({ member: false });
  await store.shareSiteAccess(person, role, page);
  const requests = [...store.get.mock.calls, ...store.post.mock.calls].map(([path]) => path);
  expect(requests.some(path => path.includes("getbyname('Tracker%20Users')"))).toBe(true);
  expect(requests.some(path => /Metrology|associatedmembergroup|sitegroups\(100\)/.test(path))).toBe(false);
  expect(store.post.mock.calls.filter(([path]) => path.includes('/sitegroups'))).toEqual([['/_api/web/sitegroups(99)/users', { body: { LoginName: person.loginName } }]]);
});
