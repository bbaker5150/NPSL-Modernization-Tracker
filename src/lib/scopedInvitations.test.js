import { describe, expect, it, vi } from 'vitest';
import { CONTAINERS, SharePointStore } from './spStore';

const webUrl = 'https://tenant.sharepoint-mil.us/sites/metsoft';
const page = `${webUrl}/SitePages/Modernization-Tracker.aspx`;
const asset = `${webUrl}/SiteAssets/Tracker/tracker.html`;
const person = { loginName: 'claims|viewer' };
function setup({ deny = '', low = '0', assets = [asset] } = {}) {
  const store = new SharePointStore({ webUrl, scopedAccess: true, invitationAssetUrls: assets });
  store.permissions.syncGroups = vi.fn(async () => null);
  store.listApi = vi.fn(async key => `/_api/web/lists/getbytitle('${key}')`);
  store.get = vi.fn(async path => {
    if (path.startsWith('/_api/web/getusereffectivepermissions')) throw new Error('Site-wide access must not be required');
    if (path.includes('sitegroups(2)/users')) return { value: [{ LoginName: person.loginName }] };
    if (path.includes('roledefinitions')) return { value: [{ Id: 124, RoleTypeKind: 2 }] };
    if (path.includes('getusereffectivepermissions')) return { d: { GetUserEffectivePermissions: { Low: deny && path.includes(deny) ? low : '33' } } };
    if (path.includes('GetFileByServerRelativePath')) return { Exists: true, Level: 1 };
    throw new Error(`Unexpected request ${path}`);
  });
  store.post = vi.fn(async () => ({ StatusCode: 0 }));
  return store;
}

describe('scoped invitations without site-wide Read', () => {
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
    const store = setup({ deny: 'SitePages', low });
    await expect(store.shareSiteAccess(person, 'Viewer', page)).rejects.toThrow('Required Read access was not confirmed');
    expect(store.post).not.toHaveBeenCalled();
  });
  it('stops before email when a resource permission check is denied', async () => {
    const store = setup();
    const get = store.get.getMockImplementation();
    store.get.mockImplementation(path => path.includes("getbytitle('users')") ? Promise.reject(new Error('403 Access denied')) : get(path));
    await expect(store.shareSiteAccess(person, 'Viewer', page)).rejects.toThrow('Modernization-Tracker - Users');
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
