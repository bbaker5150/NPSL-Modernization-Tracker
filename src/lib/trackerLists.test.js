import { describe, expect, it, vi } from 'vitest';
import { TrackerLists, TRACKER_LISTS, displayListTitle, legacyListTitle } from './trackerLists';
import { SharePointStore } from './spStore';
import { SharePointError } from './spContext';

function fixture() {
  const rows = TRACKER_LISTS.map(([key], index) => ({ Id: `00000000-0000-0000-0000-00000000000${index}`, Title: legacyListTitle('Modernization', key), BaseTemplate: 100, Hidden: true }));
  const get = vi.fn(async path => {
    if (path.endsWith('EffectiveBasePermissions')) return { Low: '2048' };
    if (path.includes('/lists?')) {
      const filter = new URLSearchParams(path.split('?')[1]).get('$filter');
      return { value: rows.filter(row => filter.includes(`'${row.Title}'`)).map(row => ({ ...row })) };
    }
    const row = rows.find(row => path.includes(row.Id));
    if (!row) throw new Error(`Unexpected GET: ${path}`);
    return { ...row };
  });
  const post = vi.fn(async (path, { body }) => {
    const row = rows.find(row => path.includes(row.Id));
    if (!row) throw new Error(`Unexpected POST: ${path}`);
    Object.assign(row, body);
  });
  return { rows, get, post, catalog: new TrackerLists('Modernization', get, post) };
}

describe('tracker list organization', () => {
  it('resolves legacy and friendly titles to the same GUID without viewer writes', async () => {
    const { rows, catalog, get, post } = fixture();
    rows[1].Title = 'NPSL Tracker - Tasks'; rows[1].Hidden = false;
    const requests = await Promise.all([
      catalog.rewrite("/_api/web/lists/getbytitle('ModernizationTasks')/items(42)"),
      catalog.rewrite("/_api/web/lists/getbytitle('NPSL%20Tracker%20-%20Tasks')/AttachmentFiles"),
    ]);
    expect(requests).toEqual([`/_api/web/lists(guid'${rows[1].Id}')/items(42)`, `/_api/web/lists(guid'${rows[1].Id}')/AttachmentFiles`]);
    expect(get).toHaveBeenCalledTimes(1);
    expect(post).not.toHaveBeenCalled();
    expect(await catalog.rewrite("/_api/web/lists/getbytitle('AnotherApp')/items")).toBe("/_api/web/lists/getbytitle('AnotherApp')/items");
  });
  it('renames all seven in place and verifies visibility, then reruns without further writes', async () => {
    const { rows, catalog, post } = fixture();
    const originalIds = rows.map(row => row.Id);
    await expect(catalog.organize()).resolves.toBe(7);
    expect(rows.map(row => row.Id)).toEqual(originalIds);
    expect(rows.every(row => !row.Hidden && row.Title.startsWith('Modernization-Tracker - '))).toBe(true);
    expect(post).toHaveBeenCalledTimes(7);
    for (const [path, options] of post.mock.calls) {
      expect(path).toMatch(/\/lists\(guid'/);
      expect(Object.keys(options.body).sort()).toEqual(['Hidden', 'Title']);
    }
    await expect(catalog.organize()).resolves.toBe(7);
    expect(post).toHaveBeenCalledTimes(7);
  });
  it('preflights duplicates and missing lists before any rename', async () => {
    for (const kind of ['duplicate', 'missing']) {
      const { rows, catalog, post } = fixture();
      if (kind === 'duplicate') rows.push({ ...rows[6], Id: '11111111-1111-1111-1111-111111111111', Title: 'NPSL Tracker - Reference Documents' });
      else rows.pop();
      await expect(catalog.organize()).rejects.toThrow(kind === 'duplicate' ? 'Multiple' : 'not found');
      expect(post).not.toHaveBeenCalled();
    }
  });
  it('requires Manage Lists and preserves read failures rather than treating them as absent lists', async () => {
    const { catalog, get, post } = fixture();
    get.mockResolvedValueOnce({ Low: '0' });
    await expect(catalog.organize()).rejects.toThrow('Manage Lists');
    expect(post).not.toHaveBeenCalled();
    get.mockRejectedValueOnce(new SharePointError('Access denied', 403));
    await expect(catalog.path('users')).rejects.toMatchObject({ status: 403 });
    await expect(catalog.path('users')).resolves.toContain("guid'");
  });
  it('resumes a partial rename with the same IDs after a failed request', async () => {
    const { rows, catalog, post } = fixture();
    const originalPost = post.getMockImplementation();
    post.mockImplementationOnce(originalPost).mockRejectedValueOnce(new Error('Connection interrupted'));
    await expect(catalog.organize()).rejects.toThrow('Connection interrupted');
    expect(rows[0].Title).toBe('Modernization-Tracker - Projects');
    expect(rows[1].Title).toBe('ModernizationTasks');
    await expect(catalog.organize()).resolves.toBe(7);
    expect(post).toHaveBeenCalledTimes(8); // one completed, one failed, six resumed
  });
  it('does not report success when SharePoint fails to apply a rename', async () => {
    const { catalog, post } = fixture();
    post.mockResolvedValueOnce({});
    await expect(catalog.organize()).rejects.toThrow('Could not verify');
  });
  it('automatically maintains names only for site owners and accepts all three aliases', async () => {
    const { rows, catalog, get, post } = fixture();
    rows[0].Title = 'NPSL Tracker - Projects';
    rows[1].Title = 'Modernization-Tracker - Tasks';
    get.mockResolvedValueOnce({ Low: '2048' });
    expect(await catalog.maintainNames()).toBe(0);
    expect(post).not.toHaveBeenCalled();
    get.mockResolvedValueOnce({ Low: '33556480' });
    expect(await catalog.maintainNames()).toBe(7);
    expect(rows.every(row => row.Title.startsWith('Modernization-Tracker - '))).toBe(true);
  });
  it('keeps custom list namespaces separate' , () => {
    expect(displayListTitle('OtherApp', 'projects')).toBe('OtherApp Tracker - Projects');
    expect(legacyListTitle('OtherApp', 'references')).toBe('OtherAppReferenceDocuments');
  });
  it('routes real store reads and permission-style writes through GUIDs after rename', async () => {
    const requests = [];
    const fetchImpl = async (url, options = {}) => {
      requests.push({ url, options });
      const body = url.includes('/lists?') ? { value: [{ Id: '11111111-1111-1111-1111-111111111111', Title: 'NPSL Tracker - Projects', BaseTemplate: 100, Hidden: false }] }
        : url.endsWith('/contextinfo') ? { FormDigestValue: 'digest', FormDigestTimeoutSeconds: 1800 }
        : url.includes('/items?') ? { value: [{ Id: 3, Title: 'Existing project' }] } : {};
      return new Response(JSON.stringify(body));
    };
    const store = new SharePointStore({ webUrl: 'https://example.invalid/sites/catalog', fetchImpl });
    expect(await store.listItems('projects', [], row => row.Title)).toEqual(['Existing project']);
    await store.post("/_api/web/lists/getbytitle('ModernizationProjects')/items(3)/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)", { body: {} });
    expect(requests.filter(({ url }) => url.includes('/lists?'))).toHaveLength(1);
    expect(requests.at(-1).url).toContain("/lists(guid'11111111-1111-1111-1111-111111111111')/items(3)/breakroleinheritance");
    expect(requests.at(-1).options.method).toBe('POST');
  });
});
