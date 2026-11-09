import { describe, it, expect, vi } from 'vitest';
import { createFeedback, validateFeedback } from './feedback';
const store = (owner = false) => ({ webUrl: 'https://example.test/sites/tracker', prefix: 'Modernization', currentUser: vi.fn().mockResolvedValue({ isSiteAdmin: owner }), get: vi.fn(async path => path.includes('associatedownergroup') ? { Id: 7 } : { value: [] }), post: vi.fn().mockResolvedValue([{ ErrorInfo: null }]) });
describe('private feedback', () => {
  it('validates and trims reports, rejecting missing, oversized, and invalid values', () => {
    expect(validateFeedback({ kind: 'Bug', title: ' Test ', details: ' Steps ' })).toEqual({ kind: 'Bug', title: 'Test', details: 'Steps' });
    for (const input of [{ kind: 'Other', title: 't', details: 'd' }, { kind: 'Bug', title: ' ', details: 'd' }, { kind: 'Feature', title: 'x'.repeat(201), details: 'd' }, { kind: 'Bug', title: 't', details: 'x'.repeat(10001) }]) expect(() => validateFeedback(input)).toThrow();
  });
  it.each(['Viewer', 'Project Engineer', 'Manager'])('allows %s submission without reading feedback or trusting author fields', async role => {
    const raw = store(); raw.currentUser.mockResolvedValue({ role });
    await createFeedback(raw).submit({ kind: 'Feature', title: 'Idea', details: 'Details', AuthorId: 99 });
    expect(raw.get).not.toHaveBeenCalled();
    expect(raw.post).toHaveBeenCalledExactlyOnceWith('/_vti_bin/client.svc/ProcessQuery', expect.objectContaining({ raw: true, body: expect.stringContaining('Modernization-Tracker - Feedback') }));
    expect(raw.post.mock.calls[0][1].body).not.toContain('AuthorId');
  });
  it('ignores an unverified Manager label before fetching submission contents', async () => {
    const raw = store(); raw.currentUser.mockResolvedValue({ role: 'Manager' });
    await expect(createFeedback(raw).list()).rejects.toThrow('Only site owners');
    expect(raw.get.mock.calls.some(([path]) => path.includes('/items'))).toBe(false);
  });
  it('recognizes associated Owners membership and site administrators', async () => {
    const raw = store(); raw.get.mockImplementation(async path => path.includes('associatedownergroup') ? { d: { Id: 7 } } : { d: { results: [{ Id: 7 }] } });
    expect(await createFeedback(raw).isOwner()).toBe(true);
    expect(await createFeedback(store(true)).isOwner()).toBe(true);
  });
  it.each([
    { Low: '4294967295', High: '2147483647' },
    { d: { EffectiveBasePermissions: { Low: '33554432', High: '0' } } },
  ])('recognizes effective site owner permissions without direct Owners membership', async permissions => {
    const raw = store(); raw.get.mockResolvedValue(permissions);
    expect(await createFeedback(raw).isOwner()).toBe(true);
    expect(raw.get).toHaveBeenCalledExactlyOnceWith('/_api/web/effectivebasepermissions');
  });
  it('does not treat ordinary site Edit or app Manager access as site ownership', async () => {
    const raw = store(); raw.currentUser.mockResolvedValue({ role: 'Manager' });
    raw.get.mockImplementation(async path => path.includes('effectivebasepermissions') ? { Low: '1011030767' } : path.includes('associatedownergroup') ? { Id: 7 } : { value: [{ Id: 9 }] });
    expect(await createFeedback(raw).isOwner()).toBe(false);
    await expect(createFeedback(raw).list()).rejects.toThrow('Only site owners');
    expect(raw.get.mock.calls.some(([path]) => path.includes('/items'))).toBe(false);
  });
  it('allows verified Owners membership when effective permissions are unavailable', async () => {
    const raw = store(); raw.get.mockImplementation(async path => {
      if (path.includes('effectivebasepermissions')) throw new Error('Unavailable');
      return path.includes('associatedownergroup') ? { Id: 7 } : { value: [{ Id: 7 }] };
    });
    expect(await createFeedback(raw).isOwner()).toBe(true);
  });
  it('fails closed when owner verification fails', async () => {
    const raw = store(); raw.get.mockRejectedValue(new Error('Denied'));
    await expect(createFeedback(raw).list()).rejects.toThrow('Denied');
    expect(raw.get.mock.calls.some(([path]) => path.includes('/items'))).toBe(false);
  });
  it('propagates failed writes without retrying', async () => {
    const raw = store(); raw.post.mockRejectedValue(new Error('Timeout'));
    await expect(createFeedback(raw).submit({ kind: 'Bug', title: 't', details: 'd' })).rejects.toThrow('Timeout');
    expect(raw.post).toHaveBeenCalledTimes(1);
  });
  it('loads paginated submissions for owners', async () => {
    const raw = store(true);
    raw.get.mockResolvedValueOnce({ value: [{ Id: '11111111-1111-1111-1111-111111111111', Title: 'Modernization-Tracker - Feedback', BaseTemplate: 100 }] }).mockResolvedValueOnce({ value: [{ Id: 1, Title: 'Bug', FeedbackKind: 'Bug', FeedbackDetails: 'Steps', Author: { Title: 'User' } }], 'odata.nextLink': `${raw.webUrl}/_api/next` }).mockResolvedValueOnce({ value: [{ Id: 2, Title: 'Idea' }] });
    const items = await createFeedback(raw).list();
    expect(items).toHaveLength(2); expect(items[0]).toMatchObject({ title: 'Bug', author: 'User', details: 'Steps' });
  });
});

const missingList = title => [{ ErrorInfo: { ErrorTypeName: 'System.ArgumentException', ErrorCode: -2147024809, ErrorMessage: `List '${title}' does not exist at site with URL 'https://example.test'.` } }];

it('explains missing storage to submitters without attempting provisioning', async () => {
  const raw = store(); raw.post.mockImplementation(async (_, { body }) => missingList(new DOMParser().parseFromString(body, 'text/xml').querySelector('Method[Name=GetByTitle] Parameter').textContent));
  await expect(createFeedback(raw).submit({ kind: 'Bug', title: 'Issue', details: 'Steps' })).rejects.toThrow('Set up private feedback');
  expect(raw.post).toHaveBeenCalledTimes(3); expect(raw.get).not.toHaveBeenCalled();
});

it('supports the old list title for add-only users without reading the catalog', async () => {
  const raw = store(); raw.post.mockResolvedValueOnce(missingList('Modernization-Tracker - Feedback')).mockResolvedValueOnce([{ ErrorInfo: null }]);
  await createFeedback(raw).submit({ kind: 'Bug', title: 'Issue', details: 'Steps' });
  expect(raw.post.mock.calls.map(([, { body }]) => new DOMParser().parseFromString(body, 'text/xml').querySelector('Method[Name=GetByTitle] Parameter').textContent)).toEqual(['Modernization-Tracker - Feedback', 'ModernizationFeedback']);
  expect(raw.get).not.toHaveBeenCalled();
});

const feedbackList = { value: [{ Id: '11111111-1111-1111-1111-111111111111', Title: 'Modernization-Tracker - Feedback', BaseTemplate: 100 }] };
it.each(['Viewer', 'Project Engineer', 'Manager'])('rejects clearing by a non-owner %s before writes', async role => {
  const raw = store(); raw.currentUser.mockResolvedValue({ role });
  await expect(createFeedback(raw).clearInbox([1])).rejects.toThrow('Only site owners');
  expect(raw.post).not.toHaveBeenCalled();
});
it('recycles only the reviewed IDs, deduplicates them, and reports progress', async () => {
  const raw = store(true); raw.get.mockResolvedValue(feedbackList); const progress = vi.fn();
  expect(await createFeedback(raw).clearInbox([1, 2, 1], progress)).toEqual({ clearedIds: [1, 2] });
  expect(raw.post.mock.calls.map(([path]) => path)).toEqual([1, 2].map(id => `/_api/web/lists(guid'11111111-1111-1111-1111-111111111111')/items(${id})/recycle()`));
  expect(progress).toHaveBeenLastCalledWith({ cleared: 2, total: 2 });
  expect(raw.get.mock.calls.filter(([path]) => path.includes('/items')).map(([path]) => path)).toEqual([1, 2].map(id => `/_api/web/lists(guid'11111111-1111-1111-1111-111111111111')/items(${id})?$select=*`));
});
it.each([[0], [1, -2], ['1'], [1, '2)/recycle()'], []].map(ids => [ids]))('rejects invalid clear IDs %j before any writes', async ids => {
  const raw = store(true); await expect(createFeedback(raw).clearInbox(ids)).rejects.toThrow('valid feedback');
  expect(raw.post).not.toHaveBeenCalled();
});
it('reports partial failure and stops instead of retrying uncertain writes', async () => {
  const raw = store(true); raw.get.mockResolvedValue(feedbackList);
  raw.post.mockResolvedValueOnce({}).mockRejectedValueOnce(new Error('Timeout'));
  await expect(createFeedback(raw).clearInbox([1, 2, 3])).rejects.toMatchObject({ clearedIds: [1], message: expect.stringContaining('Cleared 1 of 3') });
  expect(raw.post).toHaveBeenCalledTimes(2);
});
it('tolerates a reviewed submission already removed by another owner', async () => {
  const raw = store(true); raw.get.mockResolvedValue(feedbackList);
  raw.post.mockRejectedValueOnce(Object.assign(new Error('Missing'), { status: 404 })).mockResolvedValueOnce({});
  expect(await createFeedback(raw).clearInbox([1, 2])).toEqual({ clearedIds: [1, 2] });
});
it('preserves feedback received after the owner loaded the inbox', async () => {
  const api = createFeedback({ currentUser: async () => ({ isSiteAdmin: true }) });
  await api.submit({ kind: 'Bug', title: 'Reviewed', details: 'Details' });
  const snapshot = await api.list();
  await api.submit({ kind: 'Feature', title: 'New arrival', details: 'Details' });
  await api.clearInbox(snapshot.map(item => item.id));
  expect((await api.list()).map(item => item.title)).toEqual(['New arrival']);
});

const validUpdate = { kind: 'Feature', title: 'Updated', details: 'Updated details', status: 'In progress' };
const statusField = { InternalName: 'FeedbackStatus', TypeAsString: 'Choice', Choices: ['New', 'In progress', 'Resolved', 'Closed'] };
const updatingStore = () => {
  const raw = store(true);
  raw.get.mockImplementation(async path => path.includes('/fields?') ? { value: [statusField] } : feedbackList);
  return raw;
};
it.each(['Viewer', 'Project Engineer', 'Manager'])('rejects an unverified %s label before writes', async role => {
  const raw = store(); raw.currentUser.mockResolvedValue({ role });
  await expect(createFeedback(raw).update(1, validUpdate, '"2"')).rejects.toThrow('Only site owners');
  expect(raw.post).not.toHaveBeenCalled();
});
it('updates only the selected record with the reviewed version and editable fields', async () => {
  const raw = updatingStore();
  await createFeedback(raw).update(2, { ...validUpdate, AuthorId: 99, Created: 'changed' }, '"3"');
  expect(raw.post).toHaveBeenCalledExactlyOnceWith("/_api/web/lists(guid'11111111-1111-1111-1111-111111111111')/items(2)", {
    body: { Title: 'Updated', FeedbackKind: 'Feature', FeedbackDetails: 'Updated details', FeedbackStatus: 'In progress' },
    headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '"3"' },
  });
});
it.each([
  [0, validUpdate, '"1"'], ['1', validUpdate, '"1"'], [1, { ...validUpdate, title: '' }, '"1"'],
  [1, { ...validUpdate, status: 'Invalid' }, '"1"'], [1, validUpdate, undefined], [1, validUpdate, '*'],
])('rejects invalid update inputs before writes (%j)', async (id, value, etag) => {
  const raw = updatingStore(); await expect(createFeedback(raw).update(id, value, etag)).rejects.toThrow();
  expect(raw.post).not.toHaveBeenCalled();
});
it.each([[412, 'Another reviewer updated'], [404, 'was deleted'], [403, 'Denied'], [503, 'Denied']])('does not retry an update rejected with %s', async (status, message) => {
  const raw = updatingStore(); raw.post.mockRejectedValue(Object.assign(new Error('Denied'), { status }));
  await expect(createFeedback(raw).update(1, validUpdate, '"1"')).rejects.toThrow(message);
  expect(raw.post).toHaveBeenCalledTimes(1);
});
it('keeps old lists readable, requests version metadata, and defaults legacy status to New', async () => {
  const raw = store(true);
  raw.get.mockResolvedValueOnce(feedbackList).mockResolvedValueOnce({ d: { results: [{ Id: 1, Title: 'Legacy', 'odata.etag': '"2"' }, { Id: 2, FeedbackStatus: 'Resolved', __metadata: { etag: '"4"' } }] } });
  const result = await createFeedback(raw).list();
  expect(result[0]).toMatchObject({ status: 'New', etag: '"2"' });
  expect(result[1]).toMatchObject({ status: 'Resolved', etag: '"4"' });
  expect(raw.get).toHaveBeenLastCalledWith(expect.stringContaining('$select=*,Author/Title'), { Accept: 'application/json;odata=verbose' });
});
it('updates and deletes one preview record while preserving other submissions and rejecting stale versions', async () => {
  const api = createFeedback({ currentUser: async () => ({ isSiteAdmin: true }) });
  await api.submit({ kind: 'Bug', title: 'Keep me', details: 'Original' });
  await api.submit({ kind: 'Bug', title: 'Edit me', details: 'Original' });
  const [selected, kept] = await api.list();
  await api.update(selected.id, validUpdate, selected.etag);
  await expect(api.update(selected.id, validUpdate, selected.etag)).rejects.toThrow('changed');
  expect((await api.list())[1]).toEqual(kept);
  expect((await api.list())[0]).toMatchObject(validUpdate);
  await api.clearInbox([selected.id]);
  expect(await api.list()).toEqual([kept]);
});

it('allows verified Tracker Managers to read and edit history, and rechecks membership before each operation', async () => {
  const raw = store(); let manager = true;
  raw.get.mockImplementation(async path => {
    if (path.includes('effectivebasepermissions')) return { Low: '1' };
    if (path.includes('associatedownergroup')) return { Id: 7 };
    if (path.includes('currentuser/groups')) return { value: manager ? [{ Id: 22, Title: 'Tracker Managers' }] : [] };
    if (path.includes('/lists?')) return feedbackList;
    if (path.includes('/fields?')) return { value: [statusField] };
    if (path.includes('/items?')) return { value: [{ Id: 5, Title: 'Existing report', 'odata.etag': '"1"' }] };
    throw new Error('Unexpected read ' + path);
  });
  const api = createFeedback(raw);
  expect(await api.access()).toEqual({ owner: false, review: true });
  expect(await api.list()).toMatchObject([{ id: 5, title: 'Existing report' }]);
  await api.update(5, validUpdate, '"1"');
  await expect(api.clearInbox([5])).rejects.toThrow('Only site owners');
  await expect(api.setup.enableManagerReview()).rejects.toThrow('Only site owners');
  manager = false;
  await expect(api.update(5, validUpdate, '"2"')).rejects.toThrow('Only site owners');
  await expect(api.list()).rejects.toThrow('Only site owners');
  expect(raw.post).toHaveBeenCalledTimes(1);
});

it('uploads selected files before the report with no reads by an add-only submitter', async () => {
  const raw = store();
  await createFeedback(raw).submit({ kind: 'Bug', title: 'Screenshot attached', details: 'Steps', attachments: [new File(['pixels'], 'screen.png', { type: 'image/png' })] });
  expect(raw.get).not.toHaveBeenCalled();
  expect(raw.post).toHaveBeenCalledTimes(2);
  const upload = new DOMParser().parseFromString(raw.post.mock.calls[0][1].body, 'text/xml');
  expect(upload.querySelector('Actions > ObjectPath')).not.toBeNull();
  const report = new DOMParser().parseFromString(raw.post.mock.calls[1][1].body, 'text/xml');
  const fields = Object.fromEntries([...report.querySelectorAll('Actions > Method[Name=SetFieldValue]')].map(node => [...node.querySelectorAll('Parameter')].map(p => p.textContent)));
  const [file] = JSON.parse(fields.FeedbackAttachments);
  expect(file).toMatchObject({ name: 'screen.png', size: 6, type: 'image/png' });
  expect(file.key).toBe(upload.querySelector('Property[Name=Url]').textContent);
  expect(report.querySelector('Query, ObjectIdentityQuery')).toBeNull();
});
it('does not create a report after an unconfirmed upload and does not replay that upload', async () => {
  const raw = store(); raw.post.mockRejectedValue(new Error('Timeout'));
  const api = createFeedback(raw), report = { kind: 'Bug', title: 'Screenshot attached', details: 'Steps', attachments: [new File(['pixels'], 'screen.png')] };
  await expect(api.submit(report)).rejects.toThrow('No report was sent');
  await expect(api.submit(report)).rejects.toThrow('upload was not confirmed');
  expect(raw.post).toHaveBeenCalledTimes(1);
});
it('keeps attachment downloads behind verified reviewer access and a submission manifest', async () => {
  const raw = store(); raw.fetchImpl = vi.fn();
  await expect(createFeedback(raw).downloadAttachment(1, 'anything.png')).rejects.toThrow('Only site owners');
  expect(raw.fetchImpl).not.toHaveBeenCalled();
  expect(raw.get.mock.calls.some(([path]) => path.includes('/items'))).toBe(false);
});
it('allows owner and verified manager attachment downloads and rejects arbitrary file keys', async () => {
  const key = '11111111-1111-4111-8111-111111111111.png';
  const attachment = { key, name: 'screen.png', size: 6, type: 'image/png' };
  for (const owner of [true, false]) {
    const raw = store(owner);
    raw.get.mockImplementation(async path => {
      if (path.includes('effectivebasepermissions')) return { Low: '1' };
      if (path.includes('currentuser/groups')) return { value: [{ Id: 22, Title: 'Tracker Managers' }] };
      if (path.includes('associatedownergroup')) return { Id: 7 };
      if (path.includes('/lists?')) return feedbackList;
      if (path.includes('/items(1)')) return { FeedbackAttachments: JSON.stringify([attachment]) };
      if (path.includes('/RootFolder?')) return { ServerRelativeUrl: '/sites/tracker/FeedbackFiles' };
      throw new Error('Unexpected read ' + path);
    });
    raw.fetchImpl = vi.fn().mockResolvedValue({ ok: true, blob: async () => new Blob(['pixels with updated file properties']) });
    const api = createFeedback(raw);
    expect((await api.downloadAttachment(1, key)).attachment).toEqual(attachment);
    expect(raw.fetchImpl).toHaveBeenCalledWith(`https://example.test/sites/tracker/FeedbackFiles/${key}`, expect.objectContaining({ credentials: 'include' }));
    await expect(api.downloadAttachment(1, '22222222-2222-4222-8222-222222222222.png')).rejects.toThrow('no longer linked');
    expect(raw.fetchImpl).toHaveBeenCalledTimes(1);
  }
});
it('recycles linked files before the report and retains the report when a file cannot be recycled', async () => {
  const raw = store(true), file = { key: '11111111-1111-4111-8111-111111111111.png', name: 'screen.png', size: 6, type: 'image/png' };
  raw.get.mockImplementation(async path => path.includes('/items(1)') ? { FeedbackAttachments: JSON.stringify([file]) } : path.includes('/RootFolder?') ? { ServerRelativeUrl: '/sites/tracker/FeedbackFiles' } : feedbackList);
  raw.post.mockRejectedValueOnce(new Error('Denied'));
  await expect(createFeedback(raw).clearInbox([1])).rejects.toMatchObject({ clearedIds: [], message: expect.stringContaining('submission was kept') });
  expect(raw.post).toHaveBeenCalledTimes(1);
  expect(raw.post.mock.calls[0][0]).toContain('GetFileByServerRelativeUrl');
  raw.post.mockResolvedValue({});
  await createFeedback(raw).clearInbox([1]);
  expect(raw.post.mock.calls[2][0]).toContain('/items(1)/recycle()');
});

it('blocks replay of an uncertain report using already uploaded files so two reports cannot share deletion ownership', async () => {
  const raw = store(); raw.post.mockResolvedValueOnce([{ ErrorInfo: null }]).mockRejectedValueOnce(new Error('Report response timed out'));
  const api = createFeedback(raw), file = new File(['pixels'], 'screen.png');
  const report = { kind: 'Bug', title: 'Possible saved report', details: 'Steps', attachments: [file] };
  await expect(api.submit(report)).rejects.toMatchObject({ submissionUncertain: true });
  await expect(api.submit(report)).rejects.toThrow('has not been sent again');
  await expect(api.submit({ ...report, title: 'Changed draft' })).rejects.toThrow('has not been sent again');
  expect(raw.post).toHaveBeenCalledTimes(2);
});

it('gives new confirmed reports independent file ownership even when the same File object is selected twice', async () => {
  const raw = store(), api = createFeedback(raw), file = new File(['pixels'], 'screen.png');
  const report = { kind: 'Bug', title: 'Report', details: 'Steps', attachments: [file] };
  await api.submit(report); await api.submit({ ...report, title: 'Another report' });
  expect(raw.post).toHaveBeenCalledTimes(4);
  const keys = [0, 2].map(index => new DOMParser().parseFromString(raw.post.mock.calls[index][1].body, 'text/xml').querySelector('Property[Name=Url]').textContent);
  expect(keys[0]).not.toBe(keys[1]);
});
