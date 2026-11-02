import { describe, it, expect, vi } from 'vitest';
import { createFeedback, validateFeedback } from './feedback';
const store = (owner = false) => ({ webUrl: 'https://example.test/sites/tracker', prefix: 'Modernization', currentUser: vi.fn().mockResolvedValue({ isSiteAdmin: owner }), get: vi.fn(async path => path.includes('associatedownergroup') ? { Id: 7 } : { value: [] }), post: vi.fn().mockResolvedValue({}) });
describe('private feedback', () => {
  it('validates and trims reports, rejecting missing, oversized, and invalid values', () => {
    expect(validateFeedback({ kind: 'Bug', title: ' Test ', details: ' Steps ' })).toEqual({ kind: 'Bug', title: 'Test', details: 'Steps' });
    for (const input of [{ kind: 'Other', title: 't', details: 'd' }, { kind: 'Bug', title: ' ', details: 'd' }, { kind: 'Feature', title: 'x'.repeat(201), details: 'd' }, { kind: 'Bug', title: 't', details: 'x'.repeat(10001) }]) expect(() => validateFeedback(input)).toThrow();
  });
  it.each(['Viewer', 'Project Engineer', 'Manager'])('allows %s submission without reading feedback or trusting author fields', async role => {
    const raw = store(); raw.currentUser.mockResolvedValue({ role });
    await createFeedback(raw).submit({ kind: 'Feature', title: 'Idea', details: 'Details', AuthorId: 99 });
    expect(raw.get).not.toHaveBeenCalled();
    expect(raw.post).toHaveBeenCalledWith("/_api/web/lists/getbytitle('Modernization-Tracker%20-%20Feedback')/items", { body: { Title: 'Idea', FeedbackKind: 'Feature', FeedbackDetails: 'Details' } });
  });
  it('blocks non-owner managers before fetching submission contents', async () => {
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

it('explains missing storage to submitters without attempting provisioning', async () => {
  const raw = store(); raw.post.mockRejectedValue(Object.assign(new Error('Not found'), { status: 404 }));
  await expect(createFeedback(raw).submit({ kind: 'Bug', title: 'Issue', details: 'Steps' })).rejects.toThrow('Set up private feedback');
  expect(raw.post).toHaveBeenCalledTimes(3); expect(raw.get).not.toHaveBeenCalled();
});

it('supports the old list title for add-only users without reading the catalog', async () => {
  const raw = store(); raw.post.mockRejectedValueOnce(Object.assign(new Error('Missing current title'), { status: 404 })).mockResolvedValueOnce({});
  await createFeedback(raw).submit({ kind: 'Bug', title: 'Issue', details: 'Steps' });
  expect(raw.post.mock.calls.map(([path]) => path)).toEqual(["/_api/web/lists/getbytitle('Modernization-Tracker%20-%20Feedback')/items", "/_api/web/lists/getbytitle('ModernizationFeedback')/items"]);
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
  expect(raw.get.mock.calls.some(([path]) => path.includes('/items'))).toBe(false);
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
it.each(['Viewer', 'Project Engineer', 'Manager'])('rejects updates by non-owner %s before writes', async role => {
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
it.each([[412, 'Another owner updated'], [404, 'was deleted'], [403, 'Denied'], [503, 'Denied']])('does not retry an update rejected with %s', async (status, message) => {
  const raw = updatingStore(); raw.post.mockRejectedValue(Object.assign(new Error('Denied'), { status }));
  await expect(createFeedback(raw).update(1, validUpdate, '"1"')).rejects.toThrow(message);
  expect(raw.post).toHaveBeenCalledTimes(1);
});
it('keeps old lists readable, requests version metadata, and defaults legacy status to New', async () => {
  const raw = store(true);
  raw.get.mockResolvedValueOnce(feedbackList).mockResolvedValueOnce({ value: [{ Id: 1, Title: 'Legacy', 'odata.etag': '"2"' }, { Id: 2, FeedbackStatus: 'Resolved', __metadata: { etag: '"4"' } }] });
  const result = await createFeedback(raw).list();
  expect(result[0]).toMatchObject({ status: 'New', etag: '"2"' });
  expect(result[1]).toMatchObject({ status: 'Resolved', etag: '"4"' });
  expect(raw.get).toHaveBeenLastCalledWith(expect.stringContaining('$select=*,Author/Title'), { Accept: 'application/json;odata=minimalmetadata' });
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
