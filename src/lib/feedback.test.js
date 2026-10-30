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
    expect(raw.post).toHaveBeenCalledWith("/_api/web/lists/getbytitle('ModernizationFeedback')/items", { body: { Title: 'Idea', FeedbackKind: 'Feature', FeedbackDetails: 'Details' } });
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
    raw.get.mockResolvedValueOnce({ value: [{ Id: 1, Title: 'Bug', FeedbackKind: 'Bug', FeedbackDetails: 'Steps', Author: { Title: 'User' } }], 'odata.nextLink': `${raw.webUrl}/_api/next` }).mockResolvedValueOnce({ value: [{ Id: 2, Title: 'Idea' }] });
    const items = await createFeedback(raw).list();
    expect(items).toHaveLength(2); expect(items[0]).toMatchObject({ title: 'Bug', author: 'User', details: 'Steps' });
  });
});
