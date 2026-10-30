import { expect, it, vi } from 'vitest';
import { ensureFeedbackStatus, FEEDBACK_STATUSES } from './feedbackFields';
const field = { d: { results: [{ InternalName: 'FeedbackStatus', TypeAsString: 'Choice', Choices: { results: FEEDBACK_STATUSES } }] } };
it('adds and verifies the status column on an existing list without touching items or ACLs', async () => {
  const store = { get: vi.fn().mockResolvedValueOnce({ value: [] }).mockResolvedValue(field), post: vi.fn().mockResolvedValue({}) };
  await ensureFeedbackStatus(store, '/list');
  expect(store.post).toHaveBeenCalledExactlyOnceWith('/list/fields/createfieldasxml', expect.objectContaining({ verbose: true }));
  expect(store.post.mock.calls[0][1].body.parameters.SchemaXml).toContain('<Default>New</Default>');
});
it('accepts a compatible column created concurrently by another owner', async () => {
  const store = { get: vi.fn().mockResolvedValueOnce({ value: [] }).mockResolvedValue(field), post: vi.fn().mockRejectedValue(new Error('Already exists')) };
  await ensureFeedbackStatus(store, '/list');
  expect(store.post).toHaveBeenCalledTimes(1);
});
it('does not overwrite an incompatible existing column', async () => {
  const store = { get: vi.fn().mockResolvedValue({ value: [{ InternalName: 'FeedbackStatus', TypeAsString: 'Text' }] }), post: vi.fn() };
  await expect(ensureFeedbackStatus(store, '/list')).rejects.toThrow('review the FeedbackStatus');
  expect(store.post).not.toHaveBeenCalled();
});
it('does not treat denied field access as a missing column', async () => {
  const store = { get: vi.fn().mockRejectedValue(Object.assign(new Error('Denied'), { status: 403 })), post: vi.fn() };
  await expect(ensureFeedbackStatus(store, '/list')).rejects.toThrow('Denied');
  expect(store.post).not.toHaveBeenCalled();
});
it('stops when creation fails and does not retry the mutation', async () => {
  const store = { get: vi.fn().mockResolvedValue({ value: [] }), post: vi.fn().mockRejectedValue(new Error('Manage Lists required')) };
  await expect(ensureFeedbackStatus(store, '/list')).rejects.toThrow('Manage Lists required');
  expect(store.post).toHaveBeenCalledTimes(1);
});
