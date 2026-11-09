import { expect, it, vi } from 'vitest';
import { attachmentDescriptor, feedbackFilePath, feedbackFileRequest, feedbackFileUploader, parseFeedbackAttachments, validateFeedbackFiles, MAX_FEEDBACK_FILE_BYTES } from './feedbackAttachments';

const file = (name = 'screenshot.png', content = 'image bytes') => new File([content], name, { type: 'image/png' });
it('validates actual file size, counts, path characters and permitted formats before any upload', () => {
  expect(validateFeedbackFiles([file(), file('steps.pdf')])).toHaveLength(2);
  for (const files of [[file('script.html')], [file('graphic.svg')], [file('../secret.txt')], [file('empty.txt', '')], [file('large.txt', new Uint8Array(MAX_FEEDBACK_FILE_BYTES + 1))], Array.from({ length: 6 }, () => file())]) expect(() => validateFeedbackFiles(files)).toThrow();
});
it('uses opaque immutable filenames and validates stored manifests without accepting URLs or forged content types', () => {
  const descriptor = attachmentDescriptor(file('Screen & details.png'));
  expect(descriptor.key).toMatch(/^[\da-f-]{36}\.png$/);
  expect(descriptor.key).not.toContain('Screen');
  expect(parseFeedbackAttachments(JSON.stringify([descriptor]))).toEqual([descriptor]);
  for (const bad of [{ ...descriptor, key: '../other.png' }, { ...descriptor, key: 'https://outside.test/pic.png' }, { ...descriptor, type: 'text/html' }, { ...descriptor, name: 'script.svg' }]) {
    expect(parseFeedbackAttachments([bad])).toEqual([]);
    expect(() => parseFeedbackAttachments([bad], { strict: true })).toThrow('invalid attachment');
  }
});
it('uploads through a CSOM object path with base64 Content, no return-value method action and no item read-back', () => {
  const descriptor = attachmentDescriptor(file());
  const doc = new DOMParser().parseFromString(feedbackFileRequest('Private & Files', descriptor, btoa('bytes')), 'text/xml');
  expect(doc.querySelector('parsererror,Query,ObjectIdentityQuery')).toBeNull();
  expect(doc.querySelector('Actions > Method')).toBeNull();
  expect(doc.querySelector('Actions > ObjectPath').getAttribute('ObjectPathId')).toBe('7');
  expect(doc.querySelector('Method[Name=Add] Parameter').getAttribute('TypeId')).toBe('{f5c8173c-cae6-4469-a7af-3879ca3c617c}');
  expect(doc.querySelector('Property[Name=Content]').getAttribute('Type')).toBe('Base64Binary');
  expect(doc.querySelector('Property[Name=Overwrite]').textContent).toBe('false');
  expect(doc.querySelector('Property[Name=Url]').textContent).toBe(descriptor.key);
});
it('reuses confirmed uploads on an explicit retry of the same draft without rereading or overwriting files', async () => {
  const store = { post: vi.fn().mockResolvedValue([{ ErrorInfo: null }]), get: vi.fn() };
  const upload = feedbackFileUploader(store), selected = [file()];
  const first = await upload(selected), retry = await upload(selected);
  expect(first).toEqual(retry);
  expect(store.post).toHaveBeenCalledTimes(1);
  expect(store.get).not.toHaveBeenCalled();
});
it.each([undefined, [], [{ ErrorInfo: {} }], [{ ErrorInfo: { ErrorMessage: 'Access denied' } }]])('does not replay uncertain file creations (%j)', async response => {
  const store = { post: vi.fn().mockResolvedValue(response) }, upload = feedbackFileUploader(store), selected = [file()];
  await expect(upload(selected)).rejects.toThrow('No report was sent');
  await expect(upload(selected)).rejects.toThrow('upload was not confirmed');
  expect(store.post).toHaveBeenCalledTimes(1);
});
it('explains the required owner upgrade when the attachment library does not exist, and allows retry after setup', async () => {
  const missing = [{ ErrorInfo: { ErrorTypeName: 'System.ArgumentException', ErrorCode: -2147024809, ErrorMessage: "List 'Modernization-Tracker - Feedback Attachments' does not exist at site with URL 'https://example.test'" } }];
  const store = { post: vi.fn().mockResolvedValueOnce(missing).mockResolvedValueOnce([{ ErrorInfo: null }]) }, upload = feedbackFileUploader(store), selected = [file()];
  await expect(upload(selected)).rejects.toThrow('enable private attachments');
  expect(await upload(selected)).toHaveLength(1);
  expect(store.post).toHaveBeenCalledTimes(2);
});
it('bounds the final base64 request below the CSOM limit', async () => {
  const store = { post: vi.fn().mockResolvedValue([{ ErrorInfo: null }]) };
  await feedbackFileUploader(store)([file('limit.png', new Uint8Array(MAX_FEEDBACK_FILE_BYTES))]);
  expect(new TextEncoder().encode(store.post.mock.calls[0][1].body).length).toBeLessThan(2 * 1024 * 1024);
});
it.each(['/sites/other/Files', '//outside.test/files', '/sites/tracker/../other', '/sites/tracker/Files?query=1'])('rejects a returned library root outside the current site: %s', async root => {
  const store = { webUrl: 'https://example.test/sites/tracker', get: vi.fn().mockResolvedValue({ ServerRelativeUrl: root }) };
  await expect(feedbackFilePath(store, attachmentDescriptor(file()))).rejects.toThrow('unexpected feedback attachment');
});
