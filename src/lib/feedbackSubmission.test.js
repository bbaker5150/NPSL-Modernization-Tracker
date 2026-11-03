import { beforeEach, expect, it, vi } from 'vitest';
import { feedbackCreateRequest, submitPrivateFeedback } from './feedbackSubmission';
import { createFeedback, validateFeedback } from './feedback';
import { SharePointStore } from './spStore';
import { resetWebUrlCache } from './spContext';
const report = { title: 'Updates <section> & "files"', kind: 'Feature', details: 'Let owners edit.\nKeep apostrophes: owner\'s and emoji 🔧' };
beforeEach(resetWebUrlCache);
it('builds a write-only CSOM request with literal text and no load, query or forged author', () => {
  const doc = new DOMParser().parseFromString(feedbackCreateRequest('Tracker & Feedback', report), 'text/xml');
  expect(doc.querySelector('parsererror')).toBeNull();
  expect(doc.querySelector('Query, ObjectIdentityQuery')).toBeNull();
  expect([...doc.querySelectorAll('Actions > Method')].map(node => node.getAttribute('Name'))).toEqual(['SetFieldValue', 'SetFieldValue', 'SetFieldValue', 'Update']);
  const values = Object.fromEntries([...doc.querySelectorAll('Actions > Method[Name=SetFieldValue]')].map(node => [...node.querySelectorAll('Parameter')].map(p => p.textContent)));
  expect(values).toEqual({ Title: report.title, FeedbackKind: report.kind, FeedbackDetails: report.details });
  expect(doc.querySelector('Method[Name=GetByTitle] Parameter').textContent).toBe('Tracker & Feedback');
  expect(doc.querySelector('Method[Name=AddItem] Parameter').getAttribute('TypeId')).toBe('{54cdbee5-0897-44ac-829f-411557fa11be}');
  expect(() => validateFeedback({ ...report, details: 'Invalid\u0001text' })).toThrow();
});
it('uses the signed-in cookie and form digest without REST item read-back or catalog reads', async () => {
  const fetchImpl = vi.fn(async (url, options) => {
    if (url.endsWith('/_api/contextinfo')) return new Response(JSON.stringify({ FormDigestValue: 'test-digest' }));
    if (url.endsWith('/_vti_bin/client.svc/ProcessQuery')) {
      expect(options.credentials).toBe('include');
      expect(options.headers).toMatchObject({ 'X-RequestDigest': 'test-digest', 'Content-Type': 'text/xml; charset=utf-8' });
      return new Response(JSON.stringify([{ ErrorInfo: null }]));
    }
    // This is the screenshot failure when the add-only caller requests a REST item response.
    return new Response(JSON.stringify({ error: { message: 'Item does not exist. It may have been deleted by another user.' } }), { status: 500 });
  });
  const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/tracker', fetchImpl });
  await createFeedback(store).submit(report);
  expect(fetchImpl.mock.calls.map(([url]) => new URL(url).pathname)).toEqual(['/sites/tracker/_api/contextinfo', '/sites/tracker/_vti_bin/client.svc/ProcessQuery']);
});
it.each([
  [{ ErrorInfo: { ErrorMessage: 'Item does not exist. It may have been deleted by another user.', ErrorCode: -2147024809, ErrorTypeName: 'System.ArgumentException' } }],
  [{ ErrorInfo: { ErrorMessage: 'Access denied.', ErrorCode: -2147024891 } }],
  {}, [], [{ ErrorInfo: undefined }], [{ Other: null }],
])('never retries or claims success on an ambiguous CSOM response: %j', async response => {
  const store = { post: vi.fn().mockResolvedValue(response) };
  await expect(submitPrivateFeedback(store, ['Current', 'Legacy'], report)).rejects.toThrow(/check the inbox/);
  expect(store.post).toHaveBeenCalledTimes(1);
});
it.each([403, 404, 500])('never retries a transport error %s', async status => {
  const store = { post: vi.fn().mockRejectedValue(Object.assign(new Error('Item does not exist'), { status })) };
  await expect(submitPrivateFeedback(store, ['Current', 'Legacy'], report)).rejects.toThrow('before you retry');
  expect(store.post).toHaveBeenCalledTimes(1);
});
