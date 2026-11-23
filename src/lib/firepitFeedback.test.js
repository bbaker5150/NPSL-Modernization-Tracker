import { afterEach, expect, it, vi } from 'vitest';
import { createFeedback } from './feedback';
import { submitPrivateFeedback } from './feedbackSubmission';
import { feedbackSetup } from './feedbackSetup';
import { feedbackAttachmentsSetup } from './feedbackAttachmentsSetup';

const report = { kind: 'Bug', title: 'Save failed', details: 'Steps to reproduce' };
const store = () => ({
  webUrl: 'https://tenant.sharepoint-mil.us/sites/tracker',
  prefix: 'Modernization',
  get: vi.fn(),
  post: vi.fn().mockResolvedValue([{ ErrorInfo: null }]),
});
function firepit() {
  window.firepit = { sharepoint: { siteUrl: 'https://tenant.sharepoint-mil.us/sites/tracker', fetch: vi.fn() } };
}
afterEach(() => { delete window.firepit; });

it('does not provision unusable private feedback storage or alter ACLs in Firepit', async () => {
  firepit();
  const raw = store(), owner = vi.fn(), verify = vi.fn();
  await expect(feedbackSetup(raw, 'Feedback', owner).provision([1])).rejects.toThrow('No feedback or attachments have been sent');
  await expect(feedbackAttachmentsSetup(raw, owner, vi.fn(), verify).provision()).rejects.toThrow('No feedback or attachments have been sent');
  expect(raw.get).not.toHaveBeenCalled(); expect(raw.post).not.toHaveBeenCalled();
  expect(owner).not.toHaveBeenCalled(); expect(verify).not.toHaveBeenCalled();
});

it.each([{ attachments: [] }, { attachments: [new File(['steps'], 'steps.txt', { type: 'text/plain' })] }])('blocks private Firepit feedback before uploading any files or creating a report: %j', async ({ attachments }) => {
  firepit();
  const raw = store(), progress = vi.fn();
  const submit = createFeedback(raw).submit({ ...report, attachments }, progress);
  await expect(submit).rejects.toThrow('SharePoint feedback form outside Firepit');
  await expect(submit).rejects.not.toHaveProperty('submissionUncertain');
  expect(raw.get).not.toHaveBeenCalled();
  expect(raw.post).not.toHaveBeenCalled();
  expect(window.firepit.sharepoint.fetch).not.toHaveBeenCalled();
  expect(progress).not.toHaveBeenCalled();
});

it('also blocks direct private submission before CSOM or a REST retry', async () => {
  firepit();
  const raw = store();
  await expect(submitPrivateFeedback(raw, ['Feedback', 'Legacy Feedback'], report)).rejects.toThrow('No feedback or attachments have been sent');
  expect(raw.post).not.toHaveBeenCalled();
});

it('retains write-only CSOM submission outside Firepit', async () => {
  const raw = store();
  await createFeedback(raw).submit(report);
  expect(raw.post).toHaveBeenCalledExactlyOnceWith('/_vti_bin/client.svc/ProcessQuery', expect.objectContaining({ raw: true }));
  expect(raw.get).not.toHaveBeenCalled();
});

it('keeps disposable local preview available without calling SharePoint', async () => {
  firepit();
  const raw = { ...store(), webUrl: '' };
  await expect(createFeedback(raw).submit(report)).resolves.toBeUndefined();
  expect(raw.post).not.toHaveBeenCalled();
  expect(window.firepit.sharepoint.fetch).not.toHaveBeenCalled();
});

it('downloads only a verified feedback manifest attachment through the authenticated Firepit REST bridge', async () => {
  firepit();
  const raw = store();
  const key = '11111111-1111-4111-8111-111111111111.png';
  const attachment = { key, name: 'screen.png', size: 6, type: 'image/png' };
  raw.currentUser = vi.fn().mockResolvedValue({ isSiteAdmin: true });
  raw.fetchImpl = vi.fn();
  raw.get.mockImplementation(async path => {
    if (path.includes('/lists?')) return { value: [{ Id: '836fb5d6-3f10-4e6b-92dc-3aa8e297dfee', Title: 'ModernizationFeedback', BaseTemplate: 100 }] };
    if (path.includes('/items(1)')) return { FeedbackAttachments: JSON.stringify([attachment]) };
    if (path.includes('/RootFolder?')) return { ServerRelativeUrl: '/sites/tracker/FeedbackFiles' };
    throw new Error(`Unexpected read: ${path}`);
  });
  window.firepit.sharepoint.fetch.mockResolvedValue({ ok: true, blob: async () => new Blob(['pixels']) });
  const feedback = createFeedback(raw);
  expect((await feedback.downloadAttachment(1, key)).attachment).toEqual(attachment);
  expect(window.firepit.sharepoint.fetch).toHaveBeenCalledWith(
    `${raw.webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='%2Fsites%2Ftracker%2FFeedbackFiles%2F${key}')/$value`, expect.any(Object),
  );
  expect(raw.fetchImpl).not.toHaveBeenCalled();
  await expect(feedback.downloadAttachment(1, '22222222-2222-4222-8222-222222222222.png')).rejects.toThrow('no longer linked');
  expect(window.firepit.sharepoint.fetch).toHaveBeenCalledTimes(1);
});
