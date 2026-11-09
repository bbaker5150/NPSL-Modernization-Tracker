import { TrackerLists, legacyListTitle, listTitleAliases } from './trackerLists';
import { submitPrivateFeedback } from './feedbackSubmission';
import { TRACKER_GROUPS } from './trackerGroups';
import { isManager } from './access';
import { feedbackSetup } from './feedbackSetup';
import { ensureFeedbackStatus, FEEDBACK_STATUSES } from './feedbackFields';
import { attachmentDescriptor, feedbackFileUploader, feedbackFilePath, parseFeedbackAttachments, validateFeedbackFiles, MAX_FEEDBACK_FILE_BYTES } from './feedbackAttachments';
import { feedbackAttachmentsSetup } from './feedbackAttachmentsSetup';
// Feedback is deliberately separate from shared tracker data and exports.
const unwrap = value => value?.d || value;
const rows = value => value?.value || value?.d?.results || [];
export function validateFeedback(input) {
  const title = String(input.title || '').trim();
  const details = String(input.details || '').trim();
  if (!['Bug', 'Feature'].includes(input.kind)) throw new Error('Choose a bug report or feature request.');
  if (!title || title.length > 200) throw new Error('Enter a summary of 1–200 characters.');
  if (!details || details.length > 10000) throw new Error('Enter details of 1–10,000 characters.');
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/.test(title + details)) throw new Error('Remove unsupported control characters from your feedback.');
  return { kind: input.kind, title, details };
}
export function createFeedback(store) {
  const title = legacyListTitle(store.prefix, 'feedback');
  const catalog = new TrackerLists(store.prefix, path => store.get(path), (path, options) => store.post(path, options));
  // Local mode is a disposable demonstration, never a private data store.
  const preview = [];
  const local = !store.webUrl;
  async function access() {
    const user = await store.currentUser();
    if (user.isSiteAdmin) return { owner: true, review: true };
    if (local) {
      const data = await store.load?.();
      return { owner: false, review: isManager(user, data?.users) };
    }
    // Match the tracker's site-owner definition: Manage Permissions on the WEB.
    // Effective permissions include direct grants and security-group membership;
    // currentuser/groups alone can omit these. List/project permissions and app
    // Manager roles must never qualify a user as a site owner.
    let permissionError;
    try {
      const response = unwrap(await store.get('/_api/web/effectivebasepermissions'));
      const bits = response?.EffectiveBasePermissions || response;
      if ((BigInt(bits?.Low || 0) & 33554432n) !== 0n) return { owner: true, review: true };
    } catch (error) { permissionError = error; }
    const groups = rows(await store.get('/_api/web/currentuser/groups?$select=Id,Title&$top=5000'));
    const manager = groups.some(group => group.Title === TRACKER_GROUPS.Manager);
    try {
      const group = unwrap(await store.get('/_api/web/associatedownergroup?$select=Id'));
      if (Number.isInteger(group?.Id) && groups.some(entry => entry.Id === group.Id)) return { owner: true, review: true };
    } catch (error) { permissionError ||= error; }
    if (permissionError && !manager) throw permissionError;
    return { owner: false, review: manager };
  }
  const isOwner = async () => (await access()).owner;
  const setup = feedbackSetup(store, title, isOwner);
  const attachmentSetup = feedbackAttachmentsSetup(store, isOwner, () => catalog.path('feedback'), () => setup.status());
  setup.attachmentsStatus = () => attachmentSetup.status();
  setup.enableAttachments = progress => attachmentSetup.provision(progress);
  const uploadFiles = feedbackFileUploader(store);
  const uncertainReportFiles = new WeakSet();
  const localFiles = new Map();
  return {
    local,
    access,
    isOwner,
    setup,
    async submit(input, progress = () => {}) {
      const value = validateFeedback(input);
      const files = validateFeedbackFiles(input.attachments || []);
      if (local) {
        const attachments = files.map(file => { const attachment = attachmentDescriptor(file); localFiles.set(attachment.key, file); return attachment; });
        preview.unshift({ ...value, attachments, status: 'New', etag: '1', id: crypto.randomUUID(), created: new Date().toISOString(), author: 'Local preview' }); return;
      }
      if (files.some(file => uncertainReportFiles.has(file))) throw new Error('A report using these files may already be in Feedback History. Ask an owner to check before starting another report. This draft has not been sent again.');
      value.attachments = await uploadFiles(files, progress);
      progress('Sending feedback…');
      const aliases = listTitleAliases(store.prefix, 'feedback');
      try { await submitPrivateFeedback(store, [aliases[1], aliases[0], ...aliases.slice(2)], value); uploadFiles.committed(files); }
      catch (error) {
        // Reusing these immutable keys in a duplicate report would make deleting
        // either report remove the other's files. Never replay an uncertain create.
        if (error.submissionUncertain) files.forEach(file => uncertainReportFiles.add(file));
        throw error;
      }
    },
    async downloadAttachment(id, key) {
      if (!(await access()).review) throw new Error('Only site owners and Tracker Managers can read feedback attachments.');
      if (local ? typeof id !== 'string' || !id : !Number.isSafeInteger(id) || id <= 0) throw new Error('Select a valid feedback submission.');
      const record = local ? preview.find(item => item.id === id) : unwrap(await store.get(`${await catalog.path('feedback')}/items(${id})?$select=*`));
      const attachment = parseFeedbackAttachments(local ? record?.attachments : record?.FeedbackAttachments, { strict: true }).find(file => file.key === key);
      if (!attachment) throw new Error('This attachment is no longer linked to the selected submission. Reopen Feedback History.');
      if (local) return { blob: localFiles.get(key), attachment };
      // Match normal Tracker downloads: embedded hosts may stall binary REST
      // $value requests. The direct private file URL still enforces SharePoint ACLs.
      const url = await feedbackFilePath(store, attachment, { download: true });
      const response = await (store.fetchImpl || fetch)(url, { credentials: 'include', cache: 'no-store', headers: { Accept: '*/*' } });
      if (!response.ok) throw new Error(`The attachment could not be downloaded (${response.status}). Check your feedback library access with a site owner.`);
      if (/text\/html/i.test(response.headers?.get('Content-Type') || '')) throw new Error('SharePoint returned a sign-in page instead of the attachment. Refresh your session and try again.');
      const blob = await response.blob();
      // Office property promotion can change library file bytes after upload.
      // Bound downloads without treating the original size as an integrity hash.
      if (!blob.size || blob.size > MAX_FEEDBACK_FILE_BYTES * 2) throw new Error('The stored attachment is empty or unexpectedly large. Ask a site owner to verify the file.');
      return { blob: new Blob([blob], { type: attachment.type }), attachment };
    },
    async update(id, input, etag) {
      const { owner, review } = await access();
      if (!review) throw new Error('Only site owners and Tracker Managers can update submissions.');
      if (local ? typeof id !== 'string' || !id : !Number.isSafeInteger(id) || id <= 0) throw new Error('Select a valid feedback submission to update.');
      const value = validateFeedback(input);
      if (!FEEDBACK_STATUSES.includes(input.status)) throw new Error('Choose a valid feedback status.');
      if (typeof etag !== 'string' || !etag || etag === '*') throw new Error('Reopen Feedback History to load the latest submission before editing.');
      if (local) {
        const index = preview.findIndex(item => item.id === id);
        if (index < 0) throw new Error('This submission was deleted. Reopen Feedback History.');
        if (preview[index].etag !== etag) throw new Error('This submission changed. Reopen Feedback History before editing again.');
        preview[index] = { ...preview[index], ...value, status: input.status, etag: String(Number(etag) + 1) };
        return;
      }
      const path = await catalog.path('feedback');
      await ensureFeedbackStatus(store, path, { allowCreate: owner });
      try {
        await store.post(`${path}/items(${id})`, { body: { Title: value.title, FeedbackKind: value.kind, FeedbackDetails: value.details, FeedbackStatus: input.status }, headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': etag } });
      } catch (error) {
        if (error.status === 412) throw new Error('Another reviewer updated this submission. Your draft is still here; copy any changes and reopen Feedback History to review the latest version.');
        if (error.status === 404) throw new Error('This submission was deleted. Your draft is still here; reopen Feedback History to refresh.');
        throw error;
      }
    },
    async clearInbox(ids, onProgress = () => {}) {
      if (!await isOwner()) throw new Error('Only site owners can clear submissions.');
      if (!Array.isArray(ids) || !ids.length || ids.some(id => local ? typeof id !== 'string' || !id : !Number.isSafeInteger(id) || id <= 0)) throw new Error('Select valid feedback submissions to clear.');
      // Clear the reviewed snapshot only. Never query and delete new arrivals.
      const selected = [...new Set(ids)], clearedIds = [];
      const path = local ? null : await catalog.path('feedback');
      for (const id of selected) {
        try {
          if (local) {
            const index = preview.findIndex(item => item.id === id);
            if (index !== -1) { preview[index].attachments?.forEach(file => localFiles.delete(file.key)); preview.splice(index, 1); }
          } else {
            // Recycle linked files first. If one fails, retain the report and its
            // manifest so the owner can find/recover every file before retrying.
            let record;
            try { record = unwrap(await store.get(`${path}/items(${id})?$select=*`)); }
            catch (error) { if (error.status !== 404) throw error; }
            for (const file of parseFeedbackAttachments(record?.FeedbackAttachments, { strict: true })) {
              try { await store.post(`${await feedbackFilePath(store, file)}/recycle()`); }
              catch (error) { if (error.status !== 404) throw new Error(`A linked attachment could not be recycled. The submission was kept; restore any already recycled files if needed. ${error.message}`); }
            }
            try { await store.post(`${path}/items(${id})/recycle()`); }
            catch (error) { if (error.status !== 404) throw error; }
          }
          clearedIds.push(id);
          onProgress({ cleared: clearedIds.length, total: selected.length });
        } catch (cause) {
          const error = new Error(`Cleared ${clearedIds.length} of ${selected.length} submissions. ${cause.message} Reopen Feedback History to check its current contents before retrying.`);
          error.clearedIds = [...clearedIds];
          throw error;
        }
      }
      return { clearedIds };
    },
    async list() {
      if (!(await access()).review) throw new Error('Only site owners and Tracker Managers can view submissions.');
      if (local) return preview.map(item => ({ ...item }));
      const path = await catalog.path('feedback');
      // Select * so older lists without FeedbackStatus remain readable before upgrade.
      const result = []; let next = `${path}/items?$select=*,Author/Title&$expand=Author&$orderby=Created desc&$top=100`;
      while (next) {
        const body = await store.get(next, { Accept: 'application/json;odata=verbose' });
        result.push(...rows(body).map(row => ({ id: row.Id, title: row.Title, kind: row.FeedbackKind, details: row.FeedbackDetails, status: row.FeedbackStatus || 'New', etag: row['@odata.etag'] || row['odata.etag'] || row.__metadata?.etag, created: row.Created, author: row.Author?.Title || 'Unknown', attachments: parseFeedbackAttachments(row.FeedbackAttachments) })));
        const link = body['@odata.nextLink'] || body['odata.nextLink'] || body.d?.__next;
        if (!link) break;
        const url = new URL(link, `${store.webUrl}/`);
        if (!url.href.startsWith(`${store.webUrl}/_api/`)) throw new Error('Unexpected feedback pagination address.');
        next = url.href.slice(store.webUrl.length);
      }
      return result;
    },
  };
}
