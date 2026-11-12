import { displayListTitle } from './trackerLists';

// CSOM's base64 Content upload has a documented 2 MB request limit. Keeping
// each raw file at 1 MiB leaves room for XML and base64 expansion.
export const MAX_FEEDBACK_FILE_BYTES = 1024 * 1024;
export const MAX_FEEDBACK_FILES = 5;
const TYPES = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', pdf: 'application/pdf', txt: 'text/plain', log: 'text/plain', csv: 'text/csv', doc: 'application/msword', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xls: 'application/vnd.ms-excel', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', ppt: 'application/vnd.ms-powerpoint', pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', zip: 'application/zip' };
export const FEEDBACK_FILE_ACCEPT = Object.keys(TYPES).map(extension => `.${extension}`).join(',');
export const feedbackAttachmentTitle = prefix => `${displayListTitle(prefix, 'feedback')} Attachments`;
export const attachmentListPath = prefix => `/_api/web/lists/getbytitle('${encodeURIComponent(feedbackAttachmentTitle(prefix).replace(/'/g, "''"))}')`;
export const attachmentSize = size => size >= 1024 * 1024 ? `${(size / 1024 / 1024).toFixed(1)} MB` : `${Math.max(1, Math.round(size / 1024))} KB`;
const extension = name => name.split('.').pop().toLowerCase();
const xml = value => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');

export function validateFeedbackFiles(files = []) {
  if (!Array.isArray(files) || files.length > MAX_FEEDBACK_FILES) throw new Error('Attach up to 5 files per submission.');
  for (const file of files) {
    if (!file || typeof file.name !== 'string' || !file.name.trim() || file.name.length > 128 || /[\\/\u0000-\u001f<>:"|?*]/.test(file.name) || /[. ]$/.test(file.name)) throw new Error('Use a file name of 1–128 characters without special path characters.');
    if (!Object.hasOwn(TYPES, extension(file.name))) throw new Error(`${file.name}: choose an image, PDF, Office document, text file, CSV, or ZIP.`);
    if (!Number.isSafeInteger(file.size) || file.size <= 0 || file.size > MAX_FEEDBACK_FILE_BYTES) throw new Error(`${file.name}: files must be nonempty and no larger than 1 MB (1,048,576 bytes).`);
  }
  return files;
}
export function attachmentDescriptor(file) {
  validateFeedbackFiles([file]);
  const ext = extension(file.name);
  return { key: `${crypto.randomUUID()}.${ext}`, name: file.name, size: file.size, type: TYPES[ext] };
}
export function parseFeedbackAttachments(value, { strict = false } = {}) {
  if (value == null || value === '') return [];
  try {
    const files = typeof value === 'string' ? JSON.parse(value) : value;
    validateFeedbackFiles(files);
    if (files.some(file => !/^[\da-f]{8}(?:-[\da-f]{4}){3}-[\da-f]{12}\.[a-z0-9]+$/i.test(file.key || '') || extension(file.key) !== extension(file.name) || TYPES[extension(file.name)] !== file.type) || new Set(files.map(file => file.key)).size !== files.length) throw new Error('Invalid attachment manifest');
    return files.map(({ key, name, size, type }) => ({ key, name, size, type }));
  } catch {
    if (strict) throw new Error('This submission has invalid attachment details. Ask a site owner to check its FeedbackAttachments field before deleting it.');
    return [];
  }
}
export function feedbackFileRequest(title, attachment, base64) {
  parseFeedbackAttachments([attachment], { strict: true });
  if (typeof base64 !== 'string' || !/^[A-Za-z0-9+/]*={0,2}$/.test(base64)) throw new Error('Could not encode the attachment.');
  return `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="Modernization Tracker"><Actions><ObjectPath Id="8" ObjectPathId="7" /></Actions><ObjectPaths><StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="1" ParentId="0" Name="Web" /><Property Id="2" ParentId="1" Name="Lists" /><Method Id="3" ParentId="2" Name="GetByTitle"><Parameters><Parameter Type="String">${xml(title)}</Parameter></Parameters></Method><Property Id="5" ParentId="3" Name="RootFolder" /><Property Id="6" ParentId="5" Name="Files" /><Method Id="7" ParentId="6" Name="Add"><Parameters><Parameter TypeId="{f5c8173c-cae6-4469-a7af-3879ca3c617c}"><Property Name="Content" Type="Base64Binary">${base64}</Property><Property Name="Overwrite" Type="Boolean">false</Property><Property Name="Url" Type="String">${xml(attachment.key)}</Property></Parameter></Parameters></Method></ObjectPaths></Request>`;
}
async function base64File(file) {
  const buffer = typeof file.arrayBuffer === 'function' ? await file.arrayBuffer() : await new Promise((resolve, reject) => {
    const reader = new FileReader(); reader.onload = () => resolve(reader.result); reader.onerror = () => reject(new Error(`Could not read ${file.name}. Choose the file again.`)); reader.readAsArrayBuffer(file);
  });
  let binary = '';
  const bytes = new Uint8Array(buffer);
  for (let start = 0; start < bytes.length; start += 8192) binary += String.fromCharCode(...bytes.subarray(start, start + 8192));
  return btoa(binary);
}

// Upload before creating the report, so a failed file never creates a second
// report on retry. Confirmed files are reused in the same draft. An uncertain
// upload is not replayed: its immutable key may already exist in SharePoint.
export function feedbackFileUploader(store) {
  const receipts = new WeakMap();
  const upload = async (files, progress = () => {}) => {
    validateFeedbackFiles(files);
    const result = [];
    for (const [index, file] of files.entries()) {
      let receipt = receipts.get(file);
      if (!receipt) { receipt = { attachment: attachmentDescriptor(file), state: 'new' }; receipts.set(file, receipt); }
      if (receipt.state === 'uncertain') throw new Error(`${file.name}: its upload was not confirmed. Ask a site owner to check Feedback Attachments before selecting this file again. No report was sent.`);
      if (receipt.state !== 'uploaded') {
        progress(`Uploading file ${index + 1} of ${files.length}…`);
        const body = feedbackFileRequest(feedbackAttachmentTitle(store.prefix), receipt.attachment, await base64File(file));
        if (new TextEncoder().encode(body).length >= 2 * 1024 * 1024) throw new Error(`${file.name}: the encoded upload is too large.`);
        receipt.state = 'uncertain';
        try {
          const response = await store.post('/_vti_bin/client.svc/ProcessQuery', { raw: true, headers: { 'Content-Type': 'text/xml; charset=utf-8', Accept: 'application/json' }, body });
          const header = Array.isArray(response) && response[0];
          if (!header || header.ErrorInfo !== null) {
            const problem = header?.ErrorInfo;
            if (problem?.ErrorTypeName === 'System.ArgumentException' && problem.ErrorCode === -2147024809 && String(problem.ErrorMessage).startsWith(`List '${feedbackAttachmentTitle(store.prefix)}' does not exist at site with URL`)) {
              receipt.state = 'new';
              throw new Error('Ask a site owner to open feedback and enable private attachments. You can remove the files to send text feedback now.');
            }
            throw new Error(problem?.ErrorMessage || 'SharePoint returned no upload confirmation.');
          }
          receipt.state = 'uploaded';
        } catch (error) {
          if (receipt.state === 'new') throw error;
          throw new Error(`${file.name}: SharePoint did not confirm this upload. No report was sent. Ask a site owner to check Feedback Attachments before retrying. ${error.message}`);
        }
      }
      result.push(receipt.attachment);
    }
    return result;
  };
  // Once a report is confirmed, these files belong to that report. A later
  // report using the same File object must upload under new immutable keys.
  upload.committed = files => files.forEach(file => receipts.delete(file));
  return upload;
}

export async function feedbackFilePath(store, descriptor, { download = false } = {}) {
  parseFeedbackAttachments([descriptor], { strict: true });
  const response = await store.get(`${attachmentListPath(store.prefix)}/RootFolder?$select=ServerRelativeUrl`);
  const root = (response?.d || response)?.ServerRelativeUrl;
  const site = new URL(store.webUrl);
  const webPath = decodeURIComponent(site.pathname).replace(/\/$/, '');
  if (typeof root !== 'string' || !root.startsWith(`${webPath}/`) || decodeURIComponent(root).includes('..') || /[?#\\]/.test(root)) throw new Error('SharePoint returned an unexpected feedback attachment location.');
  if (download) {
    const url = new URL(`${root}/${encodeURIComponent(descriptor.key)}`, site.origin);
    if (url.origin !== site.origin || !decodeURIComponent(url.pathname).startsWith(`${webPath}/`)) throw new Error('SharePoint returned an unexpected feedback attachment address.');
    return url.href;
  }
  return `/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(`${root}/${descriptor.key}`.replace(/'/g, "''"))}')`;
}
