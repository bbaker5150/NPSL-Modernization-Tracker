export const MAX_ATTACHMENT_BYTES = 50 * 1024 * 1024;
export function validateAttachmentName(name) {
  if (typeof name !== 'string' || !name.trim()) throw new Error('Enter a document name.');
  if (name.length > 128 || /[\\/\x00-\x1f<>:"|?*]/.test(name) || /[. ]$/.test(name)) throw new Error('Use a file name of 128 characters or fewer without reserved characters or trailing dots/spaces.');
  return name.trim();
}
export function validateAttachment(file) {
  if (!file || typeof file.name !== 'string' || !file.name.trim()) throw new Error('Choose a file to attach.');
  validateAttachmentName(file.name);
  if (!Number.isFinite(file.size) || file.size <= 0 || file.size > MAX_ATTACHMENT_BYTES) throw new Error('Attachments must be nonempty and no larger than 50 MB each.');
}

// Local preview keeps file blobs separate from the JSON portfolio.
async function attachmentDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open('modernization-task-documents', 1);
    request.onupgradeneeded = () => request.result.createObjectStore('files', { keyPath: 'key' });
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(new Error('Local attachment storage is unavailable.'));
  });
}
export async function localAttachments(taskId, file, deleteName) {
  const db = await attachmentDb();
  try {
    return await new Promise((resolve, reject) => {
      const tx = db.transaction('files', file || deleteName ? 'readwrite' : 'readonly');
      const store = tx.objectStore('files');
      const request = deleteName ? store.delete(`${taskId}/${deleteName.toLowerCase()}`) : file ? store.add({ key: `${taskId}/${file.name.toLowerCase()}`, taskId, name: file.name, blob: file }) : store.getAll();
      let result;
      request.onsuccess = () => { result = request.result; };
      tx.oncomplete = () => resolve(file || deleteName ? undefined : result.filter((row) => row.taskId === taskId));
      tx.onerror = () => reject(new Error(file ? 'Could not store the file. A document with this name may already exist, or browser storage may be full.' : 'Could not read local attachments.'));
      tx.onabort = tx.onerror;
    });
  } finally { db.close(); }
}
