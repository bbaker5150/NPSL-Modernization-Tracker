import { validateAttachment } from './taskAttachments';

export function referencePath(entries, id) {
  const parts = [];
  const seen = new Set();
  while (id && !seen.has(id)) {
    seen.add(id);
    const row = entries.find((entry) => entry.id === id);
    if (!row) break;
    parts.unshift(row);
    id = row.parentId;
  }
  return parts;
}

export function validateReference(entries, row, file) {
  const existing = row.id && entries.find((entry) => entry.id === row.id);
  if (row.id && !existing) throw new Error('This reference item no longer exists.');
  const next = existing ? { ...existing, name: row.name, parentId: row.parentId || '' } : { name: row.name, parentId: row.parentId || '', kind: file ? 'file' : 'folder' };
  next.name = String(next.name || '').trim();
  if (!next.name || next.name.length > 128 || /[\\/\x00-\x1f<>:"|?*]/.test(next.name) || /[. ]$/.test(next.name)) throw new Error('Enter a name of 128 characters or fewer without reserved characters or trailing dots.');
  if (next.parentId && !entries.some((entry) => entry.id === next.parentId && entry.kind === 'folder')) throw new Error('Choose an existing folder.');
  if (existing && referencePath(entries, next.parentId).some((entry) => entry.id === existing.id)) throw new Error('A folder cannot be moved inside itself.');
  if (entries.some((entry) => entry.id !== existing?.id && entry.parentId === next.parentId && entry.name.toLowerCase() === next.name.toLowerCase())) throw new Error('An item with this name already exists in that folder.');
  if (file) {
    if (existing) throw new Error('Upload a new document instead of replacing an existing reference.');
    validateAttachment(file);
    next.fileName = file.name;
    next.size = file.size;
  }
  if (existing?.kind === 'file' && next.name.split('.').pop() !== existing.fileName.split('.').pop()) throw new Error('Keep the document’s original file extension.');
  return next;
}
