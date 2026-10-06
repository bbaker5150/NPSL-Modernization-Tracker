import JSZip from 'jszip';
import { uploadGuide } from './guide.js';

export const SOURCE = 'https://flankspeed.sharepoint-mil.us/sites/ISEAMETENG';
export const TARGET = 'https://flankspeed.sharepoint-mil.us/sites/metsoft/SitePages/Modernization-Tracker.aspx';
const TITLES = ['ModernizationReferenceDocuments', 'NPSL Tracker - Reference Documents', 'Modernization-Tracker - Reference Documents'];
const unwrap = body => body.d || body;
const collection = body => body.value || body.d?.results || body.results;
export function sourceUrl(value) {
  const url = new URL(value, `${SOURCE}/`);
  const source = new URL(SOURCE);
  if (url.origin !== source.origin || !url.pathname.toLowerCase().startsWith(`${source.pathname.toLowerCase()}/`) || url.username || url.password) throw new Error('Refusing a request outside the old ISEA METENG site.');
  return url.href;
}

// Bound the entire operation, including body decoding. Some host bridges ignore
// AbortSignal; the race still settles the UI and ignores late results.
export async function readSource(url, { fetchImpl = fetch, signal, timeoutMs = 120000, binary = false } = {}) {
  const address = sourceUrl(url);
  const controller = new AbortController();
  let timer, cancel;
  const cancelled = new Promise((_, reject) => {
    cancel = () => { controller.abort(); reject(new Error('Stopped. Downloaded files remain available while this page stays open.')); };
    if (signal?.aborted) cancel();
    else signal?.addEventListener('abort', cancel, { once: true });
    timer = setTimeout(() => { controller.abort(); reject(new Error(`Read timed out after ${timeoutMs / 1000} seconds. Retry, or download this original file and select it below.`)); }, timeoutMs);
  });
  const read = async () => {
    if (signal?.aborted) throw new Error('Stopped.');
    const response = await fetchImpl(address, { method: 'GET', credentials: 'include', cache: 'no-store', signal: controller.signal, headers: { Accept: binary ? 'application/octet-stream,*/*' : 'application/json;odata=nometadata' } });
    if (!response.ok) throw new Error(`SharePoint read failed (HTTP ${response.status}). Check access to the old site.`);
    if (response.url) sourceUrl(response.url);
    if (!binary) return response.json();
    const bytes = await response.arrayBuffer();
    const prefix = new TextDecoder().decode(bytes.slice(0, 2048)).trimStart();
    if (/^(?:<!doctype html|<html\b)/i.test(prefix) && !/\.html?(?:[?#]|$)/i.test(address)) throw new Error('SharePoint returned a sign-in or preview page instead of file bytes. Use Download original, then choose the downloaded file.');
    return bytes;
  };
  try { return await Promise.race([read(), cancelled]); }
  finally { clearTimeout(timer); signal?.removeEventListener('abort', cancel); }
}

export async function scanReferences(options = {}) {
  const filter = TITLES.map(title => `Title eq '${title}'`).join(' or ');
  const catalog = collection(await readSource(`${SOURCE}/_api/web/lists?$select=Id,Title,BaseTemplate&$filter=${encodeURIComponent(filter)}`, options));
  if (!Array.isArray(catalog) || catalog.length !== 1 || !TITLES.includes(catalog[0].Title) || catalog[0].BaseTemplate !== 100) throw new Error('Expected exactly one tracker Reference Documents list. Check access, list naming, or duplicate lists.');
  const list = catalog[0];
  if (!/^[\da-f]{8}(-[\da-f]{4}){3}-[\da-f]{12}$/i.test(list.Id)) throw new Error('SharePoint returned an invalid list ID.');
  const api = `${SOURCE}/_api/web/lists(guid'${list.Id}')`;
  const root = unwrap(await readSource(`${api}/RootFolder?$select=ServerRelativeUrl`, options)).ServerRelativeUrl;
  if (typeof root !== 'string' || !root.startsWith('/')) throw new Error('SharePoint did not return the reference list path.');
  const rows = [], visited = new Set();
  let next = `${api}/items?$select=Id,Title,RecordId,ReferenceParentId,EntryKind,FileName,FileSize,Archived,Modified,FileSystemObjectType,AttachmentFiles/FileName&$expand=AttachmentFiles&$top=500`;
  while (next) {
    next = sourceUrl(next);
    if (visited.has(next)) throw new Error('SharePoint repeated a catalog page. No files were omitted.');
    visited.add(next);
    const body = await readSource(next, options), page = collection(body);
    if (!Array.isArray(page)) throw new Error('SharePoint did not return the reference document records.');
    rows.push(...page.filter(row => Number(row.FileSystemObjectType) !== 1));
    next = body['@odata.nextLink'] || body['odata.nextLink'] || body.d?.__next;
  }
  return planReferences(rows, list, root);
}

const segment = value => {
  const clean = String(value || '').replace(/[\\/<>:"|?*\x00-\x1f]/g, '_').replace(/[. ]+$/g, '').replace(/^\.+$/, '_');
  return !clean ? 'unnamed' : /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(clean) ? `_${clean}` : clean;
};
export function planReferences(rows, list, root) {
  const byId = new Map(), files = [], folders = [], warnings = [], used = new Set();
  for (const row of rows) {
    if (!row.RecordId || byId.has(row.RecordId)) throw new Error(`Missing or duplicate reference identity at item ${row.Id}. Fix this before exporting.`);
    if (!['folder', 'file'].includes(row.EntryKind) || !Number.isInteger(row.Id) || row.Id <= 0) throw new Error(`Unrecognized reference record at item ${row.Id}.`);
    byId.set(row.RecordId, row);
  }
  const ancestors = row => {
    const result = [], seen = new Set([row.RecordId]);
    let id = row.ReferenceParentId;
    while (id) {
      const parent = byId.get(id);
      if (!parent || parent.EntryKind !== 'folder' || seen.has(id)) throw new Error(`Broken folder hierarchy at item ${row.Id}. No destination path was guessed.`);
      seen.add(id); result.unshift(parent); id = parent.ReferenceParentId;
    }
    return result;
  };
  // Use stable folder IDs to disambiguate names, including case-only collisions.
  const folderSegments = new Map();
  const siblings = new Map();
  for (const row of rows.filter(row => row.EntryKind === 'folder').sort((a, b) => a.Id - b.Id)) {
    const key = `${row.ReferenceParentId || ''}/${segment(row.Title).toLowerCase()}`;
    const name = siblings.has(key) ? `${segment(row.Title)}__item-${row.Id}` : segment(row.Title);
    siblings.set(key, true); folderSegments.set(row.RecordId, name);
  }
  for (const row of [...rows].sort((a, b) => a.Id - b.Id)) {
    const parents = ancestors(row), archived = !!row.Archived || parents.some(parent => parent.Archived);
    const destination = parents.map(parent => parent.Title);
    const path = [archived ? 'Archived' : 'Active', ...parents.map(parent => folderSegments.get(parent.RecordId))];
    if (row.EntryKind === 'folder') folders.push({ path: [...path, folderSegments.get(row.RecordId)].join('/'), destination: [...destination, row.Title], archived, sourceItemId: row.Id });
    const attachments = row.AttachmentFiles?.results || row.AttachmentFiles || [];
    if (!Array.isArray(attachments)) throw new Error(`Invalid attachment inventory at item ${row.Id}.`);
    if (row.EntryKind === 'file' && !attachments.length) {
      if (!archived) throw new Error(`Reference document has no attached file: ${row.Title} (item ${row.Id}). Restore the source attachment before exporting.`);
      warnings.push(`Archived item ${row.Id} (${row.Title}) has no remaining attachment.`);
    }
    for (const attachment of attachments) {
      const name = attachment.FileName;
      if (!name || /[\\/\x00-\x1f]/.test(name)) throw new Error(`Invalid attachment name at item ${row.Id}.`);
      const primary = row.EntryKind === 'file' && (name === row.FileName || (!row.FileName && attachments.length === 1));
      // Keep every attachment; additional attachments are explicitly marked for review.
      const zipName = segment(primary ? row.Title || name : name);
      let zipPath = [...path, zipName].join('/');
      if (used.has(zipPath.toLowerCase())) zipPath = [...path, `item-${row.Id}-${files.length + 1}-${zipName}`].join('/');
      used.add(zipPath.toLowerCase());
      const url = sourceUrl(`${new URL(SOURCE).origin}${root.split('/').map(encodeURIComponent).join('/')}/Attachments/${row.Id}/${encodeURIComponent(name)}`);
      const file = { key: `${row.Id}/${name}`, sourceItemId: row.Id, recordId: row.RecordId, sourceName: name, displayName: row.Title, destination, archived, extraAttachment: !primary, zipPath, url, expectedSize: primary && Number(row.FileSize) > 0 ? Number(row.FileSize) : null };
      if (!primary) warnings.push(`Item ${row.Id}: review additional attachment ${name} before uploading.`);
      files.push(file);
    }
  }
  // Reserve folder entries and reject file/folder name collisions instead of overwriting.
  const paths = new Set();
  for (const folder of folders) {
    const key = folder.path.toLowerCase();
    if (paths.has(key) || used.has(key)) throw new Error(`Conflicting ZIP path: ${folder.path}. Resolve the source names before exporting.`);
    paths.add(key);
  }
  return { schema: 1, source: SOURCE, target: TARGET, listId: list.Id, listTitle: list.Title, scannedAt: new Date().toISOString(), files, folders, warnings, snapshot: rows.map(row => ({ id: row.Id, modified: row.Modified, row })).sort((a, b) => a.id - b.id) };
}

export async function acceptFile(file, bytes, { name } = {}) {
  if (name && name !== file.sourceName) throw new Error(`Choose the original file named ${file.sourceName}.`);
  if (file.expectedSize !== null && bytes.byteLength !== file.expectedSize) throw new Error(`File size does not match the source record for ${file.sourceName}: expected ${file.expectedSize}, received ${bytes.byteLength} bytes.`);
  const hash = await crypto.subtle.digest('SHA-256', bytes);
  return { bytes, sha256: [...new Uint8Array(hash)].map(value => value.toString(16).padStart(2, '0')).join('') };
}

export async function makeZip(plan, downloads, onProgress = () => {}) {
  if (plan.files.some(file => !downloads.has(file.key))) throw new Error('Every inventoried attachment must be downloaded before a complete ZIP can be created.');
  const zip = new JSZip();
  const manifest = { ...plan, snapshot: undefined, exportedAt: new Date().toISOString(), files: plan.files.map(file => ({ ...file, bytes: downloads.get(file.key).bytes.byteLength, sha256: downloads.get(file.key).sha256, retrievedBy: downloads.get(file.key).retrievedBy || 'SharePoint GET' })) };
  for (const folder of plan.folders) zip.folder(folder.path);
  for (const file of plan.files) zip.file(file.zipPath, downloads.get(file.key).bytes, { binary: true });
  zip.file('Reference-Document-Upload-Guide.html', uploadGuide(manifest));
  zip.file('reference-manifest.json', JSON.stringify(manifest, null, 2));
  // STORE avoids wasting time recompressing Office/PDF files. No 150 MB JSON cap.
  // ZIP32 sizes are checked explicitly; browser memory may be a lower bound.
  const total = plan.files.reduce((sum, file) => sum + downloads.get(file.key).bytes.byteLength, 0);
  if (total >= 0xffffffff || plan.files.length + plan.folders.length >= 65000) throw new Error('This export exceeds the ZIP32 limit. Split the reference collection before exporting.');
  return { blob: await zip.generateAsync({ type: 'blob', compression: 'STORE', streamFiles: true }, meta => onProgress(meta.percent)), manifest };
}
