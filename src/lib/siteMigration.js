import { CONTAINERS, SharePointStore } from './spStore';
import { retryMigrationRead, reconcileMigrationWrite, migrationRequest } from './migrationRecovery';
import { getFormDigest } from './spContext';
import { makeMigrationBatch, checkMigrationBatch } from './migrationBatch';

export const MIGRATION_SOURCE = 'https://flankspeed.sharepoint-mil.us/sites/ISEAMETENG';
export const MIGRATION_TARGET = 'https://flankspeed.sharepoint-mil.us/sites/metsoft';
export const MIGRATION_PAGE = `${MIGRATION_TARGET}/SitePages/Modernization-Tracker.aspx`;
const normalizeUrl = value => String(value || '').replace(/\/+$/, '').toLowerCase();
export const isMigrationTarget = value => normalizeUrl(value) === normalizeUrl(MIGRATION_TARGET);
const quote = value => encodeURIComponent(String(value).replace(/'/g, "''"));
const items = body => body.value || body.d?.results || [];
const unwrap = body => body.d || body;
export const fieldsFor = container => ['Title', ...container.fields.map(field => field.name)];
const fieldTypes = new Map(CONTAINERS.flatMap(container => container.fields.map(field => [field.name, field.type])));
const clean = (value, field) => fieldTypes.get(field) === 'Boolean' ? Number(Boolean(value)) : value == null ? '' : fieldTypes.get(field) === 'DateTime' && value ? new Date(value).toISOString() : value;
export const same = (a, b, fields) => fields.every(field => clean(a[field], field) === clean(b[field], field));
// Archived users are historical records, not competing active accounts. Keep
// their RecordIds and match them only by RecordId, never by login/email.
const identity = (key, row) => key === 'users' ? row.Archived ? '' : String(row.LoginKey || row.Email || '').trim().toLowerCase().split('|').pop() : key === 'projects' ? row.ProjectKey : key === 'acronyms' ? String(row.Acronym || '').toLowerCase() : '';
export const writable = (row, fields) => Object.fromEntries(fields.map(field => [field, fieldTypes.get(field) === 'Boolean' ? Boolean(row[field]) : row[field] ?? null]));
export const sha256 = async bytes => [...new Uint8Array(await crypto.subtle.digest('SHA-256', bytes))].map(n => n.toString(16).padStart(2, '0')).join('');

// This adapter uses the current SharePoint session and existing request plumbing.
// It neither changes site permissions nor bypasses host confirmation policies.
export function migrationSite(webUrl, prefix = 'Modernization', fetchImpl = fetch, { wait, timeoutMs, onActivity = () => {} } = {}) {
  webUrl = webUrl.replace(/\/+$/, '');
  prefix = String(prefix || 'Modernization').replace(/[^A-Za-z0-9]/g, '') || 'Modernization';
  const transport = fetchImpl;
  fetchImpl = (url, options = {}) => {
    let attempt = 0;
    const request = async () => {
      onActivity(`${options.method || 'GET'} ${new URL(url).pathname} · attempt ${++attempt}`);
      const response = await migrationRequest(transport, url, options, timeoutMs);
      if ([408, 429, 502, 503, 504].includes(response.status)) {
        const retryAfter = response.headers?.get('Retry-After');
        const retryAfterMs = retryAfter ? /^\d+$/.test(retryAfter) ? Number(retryAfter) * 1000 : Math.max(0, Date.parse(retryAfter) - Date.now()) : 0;
        throw Object.assign(new Error(`SharePoint returned HTTP ${response.status}.`), { status: response.status, retryAfterMs });
      }
      return response;
    };
    return (options.method || 'GET').toUpperCase() === 'GET' ? retryMigrationRead(request, wait) : request();
  };
  const store = new SharePointStore({ webUrl, prefix, fetchImpl });
  const root = key => `/_api/web/lists/getbytitle('${quote(`${prefix}${CONTAINERS.find(c => c.key === key).suffix}`)}')`;
  const documentUrls = new Map();
  const documentKey = (key, id, name) => JSON.stringify([key, Number(id), name.toLowerCase()]);
  function rememberDocuments(key, id, files) {
    for (const file of files) {
      if (!file.ServerRelativeUrl) continue;
      // FileName is decoded metadata. Encode the last segment explicitly so
      // literal # and % characters cannot become fragments or escape sequences.
      const address = file.ServerRelativeUrl.slice(0, file.ServerRelativeUrl.lastIndexOf('/') + 1) + encodeURIComponent(file.FileName);
      const url = new URL(address, webUrl), site = new URL(webUrl);
      if (url.origin !== site.origin || !url.pathname.toLowerCase().startsWith(`${site.pathname.toLowerCase()}/`) || url.username || url.password || url.hash) throw new Error('SharePoint returned an unexpected document address.');
      documentUrls.set(documentKey(key, id, file.FileName), url.href);
    }
  }
  async function attachmentNames(key, id) {
    const files = items(await store.get(`${root(key)}/items(${Number(id)})/AttachmentFiles?$select=FileName,ServerRelativeUrl`));
    rememberDocuments(key, id, files);
    return files.map(file => file.FileName);
  }
  async function readRows(key, fields, withAttachments = false) {
    let path = `${root(key)}/items?$select=${['Id', 'Modified', ...fields, ...(withAttachments ? ['AttachmentFiles/FileName', 'AttachmentFiles/ServerRelativeUrl'] : [])].join(',')}&$top=500${withAttachments ? '&$expand=AttachmentFiles' : ''}`;
    const rows = [], visited = new Set();
    while (path) {
      if (visited.has(path)) throw new Error('SharePoint returned a repeated pagination link.');
      visited.add(path);
      const response = await fetchImpl(`${webUrl}${path}`, { credentials: 'include', cache: 'no-store', headers: { Accept: 'application/json;odata=minimalmetadata' } });
      if (!response.ok) throw new Error(`Cannot read ${key} (${response.status}).`);
      const body = await response.json(), page = items(body); rows.push(...page);
      if (withAttachments) for (const row of page) rememberDocuments(key, row.Id, row.AttachmentFiles?.results || row.AttachmentFiles || []);
      const next = body['@odata.nextLink'] || body['odata.nextLink'] || body.d?.__next;
      if (!next) break;
      const url = new URL(next, `${webUrl}/`);
      if (!normalizeUrl(url.href).startsWith(`${normalizeUrl(webUrl)}/_api/`)) throw new Error('Unexpected pagination destination.');
      path = url.href.slice(webUrl.length);
    }
    return rows;
  }
  return {
    webUrl,
    async canMigrate() {
      const body = unwrap(await store.get('/_api/web/EffectiveBasePermissions'));
      const permissions = body.EffectiveBasePermissions || body;
      // Require Manage Lists AND Manage Permissions (normally site Owners).
      const required = 2048n | 33554432n;
      return (BigInt(permissions.Low || 0) & required) === required;
    },
    async schema(container) {
      const body = await store.get(`${root(container.key)}/fields?$select=InternalName&$top=500`);
      return new Set(items(body).map(field => field.InternalName));
    },
    readRows,
    async writeBatch(key, operations) {
      if (!operations.length) return;
      const batch = makeMigrationBatch(webUrl, root(key), operations);
      const digest = await getFormDigest(webUrl, fetchImpl);
      const response = await fetchImpl(`${webUrl}/_api/$batch`, { method: 'POST', credentials: 'include', headers: { 'X-RequestDigest': digest, 'Content-Type': batch.contentType, Accept: 'multipart/mixed' }, body: batch.body });
      if (!response.ok) throw new Error(`SharePoint batch failed (${response.status}). Preview again before resuming.`);
      checkMigrationBatch(await response.text(), operations.length);
    },
    async readItem(key, id, fields) {
      const response = await fetchImpl(`${webUrl}${root(key)}/items(${Number(id)})?$select=${['Id', 'Modified', ...fields].join(',')}`, { credentials: 'include', cache: 'no-store', headers: { Accept: 'application/json;odata=minimalmetadata' } });
      if (!response.ok) throw new Error(`Cannot read ${key}/${id} (${response.status}).`);
      const row = unwrap(await response.json());
      return { ...row, etag: response.headers.get('ETag') || row['odata.etag'] || row['@odata.etag'] || row.__metadata?.etag };
    },
    async create(key, fields) { const id = await store.create(key, fields); if (!id) throw new Error(`No ID returned creating ${key}. Refresh the preview before retrying.`); return id; },
    async update(key, id, fields, etag) {
      if (!etag) throw new Error('SharePoint did not return a concurrency token. No existing record was overwritten.');
      await store.post(`${root(key)}/items(${Number(id)})`, { body: fields, headers: { 'IF-MATCH': etag, 'X-HTTP-Method': 'MERGE' } });
    },
    attachments: attachmentNames,
    async bytes(key, id, name) {
      // Use the same direct file download route as normal tracker downloads.
      // The embedded host's REST proxy may not complete binary $value requests.
      const keyForFile = documentKey(key, id, name);
      if (!documentUrls.has(keyForFile)) await attachmentNames(key, id);
      const url = documentUrls.get(keyForFile);
      if (!url) throw new Error(`SharePoint did not return a download address for ${name}. No document was skipped.`);
      try {
        const response = await fetchImpl(url, { credentials: 'include', cache: 'no-store', headers: { Accept: '*/*' } });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        if (/text\/html/i.test(response.headers.get('Content-Type') || '') && !/\.html?$/i.test(name)) throw new Error('SharePoint returned an HTML page instead of the document');
        return await response.arrayBuffer();
      } catch (error) {
        throw new Error(`Document download failed: ${name} (${key}, item ${id}). ${error.message}. Try downloading this document from the tracker document menu to check site access. No document was skipped.`);
      }
    },
    async attach(key, id, name, bytes) { await store.post(`${root(key)}/items(${Number(id)})/AttachmentFiles/add(FileName='${quote(name)}')`, { raw: true, headers: { 'Content-Type': 'application/octet-stream' }, body: bytes }); },
  };
}

function assertSites(source, target) {
  if (normalizeUrl(source.webUrl) !== normalizeUrl(MIGRATION_SOURCE) || !isMigrationTarget(target.webUrl)) throw new Error('This migration only supports ISEA METENG → metsoft.');
}

function indexRows(key, rows, site) {
  const index = new Map(), identities = new Map();
  for (const row of rows) {
    if (!row.RecordId || index.has(row.RecordId)) throw new Error(`${key}: missing or duplicate RecordId; resolve before migration.`);
    index.set(row.RecordId, row);
    const logical = identity(key, row);
    if (logical && identities.has(logical)) {
      const previous = identities.get(logical);
      if (key === 'users') throw new Error(`${site}: multiple active user records for ${logical}: ${previous.Title} (item ${previous.Id}, role ${previous.AppRole || 'unset'}) and ${row.Title} (item ${row.Id}, role ${row.AppRole || 'unset'}). Review these accounts in that site's tracker directory and remove the obsolete entry, then preview again. Archived users are preserved automatically; no role has been selected or changed.`);
      throw new Error(`${key}: duplicate logical identity; resolve before migration.`);
    }
    if (key === 'projects' && !logical) throw new Error('Project is missing ProjectKey.');
    if (logical) identities.set(logical, row);
  }
  return index;
}

export async function previewMigration(source, target, onProgress = () => {}) {
  assertSites(source, target);
  if (!(await target.canMigrate())) throw new Error('Migration requires SharePoint Manage Lists and Manage Permissions on metsoft, normally site Owner access. An app Manager role alone is insufficient.');
  const plan = { source: source.webUrl, target: target.webUrl, createdAt: new Date().toISOString(), lists: [] };
  for (const container of CONTAINERS) {
    onProgress(`Reading ${container.suffix}…`);
    const [sourceSchema, targetSchema] = await Promise.all([source.schema(container), target.schema(container)]);
    for (const required of ['Title', 'RecordId']) if (!sourceSchema.has(required)) throw new Error(`${container.suffix}: source is missing ${required}.`);
    const fields = fieldsFor(container).filter(field => sourceSchema.has(field));
    const missing = fields.filter(field => !targetSchema.has(field));
    if (missing.length) throw new Error(`${container.suffix}: destination needs fields ${missing.join(', ')}. Open the current tracker build on metsoft as a site owner to complete setup.`);
    const [rows, existing] = await Promise.all([source.readRows(container.key, fields), target.readRows(container.key, fields)]);
    indexRows(container.key, rows, 'ISEA METENG source'); const byId = indexRows(container.key, existing, 'metsoft destination');
    const matched = new Set(), entries = [];
    for (const row of rows) {
      const logical = identity(container.key, row);
      const candidates = existing.filter(item => item.RecordId === row.RecordId || (logical && identity(container.key, item) === logical));
      if (candidates.length > 1 || (candidates[0] && matched.has(candidates[0].Id))) throw new Error(`${container.suffix}: ambiguous matching records for ${row.Title || row.RecordId}.`);
      const destination = byId.get(row.RecordId) || candidates[0];
      if (destination) matched.add(destination.Id);
      const names = ['tasks', 'references'].includes(container.key) ? await source.attachments(container.key, row.Id) : [];
      entries.push({ source: row, destination, names, conflict: !!destination && !same(row, destination, fields) });
    }
    plan.lists.push({ key: container.key, title: container.suffix, fields, entries, extra: existing.filter(row => !matched.has(row.Id)).length });
  }
  const projects = new Set(plan.lists.find(list => list.key === 'projects').entries.map(entry => entry.source.ProjectKey));
  for (const list of plan.lists.filter(list => ['tasks', 'updates', 'risks'].includes(list.key))) {
    for (const { source: row } of list.entries) if (!projects.has(row.ProjectKey)) throw new Error(`${list.title}: ${row.RecordId} references a project missing from the source.`);
  }
  const refs = new Map(plan.lists.find(list => list.key === 'references').entries.map(entry => [entry.source.RecordId, entry.source]));
  for (const row of refs.values()) {
    const visited = new Set([row.RecordId]); let parent = row.ReferenceParentId;
    while (parent) {
      if (visited.has(parent) || !refs.has(parent) || refs.get(parent).EntryKind !== 'folder') throw new Error('Reference document folder relationships are invalid.');
      visited.add(parent); parent = refs.get(parent).ReferenceParentId;
    }
  }
  return plan;
}

export async function runMigration(source, target, plan, { replaceConflicts = false, onProgress = () => {}, wait } = {}) {
  assertSites(source, target);
  if (plan.source !== source.webUrl || plan.target !== target.webUrl) throw new Error('The preview does not match these sites.');
  if (!(await target.canMigrate())) throw new Error('SharePoint Manage Lists and Manage Permissions are required.');
  if (!replaceConflicts && plan.lists.some(list => list.entries.some(entry => entry.conflict))) throw new Error('Review and explicitly approve conflicting destination records first.');
  // Refresh both sides before any write. The reviewed snapshot must still match.
  for (const list of plan.lists) {
    onProgress(`Checking ${list.title} source and destination before copying…`);
    const [sourceRows, targetRows] = await Promise.all([source.readRows(list.key, list.fields), target.readRows(list.key, list.fields)]);
    const original = list.entries.map(entry => entry.source);
    if (sourceRows.length !== original.length || original.some(row => !sourceRows.some(now => now.RecordId === row.RecordId && now.Modified === row.Modified && same(now, row, list.fields)))) throw new Error('Source changed after preview. Pause edits and refresh the preview.');
    for (const entry of list.entries) {
      const logical = identity(list.key, entry.source);
      const matches = targetRows.filter(row => row.RecordId === entry.source.RecordId || (logical && identity(list.key, row) === logical));
      if (entry.destination ? matches.length !== 1 || matches[0].Id !== entry.destination.Id || matches[0].Modified !== entry.destination.Modified || !same(matches[0], entry.destination, list.fields) : matches.length > 0) throw new Error('Destination changed after preview. Refresh the preview before copying.');
    }
  }
  const report = { source: source.webUrl, target: target.webUrl, startedAt: new Date().toISOString(), verified: false, lists: [], files: [] };
  for (const list of plan.lists) {
    const result = { list: list.title, records: 0, attachments: 0, destinationExtrasRetained: list.extra };
    for (const entry of list.entries) {
      onProgress(`Copying ${list.title}: ${result.records + 1}/${list.entries.length}…`);
      let id = entry.destination?.Id;
      if (id) {
        const current = await target.readItem(list.key, id, list.fields);
        if (current.Modified !== entry.destination.Modified || !same(current, entry.destination, list.fields)) throw new Error('Destination record changed during migration. Refresh preview to resume.');
        if (entry.conflict) await reconcileMigrationWrite(
          () => target.update(list.key, id, writable(entry.source, list.fields), current.etag),
          async () => same(await target.readItem(list.key, id, list.fields), entry.source, list.fields) ? { value: undefined } : null,
          { onProgress, wait, label: `${list.title}/${entry.source.RecordId}` },
        );
      } else id = await reconcileMigrationWrite(
        () => target.create(list.key, writable(entry.source, list.fields)),
        async () => {
          const matches = (await target.readRows(list.key, list.fields)).filter(row => row.RecordId === entry.source.RecordId);
          if (matches.length > 1) throw new Error(`Multiple destination records found for ${entry.source.RecordId}. Review before resuming.`);
          return matches.length === 1 && same(matches[0], entry.source, list.fields) ? { value: matches[0].Id } : null;
        },
        { onProgress, wait, label: `${list.title}/${entry.source.RecordId}` },
      );
      entry.targetId = id;
      const existingNames = ['tasks', 'references'].includes(list.key) ? await target.attachments(list.key, id) : [];
      for (const name of entry.names) {
        onProgress(`Copying ${list.title} document: ${name}…`);
        const bytes = await source.bytes(list.key, entry.source.Id, name), hash = await sha256(bytes);
        const matchingName = existingNames.find(existing => existing.toLowerCase() === name.toLowerCase());
        if (matchingName) {
          if (await sha256(await target.bytes(list.key, id, matchingName)) !== hash) throw new Error(`Different destination file named ${name}. No file was overwritten. Resolve it in the destination, then preview again.`);
        } else await reconcileMigrationWrite(
          () => target.attach(list.key, id, name, bytes),
          async () => {
            const found = (await target.attachments(list.key, id)).find(value => value.toLowerCase() === name.toLowerCase());
            if (!found) return null;
            if (await sha256(await target.bytes(list.key, id, found)) !== hash) throw new Error(`Different destination file named ${name}. No file was overwritten.`);
            return { value: undefined };
          },
          { onProgress, wait, label: `${list.title} document ${name}` },
        );
        if (await sha256(await target.bytes(list.key, id, matchingName || name)) !== hash) throw new Error(`Document verification failed: ${name}. Preview again to resume.`);
        report.files.push({ list: list.key, recordId: entry.source.RecordId, sourceId: entry.source.Id, targetId: id, name, targetName: matchingName || name, bytes: bytes.byteLength, sha256: hash });
        result.attachments++;
      }
      const saved = await target.readItem(list.key, id, list.fields);
      if (!same(saved, entry.source, list.fields)) throw new Error(`Record verification failed: ${list.title}/${entry.source.RecordId}.`);
      result.records++;
    }
    report.lists.push(result);
  }
  onProgress('Verifying records and source consistency…');
  for (const list of plan.lists) {
    onProgress(`Verifying ${list.title} source snapshot…`);
    const rows = await source.readRows(list.key, list.fields);
    if (rows.length !== list.entries.length) throw new Error('Source changed during migration. Refresh preview and repeat before switching users.');
    for (const entry of list.entries) {
      onProgress(`Verifying ${list.title}: ${entry.source.Title || entry.source.RecordId}…`);
      const now = rows.find(row => row.RecordId === entry.source.RecordId);
      if (!now || now.Modified !== entry.source.Modified || !same(now, entry.source, list.fields)) throw new Error('Source changed during migration. Refresh preview and repeat before switching users.');
      if (!same(await target.readItem(list.key, entry.targetId, list.fields), entry.source, list.fields)) throw new Error('Destination changed during verification. Refresh preview and review.');
      if (['tasks', 'references'].includes(list.key)) {
        const names = await source.attachments(list.key, entry.source.Id);
        if (JSON.stringify([...names].sort()) !== JSON.stringify([...entry.names].sort())) throw new Error('Source attachments changed during migration. Refresh preview and repeat.');
      }
    }
  }
  for (const file of report.files) {
    onProgress(`Verifying document: ${file.name}…`);
    if (await sha256(await source.bytes(file.list, file.sourceId, file.name)) !== file.sha256 || await sha256(await target.bytes(file.list, file.targetId, file.targetName)) !== file.sha256) throw new Error(`Document changed during verification: ${file.name}. Refresh preview and repeat.`);
  }
  report.verified = true; report.completedAt = new Date().toISOString();
  return report;
}
