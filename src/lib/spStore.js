import { TrackerLists, displayListTitle, legacyListTitle } from './trackerLists';
import { normalizeDirectory } from './identity';
import { trackerPageUrl, normalizeInvitationUrl, parsePeopleResults } from './peoplePicker';
import { validateAttachment, validateAttachmentName } from './taskAttachments';
import { normalizePhaseKey, normalizeOrganization, normalizeTaskStatus, normalizeTaskOrganizations, ROLES } from '../data/workflow';
import { getCurrentUser, SharePointError, spGet, spPost } from './spContext';
import { defaultAcronyms } from '../data/defaultAcronyms';
import { TrackerPermissions } from './trackerPermissions';
import { TrackerAccessJob } from './trackerAccessJob';
import { isOwnedByUser } from './identity';

const FIELD = { TEXT: 'Text', NOTE: 'Note', NUMBER: 'Number', DATE: 'DateTime', BOOLEAN: 'Boolean' };
const ADD_FIELD = { INTERNAL_NAME_HINT: 8, TO_DEFAULT_VIEW: 16 };

export const CONTAINERS = [
  {
    key: 'projects', suffix: 'Projects', description: 'Modernization project portfolio records.', fields: [
      ['RecordId', 'Record ID', FIELD.TEXT, true], ['ProjectKey', 'Project Key', FIELD.TEXT, true],
      ['MeasurementArea', 'Measurement Area', FIELD.TEXT], ['Description', 'Description', FIELD.NOTE],
      ['OwnerName', 'Owner', FIELD.TEXT], ['OwnerEmail', 'Owner Email', FIELD.TEXT], ['OwnerKey', 'Owner Identity Key', FIELD.TEXT, true],
      ['Organization', 'Organization', FIELD.TEXT], ['OwnerCanEdit', 'Project Owner Can Edit', FIELD.BOOLEAN],
      ['Priority', 'Priority', FIELD.TEXT], ['Health', 'Health', FIELD.TEXT], ['ProjectStatus', 'Status', FIELD.TEXT],
      ['CurrentStageKey', 'Current Stage', FIELD.TEXT], ['PercentComplete', 'Percent Complete', FIELD.NUMBER], ['ProgressMode', 'Progress Mode', FIELD.TEXT],
      ['TargetFinish', 'Target Finish', FIELD.DATE], ['NextMilestone', 'Next Milestone', FIELD.TEXT],
      ['NextMilestoneDate', 'Next Milestone Date', FIELD.DATE], ['SourceNotes', 'Source Notes', FIELD.NOTE],
      ['ImportedBaseline', 'Imported Baseline', FIELD.BOOLEAN], ['TagsJson', 'Tags', FIELD.NOTE],
    ],
  },
  {
    key: 'tasks', suffix: 'Tasks', description: 'Pipeline tasks for every modernization project.', fields: [
      ['RecordId', 'Record ID', FIELD.TEXT, true], ['ProjectKey', 'Project Key', FIELD.TEXT, true],
      ['ArchivedDocuments', 'Archived Documents', FIELD.NOTE], ['TaskTitle', 'Task', FIELD.TEXT], ['PhaseKey', 'Phase', FIELD.TEXT],
      ['Organization', 'Organization', FIELD.TEXT], ['EstimatedHours', 'Est. Hours', FIELD.NUMBER], ['SortOrder', 'Sort Order', FIELD.NUMBER], ['TaskStatus', 'Status', FIELD.TEXT],
      ['AssignedDate', 'Assigned / Creation Date', FIELD.DATE], ['StartDate', 'Start Date', FIELD.DATE], ['FinishDate', 'Finish Date', FIELD.DATE], ['DueDate', 'Due Date', FIELD.DATE],
      ['OwnerName', 'Owner', FIELD.TEXT], ['OwnerEmail', 'Owner Email', FIELD.TEXT], ['OwnerKey', 'Owner Identity Key', FIELD.TEXT, true], ['Notes', 'Notes', FIELD.NOTE],
      ['BlockedReason', 'Blocked Reason', FIELD.NOTE], ['SourceStartLabel', 'Source Start Label', FIELD.TEXT],
      ['DataIssue', 'Data Issue', FIELD.NOTE],
      ['DeferredDate', 'Deferred Date', FIELD.DATE], ['DeferredJustification', 'Deferral Justification', FIELD.NOTE], ['NotRequiredJustification', 'Not Required Justification', FIELD.NOTE],
    ],
  },
  {
    key: 'updates', suffix: 'Updates', description: 'Project status updates and decisions.', fields: [
      ['RecordId', 'Record ID', FIELD.TEXT, true], ['ProjectKey', 'Project Key', FIELD.TEXT, true],
      ['UpdateType', 'Update Type', FIELD.TEXT], ['Summary', 'Summary', FIELD.NOTE], ['EntryDate', 'Entry Date', FIELD.DATE],
      ['AuthorName', 'Author', FIELD.TEXT], ['AuthorEmail', 'Author Email', FIELD.TEXT], ['AuthorKey', 'Author Identity Key', FIELD.TEXT, true],
    ],
  },
  {
    key: 'risks', suffix: 'Risks', description: 'Project risks, issues, and mitigation actions.', fields: [
      ['RecordId', 'Record ID', FIELD.TEXT, true], ['ProjectKey', 'Project Key', FIELD.TEXT, true],
      ['RiskTitle', 'Risk', FIELD.TEXT], ['Severity', 'Severity', FIELD.TEXT], ['Probability', 'Probability', FIELD.TEXT],
      ['Mitigation', 'Mitigation', FIELD.NOTE], ['OwnerName', 'Owner', FIELD.TEXT], ['OwnerKey', 'Owner Identity Key', FIELD.TEXT, true], ['RiskStatus', 'Status', FIELD.TEXT],
      ['DueDate', 'Due Date', FIELD.DATE],
    ],
  },
  {
    key: 'acronyms', suffix: 'Acronyms', description: 'Shared modernization acronym glossary.', fields: [
      ['RecordId', 'Record ID', FIELD.TEXT, true], ['Acronym', 'Acronym', FIELD.TEXT, true],
      ['FullTerm', 'Full Term', FIELD.TEXT], ['Definition', 'Definition', FIELD.NOTE], ['SeedVersion', 'Seed Version', FIELD.TEXT],
    ],
  },
  { key: 'users', suffix: 'Users', description: 'Tracker user directory and application roles.', fields: [
    ['RecordId', 'Record ID', FIELD.TEXT, true], ['LoginKey', 'Login Key', FIELD.TEXT, true], ['Email', 'Email', FIELD.TEXT], ['AppRole', 'Application Role', FIELD.TEXT],
  ] },
  { key: 'references', suffix: 'ReferenceDocuments', description: 'Shared reference templates and folder organization.', fields: [
    ['RecordId', 'Record ID', FIELD.TEXT, true], ['ReferenceParentId', 'Parent ID', FIELD.TEXT, true],
    ['EntryKind', 'Entry Kind', FIELD.TEXT], ['FileName', 'File Name', FIELD.TEXT], ['FileSize', 'File Size', FIELD.NUMBER],
  ] },
].map((container) => ({
  ...container,
  fields: [...container.fields, ['Archived', 'Archived', FIELD.BOOLEAN]].map(([name, title, type, indexed = false]) => ({ name, title, type, indexed, inView: type !== FIELD.NOTE })),
}));

const escapeXml = (value) => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const escapeOData = (value) => encodeURIComponent(String(value).replace(/'/g, "''"));
const titleFor = legacyListTitle;
const apiFor = (prefix, key) => `/_api/web/lists/getbytitle('${escapeOData(titleFor(prefix, key))}')`;

function schemaXml(field) {
  const attrs = { Type: field.type, DisplayName: field.title, Name: field.name, StaticName: field.name };
  if (field.type === FIELD.NOTE) Object.assign(attrs, { NumLines: 8, RichText: 'FALSE', AppendOnly: 'FALSE' });
  if (field.indexed && field.type !== FIELD.NOTE) attrs.Indexed = 'TRUE';
  if (field.type === FIELD.DATE) Object.assign(attrs, { Format: 'DateOnly', FriendlyDisplayFormat: 'Disabled' });
  return `<Field ${Object.entries(attrs).map(([key, value]) => `${key}="${escapeXml(value)}"`).join(' ')} />`;
}

const dateOnly = (value) => value ? String(value).slice(0, 10) : '';
const sharePointDate = (value) => value ? `${dateOnly(value)}T12:00:00Z` : null;
const DATE_FIELDS = new Set(['TargetFinish', 'NextMilestoneDate', 'AssignedDate', 'StartDate', 'FinishDate', 'DueDate', 'DeferredDate', 'EntryDate']);
const sharePointFormDate = (value) => {
  if (!value) return '';
  const [year, month, day] = dateOnly(value).split('-').map(Number);
  return year && month && day ? `${month}/${day}/${year}` : '';
};
const phaseKey = normalizePhaseKey;
const safeJson = (value, fallback) => {
  try { return value ? JSON.parse(value) : fallback; } catch { return fallback; }
};

const projectFields = (row) => ({
  Title: row.title, RecordId: row.id, ProjectKey: row.projectKey, MeasurementArea: row.measurementArea,
  Description: row.description, OwnerName: row.ownerName, OwnerEmail: row.ownerEmail, OwnerKey: row.ownerKey || '',
  Priority: row.priority, Health: row.health, ProjectStatus: row.status,
  Organization: normalizeOrganization(row.organization), OwnerCanEdit: !!row.ownerCanEdit,
  CurrentStageKey: row.currentStageKey, PercentComplete: row.percentComplete, ProgressMode: row.progressMode || 'phases', TargetFinish: sharePointDate(row.targetFinish),
  NextMilestone: row.nextMilestone, NextMilestoneDate: sharePointDate(row.nextMilestoneDate), SourceNotes: row.sourceNotes,
  ImportedBaseline: !!row.importedBaseline, TagsJson: JSON.stringify(row.tags || []),
});

const taskFields = (row) => ({
  Title: row.title, RecordId: row.id, ProjectKey: row.projectKey,
  Organization: normalizeOrganization(row.organization), TaskTitle: row.title, PhaseKey: row.phaseKey, SortOrder: row.order, TaskStatus: row.status,
  EstimatedHours: row.estimatedHours === '' || row.estimatedHours == null ? null : Number(row.estimatedHours), AssignedDate: sharePointDate(row.assignedDate), StartDate: sharePointDate(row.startDate), FinishDate: sharePointDate(row.finishDate), DueDate: sharePointDate(row.dueDate),
  OwnerName: row.ownerName, OwnerEmail: row.ownerEmail, OwnerKey: row.ownerKey || '', Notes: row.notes, BlockedReason: row.blockedReason,
  SourceStartLabel: row.sourceStartLabel, DataIssue: row.dataIssue,
  DeferredDate: sharePointDate(row.deferredDate), DeferredJustification: row.deferredJustification || '', NotRequiredJustification: row.notRequiredJustification || '',
});

const updateFields = (row) => ({
  Title: `${row.projectKey} update`, RecordId: row.id, ProjectKey: row.projectKey, UpdateType: row.type,
  Summary: row.summary, EntryDate: sharePointDate(row.entryDate), AuthorName: row.authorName, AuthorEmail: row.authorEmail, AuthorKey: row.authorKey || '',
});

const riskFields = (row) => ({
  Title: row.title, RecordId: row.id, ProjectKey: row.projectKey, RiskTitle: row.title, Severity: row.severity,
  Probability: row.probability, Mitigation: row.mitigation, OwnerName: row.ownerName, OwnerKey: row.ownerKey || '', RiskStatus: row.status,
  DueDate: sharePointDate(row.dueDate),
});

const acronymFields = (row) => ({
  Title: row.acronym, RecordId: row.id, Acronym: row.acronym, FullTerm: row.term, Definition: row.definition, SeedVersion: row.seedVersion || '',
});

function fromProject(item) {
  return {
    spId: item.Id, id: item.RecordId, projectKey: item.ProjectKey, title: item.Title,
    measurementArea: item.MeasurementArea || item.Title, description: item.Description || '', ownerName: item.OwnerName || 'Unassigned',
    ownerEmail: item.OwnerEmail || '', ownerKey: item.OwnerKey || '',
    organization: normalizeOrganization(item.Organization), ownerCanEdit: !!item.OwnerCanEdit,
    priority: item.Priority || 'Medium', health: item.Health || 'Needs Review', status: item.ProjectStatus || 'Planned',
    currentStageKey: phaseKey(item.CurrentStageKey), percentComplete: Number(item.PercentComplete || 0), progressMode: item.ProgressMode === 'tasks' ? 'tasks' : 'phases',
    targetFinish: dateOnly(item.TargetFinish), nextMilestone: item.NextMilestone || '', nextMilestoneDate: dateOnly(item.NextMilestoneDate),
    sourceNotes: item.SourceNotes || '', importedBaseline: !!item.ImportedBaseline, tags: safeJson(item.TagsJson, []),
  };
}

function fromTask(item) {
  const status = normalizeTaskStatus(item.TaskStatus);
  return {
    spId: item.Id, id: item.RecordId, projectKey: item.ProjectKey, title: item.TaskTitle,
    organization: item.Organization || '', phaseKey: phaseKey(item.PhaseKey), order: Number(item.SortOrder || 0), status,
    estimatedHours: item.EstimatedHours == null ? null : Number(item.EstimatedHours), assignedDate: dateOnly(item.AssignedDate), startDate: dateOnly(item.StartDate), finishDate: dateOnly(item.FinishDate), dueDate: dateOnly(item.DueDate), deferredDate: dateOnly(item.DeferredDate), deferredJustification: item.DeferredJustification || '', notRequiredJustification: item.NotRequiredJustification || '',
    ownerName: item.OwnerName || 'Unassigned', ownerEmail: item.OwnerEmail || '', ownerKey: item.OwnerKey || '', notes: item.Notes || '',
    blockedReason: item.BlockedReason || '', sourceStartLabel: item.SourceStartLabel || '', dataIssue: item.DataIssue || '',
  };
}

const fromUser = (item) => ({ spId: item.Id, id: item.RecordId, title: item.Title, loginName: item.LoginKey, email: item.Email || '', role: item.AppRole || 'User' });

const fromUpdate = (item) => ({ spId: item.Id, id: item.RecordId, projectKey: item.ProjectKey, type: item.UpdateType, summary: item.Summary, entryDate: dateOnly(item.EntryDate), authorName: item.AuthorName, authorEmail: item.AuthorEmail, authorKey: item.AuthorKey || '' });
const fromRisk = (item) => ({ spId: item.Id, id: item.RecordId, projectKey: item.ProjectKey, title: item.RiskTitle || item.Title, severity: item.Severity, probability: item.Probability, mitigation: item.Mitigation || '', ownerName: item.OwnerName || '', ownerKey: item.OwnerKey || '', status: item.RiskStatus || 'Open', dueDate: dateOnly(item.DueDate) });
const fromAcronym = (item) => ({ spId: item.Id, id: item.RecordId, acronym: item.Acronym || item.Title, term: item.FullTerm || '', definition: item.Definition || '', seedVersion: item.SeedVersion || '' });

export class SharePointStore {
  constructor({ webUrl, prefix = 'Modernization', fetchImpl = fetch, scopedAccess = false, invitationAssetUrls = [], invitationAssetFolders = [] }) {
    this.webUrl = String(webUrl || '').replace(/\/+$/, '');
    this.prefix = prefix;
    this.fetchImpl = !scopedAccess ? fetchImpl : async (url, options = {}) => {
      const controller = new AbortController();
      const abort = () => controller.abort();
      if (options.signal?.aborted) abort();
      options.signal?.addEventListener('abort', abort, { once: true });
      const timer = setTimeout(abort, 60000);
      try {
        const response = await fetchImpl(url, { ...options, signal: controller.signal });
        const bytes = await response.arrayBuffer();
        return new Response([204, 205, 304].includes(response.status) ? null : bytes, { status: response.status, statusText: response.statusText, headers: response.headers });
      }
      catch (error) { if (controller.signal.aborted) throw new Error('SharePoint request timed out after 60 seconds. A write may have completed; reload or rerun permission setup to verify it.'); throw error; }
      finally { clearTimeout(timer); options.signal?.removeEventListener('abort', abort); }
    };
    this.lists = new TrackerLists(prefix, path => spGet(this.webUrl, path, this.fetchImpl), (path, options) => spPost(this.webUrl, path, options, this.fetchImpl));
    this.userPromise = null;
    this.scopedAccess = scopedAccess;
    this.invitationAssetUrls = invitationAssetUrls;
    this.invitationAssetFolders = invitationAssetFolders;
    this.permissions = scopedAccess ? new TrackerPermissions(this) : null;
    this.trackerAccessJob = scopedAccess ? new TrackerAccessJob(this) : null;
  }

  get = async (path, headers) => spGet(this.webUrl, await this.lists.rewrite(path), this.fetchImpl, headers);
  post = async (path, options) => spPost(this.webUrl, await this.lists.rewrite(path), options, this.fetchImpl);
  listApi = key => this.lists.path(key);
  organizeTrackerLists = onProgress => this.lists.organize(onProgress);
  maintainListNames = () => this.listNameMaintenance ||= this.lists.maintainNames();
  currentUser = () => this.userPromise ||= getCurrentUser(this.webUrl, this.fetchImpl);

  async listExists(key) {
    try { await this.get(`${apiFor(this.prefix, key)}?$select=Id`); return true; }
    catch (error) { if (error instanceof SharePointError && error.status === 404) return false; throw error; }
  }

  async readiness() {
    const checks = await Promise.all(CONTAINERS.map(async (container) => {
      const exists = await this.listExists(container.key);
      if (!exists) return { key: container.key, exists, missingFields: container.fields.map((field) => field.name) };
      const body = await this.get(`${apiFor(this.prefix, container.key)}/fields?$select=InternalName&$top=500`);
      const existing = new Set((body.value || body.d?.results || []).map((field) => field.InternalName));
      return { key: container.key, exists, missingFields: container.fields.filter((field) => !existing.has(field.name)).map((field) => field.name) };
    }));
    return { ready: checks.every((check) => check.exists && check.missingFields.length === 0), checks };
  }

  async provision() {
    // An unreadable/security-trimmed list can look missing to readiness().
    // Never turn that into schema writes by a viewer, engineer, or list manager.
    // Check actual site permissions, independently of the app/directory role.
    const response = await this.get('/_api/web/EffectiveBasePermissions');
    const permissions = response?.d || response;
    const low = Number((permissions?.EffectiveBasePermissions || permissions)?.Low);
    const setupPermissions = 2048 | 33554432; // Manage Lists + Manage Permissions
    if (!Number.isInteger(low) || low < 0 || low > 0xffffffff || (low & setupPermissions) !== setupPermissions) {
      throw new Error('Tracker lists or required fields are unavailable to your account. Ask a site owner to check your Read access to all seven Tracker lists and confirm the configured site and list names. If setup is needed, an administrator with site-level Manage Lists and Manage Permissions must open the app. No setup changes were made; ordinary users do not need site-wide Edit.');
    }
    const steps = [];
    for (const container of CONTAINERS) {
      if (!(await this.listExists(container.key))) {
        await this.post('/_api/web/lists', { body: { Title: displayListTitle(this.prefix, container.key), Description: container.description, BaseTemplate: 100, ...(['tasks', 'references'].includes(container.key) ? { EnableAttachments: true } : {}), AllowContentTypes: false, ContentTypesEnabled: false, Hidden: false } });
        this.lists.cache.delete(container.key);
        steps.push(`Created ${displayListTitle(this.prefix, container.key)}`);
      }
      const body = await this.get(`${apiFor(this.prefix, container.key)}/fields?$select=InternalName&$top=500`);
      const existing = new Set((body.value || []).map((field) => field.InternalName));
      const seedMarker = container.key === 'acronyms' ? container.fields.find((field) => field.name === 'SeedVersion') : null;
      const addField = async (field) => {
        await this.post(`${apiFor(this.prefix, container.key)}/fields/createfieldasxml`, {
          verbose: true,
          body: { parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: schemaXml(field), Options: ADD_FIELD.INTERNAL_NAME_HINT | (field.inView ? ADD_FIELD.TO_DEFAULT_VIEW : 0) } },
        });
        steps.push(`Added ${container.key}.${field.name}`);
      };
      for (const field of container.fields) {
        if (existing.has(field.name) || field === seedMarker) continue;
        await addField(field);
      }
      if (seedMarker && !existing.has(seedMarker.name)) {
        const fieldsBeforeMarker = container.fields.filter((field) => field !== seedMarker).map((field) => field.name);
        const existingRows = await this.listItems('acronyms', fieldsBeforeMarker, fromAcronym);
        const existingAcronyms = new Set(existingRows.map((entry) => String(entry.acronym || '').toUpperCase()));
        const missing = defaultAcronyms.filter((entry) => !existingAcronyms.has(entry.acronym.toUpperCase()));
        for (const entry of missing) {
          const fields = acronymFields(entry);
          delete fields.SeedVersion;
          await this.create('acronyms', fields);
        }
        if (missing.length) steps.push(`Added ${missing.length} default acronyms`);
        await addField(seedMarker);
      }
    }
    return steps;
  }

  async listItems(key, fields, converter, includeArchived = false) {
    const select = ['Id', 'Title', 'FileSystemObjectType', ...fields].join(',');
    let path = `${apiFor(this.prefix, key)}/items?$select=${select}&$top=5000`;
    const rows = [];
    while (path) {
      const body = await this.get(path);
      rows.push(...(body.value || body.d?.results || []));
      const next = body['@odata.nextLink'] || body.d?.__next || '';
      if (!next) path = '';
      else if (next.startsWith(this.webUrl)) path = next.slice(this.webUrl.length);
      else if (/^https?:/i.test(next)) {
        const parsed = new URL(next);
        path = `${parsed.pathname}${parsed.search}`;
      } else path = next;
    }
    return rows.filter((row) => row.FileSystemObjectType !== 1 && (includeArchived || !row.Archived)).map(converter);
  }

  async load() {
    const [projects, tasks, updates, risks, acronyms, users] = await Promise.all([
      this.listItems('projects', CONTAINERS[0].fields.map((field) => field.name), fromProject),
      this.listItems('tasks', CONTAINERS[1].fields.map((field) => field.name), fromTask),
      this.listItems('updates', CONTAINERS[2].fields.map((field) => field.name), fromUpdate),
      this.listItems('risks', CONTAINERS[3].fields.map((field) => field.name), fromRisk),
      this.listItems('acronyms', CONTAINERS[4].fields.map((field) => field.name), fromAcronym),
      this.listItems('users', CONTAINERS[5].fields.map((field) => field.name), fromUser),
    ]);
    const directory = normalizeDirectory(users, projects);
    if (this.permissions) {
      const current = await this.currentUser();
      const { role } = await this.permissions.currentRole(current);
      const entry = directory.find(row => isOwnedByUser({ ownerKey: row.loginName, ownerEmail: row.email }, current));
      if (entry) { entry.role = role; delete entry.roleMigrationPending; }
      else directory.push({ id: `session-${current.id}`, title: current.title, email: current.email, loginName: current.loginName, role, sessionOnly: true });
    }
    return { projects, tasks: normalizeTaskOrganizations(tasks, projects), updates, risks, acronyms, users: directory };
  }

  async create(key, fields) {
    if (this.permissions && ['tasks', 'updates', 'risks'].includes(key)) return this.permissions.createChild(key, fields);
    const created = await this.post(`${apiFor(this.prefix, key)}/items`, { body: fields });
    return created?.Id;
  }

  formValues(fields) {
    // ValidateUpdateListItem parses field values as if they came from a
    // SharePoint edit form. Unlike normal REST item payloads, its DateTime
    // parser rejects ISO-8601 strings, so date fields use the site's standard
    // numeric form representation while empty values remain clearable.
    const serialize = (fieldName, value) => DATE_FIELDS.has(fieldName)
      ? sharePointFormDate(value)
      : value == null ? '' : typeof value === 'boolean' ? (value ? '1' : '0') : String(value);
    return Object.entries(fields).map(([FieldName, value]) => ({ FieldName, FieldValue: serialize(FieldName, value) }));
  }

  async update(key, spId, fields) {
    const result = await this.post(`${apiFor(this.prefix, key)}/items(${spId})/validateupdatelistitem`, {
      body: {
        formValues: this.formValues(fields),
        bNewDocumentUpdate: true,
      },
    });
    const rows = result?.value || result?.ValidateUpdateListItem?.results || result?.ValidateUpdateListItem || result?.d?.ValidateUpdateListItem?.results || [];
    const failed = rows.find((row) => row.HasException || row.ErrorMessage);
    if (failed) throw new SharePointError(`SharePoint rejected ${failed.FieldName || 'a field'}: ${failed.ErrorMessage || 'validation failed'}`, 400, failed);
  }

  async recycle(key, spId) {
    if (this.permissions && key === 'users') {
      const users = await this.directoryRows();
      const person = users.find(row => row.spId === spId);
      if (!person) throw new Error('This user no longer exists.');
      await this.syncUserAccess(person, null, users);
    }
    if (this.permissions && key === 'projects') {
      const project = (await this.permissionProjects()).find(row => row.spId === spId);
      if (project) await this.permissions.syncProject({ ...project, archived: true }, []);
    }
    // Like Uncertalytics, remove from the app through a normal metadata update.
    await this.update(key, spId, { Archived: true });
  }

  async listTaskAttachments(task) {
    if (!Number.isInteger(task.spId) || task.spId <= 0) throw new Error('Save the task before attaching documents.');
    const body = await this.get(`${apiFor(this.prefix, 'tasks')}/items(${task.spId})/AttachmentFiles?$select=FileName,ServerRelativeUrl`);
    const archived = await this.archivedDocuments(task);
    return (body.value || body.d?.results || []).filter((file) => !archived.includes(file.FileName)).map((file) => {
      const url = new URL(file.ServerRelativeUrl, this.webUrl);
      if (url.origin !== new URL(this.webUrl).origin) throw new Error('SharePoint returned an unexpected attachment address.');
      return { name: file.FileName, url: url.href };
    });
  }

  async archivedDocuments(task) {
    const response = await this.get(`${apiFor(this.prefix, 'tasks')}/items(${task.spId})?$select=ArchivedDocuments`);
    const names = safeJson((response.d || response).ArchivedDocuments, []);
    if (!Array.isArray(names)) throw new Error('Document archive metadata is invalid.');
    return names;
  }

  async downloadTaskAttachment(task, name) {
    const file = (await this.listTaskAttachments(task)).find((entry) => entry.name === name);
    if (!file) throw new Error('This document no longer exists. Refresh the documents list.');
    const response = await this.fetchImpl(file.url, { credentials: 'include', headers: { Accept: '*/*' } });
    if (!response.ok) throw new Error(`Document download failed (${response.status}).`);
    return response.blob();
  }

  async deleteTaskAttachment(task, name) {
    if (!(await this.listTaskAttachments(task)).some((entry) => entry.name === name)) throw new Error('This document no longer exists. Refresh the documents list.');
    const archived = await this.archivedDocuments(task);
    await this.update('tasks', task.spId, { ArchivedDocuments: JSON.stringify([...new Set([...archived, name])]) });
    if (!(await this.archivedDocuments(task)).includes(name)) throw new Error('SharePoint did not confirm the document deletion. Refresh and try again.');
  }

  async addTaskAttachment(task, file) {
    validateAttachment(file);
    const files = await this.listTaskAttachments(task);
    if ((await this.archivedDocuments(task)).some((name) => name.toLowerCase() === file.name.toLowerCase())) throw new Error('A deleted document uses this name. Rename the file before uploading.');
    if (files.some((entry) => entry.name.toLowerCase() === file.name.toLowerCase())) throw new Error('A document with this name is already attached. Rename the new file before uploading.');
    const settings = await this.get(`${apiFor(this.prefix, 'tasks')}?$select=EnableAttachments`);
    if ((settings.d || settings).EnableAttachments === false) throw new Error('Attachments are disabled on the tasks list. Ask a site owner to enable list attachments.');
    await this.post(`${apiFor(this.prefix, 'tasks')}/items(${task.spId})/AttachmentFiles/add(FileName='${escapeOData(file.name)}')`, { raw: true, headers: { 'Content-Type': 'application/octet-stream' }, body: await file.arrayBuffer() });
    return this.listTaskAttachments(task);
  }

  async renameTaskAttachment(task, name, nextName) {
    const cleanName = validateAttachmentName(nextName);
    const current = (await this.listTaskAttachments(task)).find((entry) => entry.name === name);
    if (!current) throw new Error('This document no longer exists. Refresh the documents list.');
    if (current.name.toLowerCase() === cleanName.toLowerCase()) return this.listTaskAttachments(task);
    const blob = await this.downloadTaskAttachment(task, name);
    await this.addTaskAttachment(task, new File([blob], cleanName, { type: blob.type, lastModified: Date.now() }));
    await this.deleteTaskAttachment(task, name);
    return this.listTaskAttachments(task);
  }

  async listReferenceEntries() {
    return this.listItems('references', ['RecordId', 'ReferenceParentId', 'EntryKind', 'FileName', 'FileSize', 'Archived'], (item) => ({
      id: item.RecordId, spId: item.Id, name: item.Title, parentId: item.ReferenceParentId || '',
      kind: item.EntryKind, fileName: item.FileName || '', size: item.FileSize || 0,
    }));
  }

  async saveReferenceEntry(row, file) {
    const next = { ...row, id: row.id || `ref-${crypto.randomUUID()}` };
    const fields = { Title: next.name, RecordId: next.id, ReferenceParentId: next.parentId, EntryKind: next.kind, FileName: next.fileName || '', FileSize: next.size || 0 };
    if (next.spId) await this.update('references', next.spId, fields);
    else {
      // Incomplete uploads stay hidden, and retries never replace another file.
      next.spId = await this.create('references', { ...fields, Archived: !!file });
      if (!next.spId) throw new Error('SharePoint did not return a reference item ID.');
      if (file) {
        await this.post(`${apiFor(this.prefix, 'references')}/items(${next.spId})/AttachmentFiles/add(FileName='${escapeOData(file.name)}')`, { raw: true, headers: { 'Content-Type': 'application/octet-stream' }, body: await file.arrayBuffer() });
        await this.update('references', next.spId, { Archived: false });
      }
    }
    const saved = (await this.listReferenceEntries()).find((entry) => entry.id === next.id);
    if (!saved || saved.name !== next.name || saved.parentId !== next.parentId) throw new Error('SharePoint did not confirm the reference document changes.');
    return saved;
  }

  async deleteReferenceEntry(row) {
    await this.update('references', row.spId, { Archived: true });
    const response = await this.get(`${apiFor(this.prefix, 'references')}/items(${row.spId})?$select=Archived`);
    if ((response.d || response).Archived !== true) throw new Error('SharePoint did not confirm reference deletion.');
  }

  async downloadReferenceEntry(row) {
    const body = await this.get(`${apiFor(this.prefix, 'references')}/items(${row.spId})/AttachmentFiles?$select=FileName,ServerRelativeUrl`);
    const file = (body.value || body.d?.results || []).find((entry) => entry.FileName === row.fileName);
    if (!file) throw new Error('This reference document is unavailable.');
    const url = new URL(file.ServerRelativeUrl, this.webUrl);
    if (url.origin !== new URL(this.webUrl).origin) throw new Error('Unexpected reference document address.');
    const response = await this.fetchImpl(url.href, { credentials: 'include', headers: { Accept: '*/*' } });
    if (!response.ok) throw new Error(`Document download failed (${response.status}).`);
    return response.blob();
  }

  async verifyInvitationRead(scope, loginName, required, label) {
    try {
      const response = await this.get(`${scope}/getusereffectivepermissions(@u)?@u='${escapeOData(loginName)}'`);
      const permissions = response?.d?.GetUserEffectivePermissions || response?.GetUserEffectivePermissions || response?.d || response;
      const raw = permissions?.Low;
      const low = (typeof raw === 'string' && /^\d+$/.test(raw)) || typeof raw === 'number' ? Number(raw) : NaN;
      if (!Number.isInteger(low) || low < 0 || low > 0xffffffff || (low & required) !== required) throw new Error('Required Read access was not confirmed.');
    } catch (error) {
      throw new Error(`Tracker access could not be verified for ${label}. Group membership may already be saved. Ask a site owner to check this resource's permissions. No invitation email was requested. ${error.message}`);
    }
  }

  async shareSiteAccess(person, role, appUrl) {
    if (!ROLES.includes(role) || !person.loginName) throw new Error('Select a resolved person and a valid role.');
    const link = new URL(trackerPageUrl(appUrl, this.webUrl));
    const fileApi = `/_api/web/GetFileByServerRelativePath(decodedurl='${escapeOData(decodeURIComponent(link.pathname))}')`;
    let page;
    try { const response = await this.get(`${fileApi}?$select=Exists,Level`); page = response.d || response; }
    catch (error) { throw new Error(`The tracker page could not be checked. Confirm its direct URL and your access. ${error.message}`); }
    if (page.Exists !== true) throw new Error('The tracker page was not found. Use its direct published URL.');
    if (page.Level !== 1) throw new Error('Publish or republish the tracker page in SharePoint before inviting users. The current page is a draft or checked out.');
    // Baseline app membership is managed outside Tracker. Viewer has no role group.
    let group;
    if (this.permissions) {
      group = await this.permissions.syncGroups(person, role);
    } else {
      const groupResponse = await this.get('/_api/web/associatedmembergroup?$select=Id,Title');
      group = groupResponse.d || groupResponse;
      if (!Number.isInteger(group.Id) || group.Id <= 0) throw new Error('The site Members group is unavailable. Ask a site owner to configure it.');
      const membersApi = `/_api/web/sitegroups(${group.Id})/users`;
      const membershipUrl = `${membersApi}?$select=Id,LoginName&$filter=LoginName eq '${escapeOData(person.loginName)}'`;
      const isMember = async () => {
        const response = await this.get(membershipUrl);
        return (response.value || response.d?.results || []).some((user) => user.LoginName?.toLowerCase() === person.loginName.toLowerCase());
      };
      try {
        if (!(await isMember())) await this.post(membersApi, { body: { LoginName: person.loginName } });
        if (!(await isMember())) throw new Error('SharePoint did not confirm group membership.');
      } catch (error) {
        throw new Error(`Could not verify membership in ${group.Title || 'the site Members group'}. The inviting account must be allowed to manage this group. No invitation email was requested. ${error.message}`);
      }
    }
    const readDefinitions = await this.get('/_api/web/roledefinitions?$select=Id,RoleTypeKind&$filter=RoleTypeKind eq 2');
    const readRole = (readDefinitions.value || readDefinitions.d?.results || []).find((entry) => entry.RoleTypeKind === 2);
    if (!readRole?.Id) throw new Error('The Read permission level is unavailable. Ask a site owner.');
    if (!this.permissions) {
      // Retain the legacy deployment's Members-based invitation behavior.
      await this.verifyInvitationRead('/_api/web', person.loginName, 1, 'the legacy site');
    }
    await this.verifyInvitationRead(`${fileApi}/ListItemAllFields`, person.loginName, 33, 'the published Tracker page');
    if (this.permissions) {
      for (const container of CONTAINERS) {
        await this.verifyInvitationRead(await this.listApi(container.key), person.loginName, 1, displayListTitle(this.prefix, container.key));
      }
      if (!Array.isArray(this.invitationAssetFolders)) throw new Error('Configure invitationAssetFolders as an array of Tracker folder URLs. No invitation email was requested.');
      for (const value of [...new Set(this.invitationAssetFolders)]) {
        const folder = new URL(normalizeInvitationUrl(value, this.webUrl));
        const path = `/_api/web/GetFolderByServerRelativePath(decodedurl='${escapeOData(decodeURIComponent(folder.pathname).replace(/\/$/, ''))}')/ListItemAllFields`;
        await this.verifyInvitationRead(path, person.loginName, 33, `Tracker assets folder ${folder.pathname}`);
      }
      if (!Array.isArray(this.invitationAssetUrls)) throw new Error('Configure invitationAssetUrls as an array of Tracker file URLs. No invitation email was requested.');
      for (const value of [...new Set(this.invitationAssetUrls)]) {
        const asset = new URL(normalizeInvitationUrl(value, this.webUrl));
        const path = `/_api/web/GetFileByServerRelativePath(decodedurl='${escapeOData(decodeURIComponent(asset.pathname))}')/ListItemAllFields`;
        await this.verifyInvitationRead(path, person.loginName, 33, `Tracker asset ${asset.pathname}`);
      }
    }
    // Flank Speed has retired SP.Utilities.Utility.SendEmail. Make exactly one
    // page-targeted ShareObject request after membership and access pass.
    // ShareObject renders this field as plain text in Flank Speed invitation
    // emails, so keep it free of URLs and markup. The generated card and Open
    // button below the message provide the direct tracker link.
    const emailBody = `You have been added to the NPSL Modernization Tracker. Assigned role: ${role.toUpperCase()}.`;
    let response;
    try { response = await this.post('/_api/SP.Web.ShareObject', { body: {
      url: link.href,
      peoplePickerInput: JSON.stringify([{ Key: person.loginName }]),
      roleValue: 'role:' + readRole.Id,
      groupId: 0, propagateAcl: false, sendEmail: true,
      includeAnonymousLinkInEmail: false,
      emailSubject: 'Invitation to the NPSL Modernization Tracker',
      emailBody,
      useSimplifiedRoles: false,
    } }); } catch (error) {
      if (this.permissions) throw new Error(`Tracker group membership was verified, but the invitation email was not confirmed. A site owner may need to send the page invitation if your account cannot share it. Retry only the invitation. ${error.message}`);
      throw error;
    }
    const result = response?.d?.ShareObject || response?.ShareObject || response?.d || response;
    if (result?.StatusCode !== 0 || result?.ErrorMessage) throw new Error(result?.ErrorMessage || 'SharePoint did not confirm the invitation email request.');
    return { access: group?.Title || (this.permissions ? 'Viewer (existing app access)' : 'Site Members'), emailRequested: true };
  }

  async searchPeople(query) {
    if (query.trim().length < 2) return [];
    const result = await this.post('/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser', {
      body: { queryParams: { AllowEmailAddresses: true, AllowMultipleEntities: false, AllUrlZones: false, MaximumEntitySuggestions: 15, PrincipalSource: 15, PrincipalType: 1, QueryString: query.trim() } },
    });
    return parsePeopleResults(result);
  }

  async resolvePerson(loginName) {
    if (!loginName?.trim()) throw new Error('Select a person from the directory.');
    const response = await this.post('/_api/web/ensureuser', { body: { logonName: loginName } });
    const person = response.d || response;
    if (!person.LoginName || !person.Title) throw new Error('SharePoint could not resolve this person. Search and select them again.');
    return { title: person.Title, email: person.Email || '', loginName: person.LoginName };
  }

  async saveUser(row) {
    if (this.permissions) {
      const users = await this.directoryRows();
      const existing = users.find(entry => entry.id === row.id);
      if (existing && existing.loginName.toLowerCase() !== row.loginName.toLowerCase()) throw new Error('An existing SharePoint identity cannot be changed. Invite the new identity separately.');
      // Session-only owner entries become real directory records when explicitly saved.
      row = { ...row, spId: existing?.spId || row.spId };
      await this.syncUserAccess(row, row.role, users);
    }
    const fields = { Title: row.title, RecordId: row.id, LoginKey: row.loginName, Email: row.email || '', AppRole: row.role };
    let spId = row.spId;
    if (spId) await this.update('users', spId, fields);
    else spId = await this.create('users', fields);
    if (!spId) throw new Error('SharePoint did not return a user record ID. Your role could not be verified.');
    const response = await this.get(`${apiFor(this.prefix, 'users')}/items(${spId})?$select=Id,RecordId,Title,LoginKey,Email,AppRole`);
    const item = response?.d || response;
    if (item?.RecordId !== row.id || item?.LoginKey !== row.loginName || item?.AppRole !== row.role) {
      throw new Error(`SharePoint did not confirm the requested ${row.role} role. Check edit permission on ${titleFor(this.prefix, 'users')} and try again. Current stored role: ${item?.AppRole || 'unknown'}.`);
    }
    return fromUser(item);
  }

  async saveProject(row) {
    if (!this.permissions) return row.spId ? (await this.update('projects', row.spId, projectFields(row)), row) : { ...row, spId: await this.create('projects', projectFields(row)) };
    const { role } = await this.permissions.currentRole(await this.currentUser());
    // A retry after a partial permission update reuses the already-saved project.
    const existing = (await this.permissionProjects()).find(project => project.id === row.id);
    const saved = { ...row, spId: existing?.spId || row.spId };
    if (saved.spId) await this.update('projects', saved.spId, projectFields(saved));
    else saved.spId = await this.create('projects', projectFields(saved));
    if (role === 'Manager') {
      try { await this.permissions.syncProject(saved, await this.directoryRows()); }
      catch (error) { throw new Error(`Project saved, but its access setup is incomplete. Retry Save or Apply tracker permissions. ${error.message}`); }
    }
    return saved;
  }
  async saveTask(row) { return row.spId ? (await this.update('tasks', row.spId, taskFields(row)), row) : { ...row, spId: await this.create('tasks', taskFields(row)) }; }
  async saveTasks(rows) { return Promise.all(rows.map((row) => this.saveTask(row))); }
  async saveUpdate(row) { return row.spId ? (await this.update('updates', row.spId, updateFields(row)), row) : { ...row, spId: await this.create('updates', updateFields(row)) }; }
  async saveRisk(row) { return row.spId ? (await this.update('risks', row.spId, riskFields(row)), row) : { ...row, spId: await this.create('risks', riskFields(row)) }; }
  async saveAcronym(row) { return row.spId ? (await this.update('acronyms', row.spId, acronymFields(row)), row) : { ...row, spId: await this.create('acronyms', acronymFields(row)) }; }

  directoryRows() { return this.listItems('users', CONTAINERS[5].fields.map(field => field.name), fromUser); }
  permissionProjects() { return this.listItems('projects', CONTAINERS[0].fields.map(field => field.name), item => ({ ...fromProject(item), archived: !!item.Archived }), true); }
  async syncUserAccess(person, role, users) {
    const existing = users.find(row => row.id === person.id);
    const next = [...users.filter(row => row.id !== person.id), ...(role ? [{ ...person, role }] : [])];
    await this.permissions.syncGroups(person, role, { beforeChange: async memberships => {
      const liveEngineer = memberships.some(entry => entry.name === 'Project Engineer' && entry.member);
      const liveManagerDemotion = role === 'Viewer' && memberships.some(entry => entry.name === 'Manager' && entry.member);
      const cleanNonEngineer = !existing || existing.role === 'Viewer' || (existing.role === 'Manager' && role === 'Manager');
      // Viewer → Manager and ordinary non-engineer profile saves are group-only.
      // Demotions, removals, legacy roles, and actual engineer membership still
      // check existing grants, including those left by interrupted changes.
      if (role !== 'Project Engineer' && (!role || !cleanNonEngineer || liveEngineer || liveManagerDemotion)) {
        await this.permissions.revokeEngineerAccess(person, memberships);
      }
    } });
    if (role === 'Project Engineer') {
      const projects = await this.permissionProjects();
      for (const project of projects.filter(project => isOwnedByUser(project, person))) await this.permissions.syncProject(project, next);
    }
  }
  async prepareTrackerAccess(onProgress = () => {}) {
    if (!this.permissions) throw new Error('Scoped tracker access is not enabled on this site.');
    const { role, siteOwner } = await this.permissions.currentRole(await this.currentUser());
    if (role !== 'Manager') throw new Error('Only tracker managers or site owners can apply permissions.');
    const baseline = await this.permissions.baselineGrants();
    if (siteOwner) await this.permissions.enableFolders();
    else for (const key of ['tasks', 'updates', 'risks']) if (!(await this.permissions.metadata(key)).EnableFolderCreation) throw new Error('A site owner must run this setup once to enable project folders.');
    const users = normalizeDirectory(await this.directoryRows(), await this.permissionProjects());
    for (const person of users) {
      onProgress(`Checking group membership: ${person.title}`);
      await this.permissions.syncGroups(person, person.role);
      if (person.roleMigrationPending) await this.update('users', person.spId, { AppRole: person.role });
    }
    const projects = await this.permissionProjects();
    for (const project of projects) await this.permissions.syncProject(project, users, onProgress, baseline);
    return { users: users.length, projects: projects.length };
  }

}

export { titleFor };
