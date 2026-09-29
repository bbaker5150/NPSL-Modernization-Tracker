import { normalizeInvitationUrl, parsePeopleResults } from './peoplePicker';
import { validateAttachment } from './taskAttachments';
import { normalizePhaseKey } from '../data/workflow';
import { getCurrentUser, SharePointError, spGet, spPost } from './spContext';
import { defaultAcronyms } from '../data/defaultAcronyms';

const FIELD = { TEXT: 'Text', NOTE: 'Note', NUMBER: 'Number', DATE: 'DateTime', BOOLEAN: 'Boolean' };
const ADD_FIELD = { INTERNAL_NAME_HINT: 8, TO_DEFAULT_VIEW: 16 };

export const CONTAINERS = [
  {
    key: 'projects', suffix: 'Projects', description: 'Modernization project portfolio records.', fields: [
      ['RecordId', 'Record ID', FIELD.TEXT, true], ['ProjectKey', 'Project Key', FIELD.TEXT, true],
      ['MeasurementArea', 'Measurement Area', FIELD.TEXT], ['Description', 'Description', FIELD.NOTE],
      ['OwnerName', 'Owner', FIELD.TEXT], ['OwnerEmail', 'Owner Email', FIELD.TEXT], ['OwnerKey', 'Owner Identity Key', FIELD.TEXT, true],
      ['ManagerName', 'Manager', FIELD.TEXT], ['ManagerEmail', 'Manager Email', FIELD.TEXT],
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
      ['EstimatedHours', 'Est. Hours', FIELD.NUMBER], ['SortOrder', 'Sort Order', FIELD.NUMBER], ['TaskStatus', 'Status', FIELD.TEXT],
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
].map((container) => ({
  ...container,
  fields: [...container.fields, ['Archived', 'Archived', FIELD.BOOLEAN]].map(([name, title, type, indexed = false]) => ({ name, title, type, indexed, inView: type !== FIELD.NOTE })),
}));

const escapeXml = (value) => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const escapeOData = (value) => encodeURIComponent(String(value).replace(/'/g, "''"));
const titleFor = (prefix, key) => `${String(prefix || 'Modernization').replace(/[^A-Za-z0-9]/g, '') || 'Modernization'}${CONTAINERS.find((c) => c.key === key).suffix}`;
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
  Description: row.description, OwnerName: row.ownerName, OwnerEmail: row.ownerEmail, OwnerKey: row.ownerKey || '', ManagerName: row.managerName,
  ManagerEmail: row.managerEmail || '', Priority: row.priority, Health: row.health, ProjectStatus: row.status,
  CurrentStageKey: row.currentStageKey, PercentComplete: row.percentComplete, ProgressMode: row.progressMode || 'phases', TargetFinish: sharePointDate(row.targetFinish),
  NextMilestone: row.nextMilestone, NextMilestoneDate: sharePointDate(row.nextMilestoneDate), SourceNotes: row.sourceNotes,
  ImportedBaseline: !!row.importedBaseline, TagsJson: JSON.stringify(row.tags || []),
});

const taskFields = (row) => ({
  Title: row.title, RecordId: row.id, ProjectKey: row.projectKey,
  TaskTitle: row.title, PhaseKey: row.phaseKey, SortOrder: row.order, TaskStatus: row.status,
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
    ownerEmail: item.OwnerEmail || '', ownerKey: item.OwnerKey || '', managerName: item.ManagerName || 'Unassigned', managerEmail: item.ManagerEmail || '',
    priority: item.Priority || 'Medium', health: item.Health || 'Needs Review', status: item.ProjectStatus || 'Planned',
    currentStageKey: phaseKey(item.CurrentStageKey), percentComplete: Number(item.PercentComplete || 0), progressMode: item.ProgressMode === 'tasks' ? 'tasks' : 'phases',
    targetFinish: dateOnly(item.TargetFinish), nextMilestone: item.NextMilestone || '', nextMilestoneDate: dateOnly(item.NextMilestoneDate),
    sourceNotes: item.SourceNotes || '', importedBaseline: !!item.ImportedBaseline, tags: safeJson(item.TagsJson, []),
  };
}

function fromTask(item) {
  const status = item.TaskStatus || 'Not Started';
  return {
    spId: item.Id, id: item.RecordId, projectKey: item.ProjectKey, title: item.TaskTitle,
    phaseKey: phaseKey(item.PhaseKey), order: Number(item.SortOrder || 0), status,
    estimatedHours: item.EstimatedHours == null ? null : Number(item.EstimatedHours), assignedDate: dateOnly(item.AssignedDate), startDate: dateOnly(item.StartDate), finishDate: dateOnly(item.FinishDate), dueDate: dateOnly(item.DueDate), deferredDate: dateOnly(item.DeferredDate), deferredJustification: item.DeferredJustification || '', notRequiredJustification: item.NotRequiredJustification || '',
    ownerName: item.OwnerName || 'Unassigned', ownerEmail: item.OwnerEmail || '', ownerKey: item.OwnerKey || '', notes: item.Notes || '',
    blockedReason: item.BlockedReason || '', sourceStartLabel: item.SourceStartLabel || '', dataIssue: item.DataIssue || '',
  };
}

const fromUser = (item) => ({ spId: item.Id, id: item.RecordId, title: item.Title, loginName: item.LoginKey, email: item.Email || '', role: ['Manager', 'SME'].includes(item.AppRole) ? item.AppRole : 'User' });

const fromUpdate = (item) => ({ spId: item.Id, id: item.RecordId, projectKey: item.ProjectKey, type: item.UpdateType, summary: item.Summary, entryDate: dateOnly(item.EntryDate), authorName: item.AuthorName, authorEmail: item.AuthorEmail, authorKey: item.AuthorKey || '' });
const fromRisk = (item) => ({ spId: item.Id, id: item.RecordId, projectKey: item.ProjectKey, title: item.RiskTitle || item.Title, severity: item.Severity, probability: item.Probability, mitigation: item.Mitigation || '', ownerName: item.OwnerName || '', ownerKey: item.OwnerKey || '', status: item.RiskStatus || 'Open', dueDate: dateOnly(item.DueDate) });
const fromAcronym = (item) => ({ spId: item.Id, id: item.RecordId, acronym: item.Acronym || item.Title, term: item.FullTerm || '', definition: item.Definition || '', seedVersion: item.SeedVersion || '' });

export class SharePointStore {
  constructor({ webUrl, prefix = 'Modernization', fetchImpl = fetch, hideLists = true }) {
    this.webUrl = String(webUrl || '').replace(/\/+$/, '');
    this.prefix = prefix;
    this.fetchImpl = fetchImpl;
    this.hideLists = hideLists;
    this.userPromise = null;
  }

  get = (path) => spGet(this.webUrl, path, this.fetchImpl);
  post = (path, options) => spPost(this.webUrl, path, options, this.fetchImpl);
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
    const steps = [];
    for (const container of CONTAINERS) {
      if (!(await this.listExists(container.key))) {
        await this.post('/_api/web/lists', { body: { Title: titleFor(this.prefix, container.key), Description: container.description, BaseTemplate: 100, ...(container.key === 'tasks' ? { EnableAttachments: true } : {}), AllowContentTypes: false, ContentTypesEnabled: false, Hidden: this.hideLists } });
        steps.push(`Created ${titleFor(this.prefix, container.key)}`);
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

  async listItems(key, fields, converter) {
    const select = ['Id', 'Title', ...fields].join(',');
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
    return rows.filter((row) => !row.Archived).map(converter);
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
    return { projects, tasks, updates, risks, acronyms, users };
  }

  async create(key, fields) {
    const created = await this.post(`${apiFor(this.prefix, key)}/items`, { body: fields });
    return created?.Id;
  }

  async update(key, spId, fields) {
    // ValidateUpdateListItem parses field values as if they came from a
    // SharePoint edit form. Unlike normal REST item payloads, its DateTime
    // parser rejects ISO-8601 strings, so date fields use the site's standard
    // numeric form representation while empty values remain clearable.
    const serialize = (fieldName, value) => DATE_FIELDS.has(fieldName)
      ? sharePointFormDate(value)
      : value == null ? '' : typeof value === 'boolean' ? (value ? '1' : '0') : String(value);
    const result = await this.post(`${apiFor(this.prefix, key)}/items(${spId})/validateupdatelistitem`, {
      body: {
        formValues: Object.entries(fields).map(([FieldName, value]) => ({ FieldName, FieldValue: serialize(FieldName, value) })),
        bNewDocumentUpdate: true,
      },
    });
    const rows = result?.value || result?.ValidateUpdateListItem?.results || result?.ValidateUpdateListItem || result?.d?.ValidateUpdateListItem?.results || [];
    const failed = rows.find((row) => row.HasException || row.ErrorMessage);
    if (failed) throw new SharePointError(`SharePoint rejected ${failed.FieldName || 'a field'}: ${failed.ErrorMessage || 'validation failed'}`, 400, failed);
  }

  async recycle(key, spId) {
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

  async shareSiteAccess(person, role, appUrl) {
    if (!['User', 'SME', 'Manager'].includes(role) || !person.loginName) throw new Error('Select a resolved person and a valid role.');
    const link = new URL(normalizeInvitationUrl(appUrl, this.webUrl));
    // Use Contribute, not Full Control or site-design Edit, for working roles.
    const kind = role === 'SME' ? 2 : 3;
    const definitions = await this.get(`/_api/web/roledefinitions?$select=Id,RoleTypeKind&$filter=RoleTypeKind eq ${kind}`);
    const definition = (definitions.value || definitions.d?.results || []).find((entry) => entry.RoleTypeKind === kind);
    if (!definition?.Id) throw new Error('The required SharePoint permission level is unavailable. Ask a site owner.');
    const response = await this.post('/_api/SP.Web.ShareObject', { body: {
      url: this.webUrl,
      peoplePickerInput: JSON.stringify([{ Key: person.loginName }]),
      roleValue: `role:${definition.Id}`, groupId: 0, propagateAcl: false,
      sendEmail: true, includeAnonymousLinkInEmail: false,
      emailSubject: 'Invitation to the NPSL Modernization Tracker',
      emailBody: `You have been added to the NPSL Modernization Tracker as ${role}. Open the tracker: ${link.href}`,
      useSimplifiedRoles: false,
    } });
    const result = response?.d?.ShareObject || response?.ShareObject || response?.d || response;
    if (result?.StatusCode !== 0 || result?.ErrorMessage) throw new Error(result?.ErrorMessage || 'SharePoint did not confirm sharing. The inviting manager must have permission to share this site.');
    const permissions = await this.get(`/_api/web/getusereffectivepermissions(@u)?@u='${escapeOData(person.loginName)}'`);
    const effective = permissions?.d?.GetUserEffectivePermissions || permissions?.GetUserEffectivePermissions || permissions?.d || permissions;
    const required = role === 'SME' ? 1 : 7;
    if (!Number.isFinite(Number(effective?.Low)) || (Number(effective.Low) & required) !== required) throw new Error('Sharing was accepted, but site access could not be verified. Ask a site owner to check permissions before retrying.');
    return { access: role === 'SME' ? 'Read' : 'Contribute', emailRequested: true };
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

  async saveProject(row) { return row.spId ? (await this.update('projects', row.spId, projectFields(row)), row) : { ...row, spId: await this.create('projects', projectFields(row)) }; }
  async saveTask(row) { return row.spId ? (await this.update('tasks', row.spId, taskFields(row)), row) : { ...row, spId: await this.create('tasks', taskFields(row)) }; }
  async saveTasks(rows) { return Promise.all(rows.map((row) => this.saveTask(row))); }
  async saveUpdate(row) { return { ...row, spId: await this.create('updates', updateFields(row)) }; }
  async saveRisk(row) { return row.spId ? (await this.update('risks', row.spId, riskFields(row)), row) : { ...row, spId: await this.create('risks', riskFields(row)) }; }
  async saveAcronym(row) { return row.spId ? (await this.update('acronyms', row.spId, acronymFields(row)), row) : { ...row, spId: await this.create('acronyms', acronymFields(row)) }; }

}

export { titleFor };
