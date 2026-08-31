import { getCurrentUser, SharePointError, spGet, spPost } from './spContext';

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
      ['CurrentStageKey', 'Current Stage', FIELD.TEXT], ['PercentComplete', 'Percent Complete', FIELD.NUMBER],
      ['TargetFinish', 'Target Finish', FIELD.DATE], ['NextMilestone', 'Next Milestone', FIELD.TEXT],
      ['NextMilestoneDate', 'Next Milestone Date', FIELD.DATE], ['SourceNotes', 'Source Notes', FIELD.NOTE],
      ['ImportedBaseline', 'Imported Baseline', FIELD.BOOLEAN], ['TagsJson', 'Tags', FIELD.NOTE],
    ],
  },
  {
    key: 'tasks', suffix: 'Tasks', description: 'WBS tasks for every modernization project.', fields: [
      ['RecordId', 'Record ID', FIELD.TEXT, true], ['ProjectKey', 'Project Key', FIELD.TEXT, true],
      ['WBS', 'WBS', FIELD.TEXT], ['TaskTitle', 'Task', FIELD.TEXT], ['PhaseKey', 'Phase', FIELD.TEXT],
      ['SortOrder', 'Sort Order', FIELD.NUMBER], ['TaskStatus', 'Status', FIELD.TEXT],
      ['StartDate', 'Start Date', FIELD.DATE], ['FinishDate', 'Finish Date', FIELD.DATE], ['DueDate', 'Due Date', FIELD.DATE],
      ['OwnerName', 'Owner', FIELD.TEXT], ['OwnerEmail', 'Owner Email', FIELD.TEXT], ['OwnerKey', 'Owner Identity Key', FIELD.TEXT, true], ['Notes', 'Notes', FIELD.NOTE],
      ['BlockedReason', 'Blocked Reason', FIELD.NOTE], ['SourceStartLabel', 'Source Start Label', FIELD.TEXT],
      ['DataIssue', 'Data Issue', FIELD.NOTE],
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
].map((container) => ({
  ...container,
  fields: container.fields.map(([name, title, type, indexed = false]) => ({ name, title, type, indexed, inView: type !== FIELD.NOTE })),
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

function toFormValues(fields) {
  return Object.entries(fields).map(([FieldName, value]) => ({
    FieldName,
    FieldValue: value === null || value === undefined ? '' : typeof value === 'boolean' ? (value ? '1' : '0') : String(value),
  }));
}

function assertUpdate(result, context) {
  const rows = result?.value || result?.d?.ValidateUpdateListItem?.results || [];
  const failures = rows.filter((row) => row?.HasException || row?.ErrorMessage);
  if (failures.length) throw new SharePointError(`${context}: ${failures.map((row) => row.ErrorMessage).join('; ')}`, 400);
}

const dateOnly = (value) => value ? String(value).slice(0, 10) : '';
const safeJson = (value, fallback) => {
  try { return value ? JSON.parse(value) : fallback; } catch { return fallback; }
};

const projectFields = (row) => ({
  Title: row.title, RecordId: row.id, ProjectKey: row.projectKey, MeasurementArea: row.measurementArea,
  Description: row.description, OwnerName: row.ownerName, OwnerEmail: row.ownerEmail, OwnerKey: row.ownerKey || '', ManagerName: row.managerName,
  ManagerEmail: row.managerEmail || '', Priority: row.priority, Health: row.health, ProjectStatus: row.status,
  CurrentStageKey: row.currentStageKey, PercentComplete: row.percentComplete, TargetFinish: row.targetFinish || null,
  NextMilestone: row.nextMilestone, NextMilestoneDate: row.nextMilestoneDate || null, SourceNotes: row.sourceNotes,
  ImportedBaseline: !!row.importedBaseline, TagsJson: JSON.stringify(row.tags || []),
});

const taskFields = (row) => ({
  Title: `${row.wbs} ${row.title}`, RecordId: row.id, ProjectKey: row.projectKey, WBS: row.wbs,
  TaskTitle: row.title, PhaseKey: row.phaseKey, SortOrder: row.order, TaskStatus: row.status,
  StartDate: row.startDate || null, FinishDate: row.finishDate || null, DueDate: row.dueDate || null,
  OwnerName: row.ownerName, OwnerEmail: row.ownerEmail, OwnerKey: row.ownerKey || '', Notes: row.notes, BlockedReason: row.blockedReason,
  SourceStartLabel: row.sourceStartLabel, DataIssue: row.dataIssue,
});

const updateFields = (row) => ({
  Title: `${row.projectKey} update`, RecordId: row.id, ProjectKey: row.projectKey, UpdateType: row.type,
  Summary: row.summary, EntryDate: row.entryDate, AuthorName: row.authorName, AuthorEmail: row.authorEmail, AuthorKey: row.authorKey || '',
});

const riskFields = (row) => ({
  Title: row.title, RecordId: row.id, ProjectKey: row.projectKey, RiskTitle: row.title, Severity: row.severity,
  Probability: row.probability, Mitigation: row.mitigation, OwnerName: row.ownerName, OwnerKey: row.ownerKey || '', RiskStatus: row.status,
  DueDate: row.dueDate || null,
});

function fromProject(item) {
  return {
    spId: item.Id, id: item.RecordId, projectKey: item.ProjectKey, title: item.Title,
    measurementArea: item.MeasurementArea || item.Title, description: item.Description || '', ownerName: item.OwnerName || 'Unassigned',
    ownerEmail: item.OwnerEmail || '', ownerKey: item.OwnerKey || '', managerName: item.ManagerName || 'Unassigned', managerEmail: item.ManagerEmail || '',
    priority: item.Priority || 'Medium', health: item.Health || 'Needs Review', status: item.ProjectStatus || 'Planned',
    currentStageKey: item.CurrentStageKey || 'need-scope', percentComplete: Number(item.PercentComplete || 0),
    targetFinish: dateOnly(item.TargetFinish), nextMilestone: item.NextMilestone || '', nextMilestoneDate: dateOnly(item.NextMilestoneDate),
    sourceNotes: item.SourceNotes || '', importedBaseline: !!item.ImportedBaseline, tags: safeJson(item.TagsJson, []),
  };
}

function fromTask(item) {
  return {
    spId: item.Id, id: item.RecordId, projectKey: item.ProjectKey, wbs: item.WBS, title: item.TaskTitle,
    phaseKey: item.PhaseKey, order: Number(item.SortOrder || 0), status: item.TaskStatus || 'Not Started',
    startDate: dateOnly(item.StartDate), finishDate: dateOnly(item.FinishDate), dueDate: dateOnly(item.DueDate),
    ownerName: item.OwnerName || 'Unassigned', ownerEmail: item.OwnerEmail || '', ownerKey: item.OwnerKey || '', notes: item.Notes || '',
    blockedReason: item.BlockedReason || '', sourceStartLabel: item.SourceStartLabel || '', dataIssue: item.DataIssue || '',
  };
}

const fromUpdate = (item) => ({ spId: item.Id, id: item.RecordId, projectKey: item.ProjectKey, type: item.UpdateType, summary: item.Summary, entryDate: dateOnly(item.EntryDate), authorName: item.AuthorName, authorEmail: item.AuthorEmail, authorKey: item.AuthorKey || '' });
const fromRisk = (item) => ({ spId: item.Id, id: item.RecordId, projectKey: item.ProjectKey, title: item.RiskTitle || item.Title, severity: item.Severity, probability: item.Probability, mitigation: item.Mitigation || '', ownerName: item.OwnerName || '', ownerKey: item.OwnerKey || '', status: item.RiskStatus || 'Open', dueDate: dateOnly(item.DueDate) });

export class SharePointStore {
  constructor({ webUrl, prefix = 'Modernization', fetchImpl = fetch }) {
    this.webUrl = String(webUrl || '').replace(/\/+$/, '');
    this.prefix = prefix;
    this.fetchImpl = fetchImpl;
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
        await this.post('/_api/web/lists', { body: { Title: titleFor(this.prefix, container.key), Description: container.description, BaseTemplate: 100, AllowContentTypes: false, ContentTypesEnabled: false } });
        steps.push(`Created ${titleFor(this.prefix, container.key)}`);
      }
      const body = await this.get(`${apiFor(this.prefix, container.key)}/fields?$select=InternalName&$top=500`);
      const existing = new Set((body.value || []).map((field) => field.InternalName));
      for (const field of container.fields) {
        if (existing.has(field.name)) continue;
        await this.post(`${apiFor(this.prefix, container.key)}/fields/createfieldasxml`, {
          verbose: true,
          body: { parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: schemaXml(field), Options: ADD_FIELD.INTERNAL_NAME_HINT | (field.inView ? ADD_FIELD.TO_DEFAULT_VIEW : 0) } },
        });
        steps.push(`Added ${container.key}.${field.name}`);
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
    return rows.map(converter);
  }

  async load() {
    const [projects, tasks, updates, risks] = await Promise.all([
      this.listItems('projects', CONTAINERS[0].fields.map((field) => field.name), fromProject),
      this.listItems('tasks', CONTAINERS[1].fields.map((field) => field.name), fromTask),
      this.listItems('updates', CONTAINERS[2].fields.map((field) => field.name), fromUpdate),
      this.listItems('risks', CONTAINERS[3].fields.map((field) => field.name), fromRisk),
    ]);
    return { projects, tasks, updates, risks };
  }

  async create(key, fields) {
    const created = await this.post(`${apiFor(this.prefix, key)}/items`, { body: fields });
    return created?.Id;
  }

  async update(key, spId, fields) {
    const result = await this.post(`${apiFor(this.prefix, key)}/items(${spId})/ValidateUpdateListItem()`, { body: { formValues: toFormValues(fields), bNewDocumentUpdate: false } });
    assertUpdate(result, `Updating ${key} item ${spId}`);
  }

  async recycle(key, spId) { await this.post(`${apiFor(this.prefix, key)}/items(${spId})/recycle()`, {}); }

  async saveProject(row) { return row.spId ? (await this.update('projects', row.spId, projectFields(row)), row) : { ...row, spId: await this.create('projects', projectFields(row)) }; }
  async saveTask(row) { return row.spId ? (await this.update('tasks', row.spId, taskFields(row)), row) : { ...row, spId: await this.create('tasks', taskFields(row)) }; }
  async saveUpdate(row) { return { ...row, spId: await this.create('updates', updateFields(row)) }; }
  async saveRisk(row) { return row.spId ? (await this.update('risks', row.spId, riskFields(row)), row) : { ...row, spId: await this.create('risks', riskFields(row)) }; }

}

export { titleFor };
