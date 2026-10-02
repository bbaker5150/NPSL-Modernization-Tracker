import { normalizeDirectory } from './identity';
export { isOwnedByUser, userIdentityKey } from './identity';
import { localAttachments, validateAttachment, validateAttachmentName } from './taskAttachments';
import { workflowData, normalizePhaseKey, normalizeOrganization, normalizeTaskStatus, normalizeTaskOrganizations } from '../data/workflow';
import { DEFAULT_ACRONYM_VERSION, defaultAcronyms } from '../data/defaultAcronyms';
import { resolveWebUrl } from './spContext';
import { SharePointStore } from './spStore';

const STORAGE_KEY = 'modernization-project-tracker:v2';
const ACRONYM_DEFAULTS_KEY = `${STORAGE_KEY}:acronyms:${DEFAULT_ACRONYM_VERSION}`;
const EMPTY_DATA = { projects: [], tasks: [], updates: [], risks: [], acronyms: [] };
const clone = (value) => JSON.parse(JSON.stringify(value));
const uid = (prefix) => `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;

class LocalStore {
  constructor() {
    const stored = localStorage.getItem(STORAGE_KEY);
    this.data = stored ? JSON.parse(stored) : clone(EMPTY_DATA);
    this.data.users ||= [];
    this.data.references ||= [];
    this.data.updates ||= [];
    this.data.risks ||= [];
    this.data.acronyms ||= [];
    if (!localStorage.getItem(ACRONYM_DEFAULTS_KEY)) {
      const existing = new Set(this.data.acronyms.map((entry) => String(entry.acronym || '').toUpperCase()));
      this.data.acronyms.push(...defaultAcronyms.filter((entry) => !existing.has(entry.acronym.toUpperCase())));
      localStorage.setItem(ACRONYM_DEFAULTS_KEY, '1');
      this.persist();
    }
  }

  persist() { localStorage.setItem(STORAGE_KEY, JSON.stringify(this.data)); }
  async currentUser() { return { id: 1, title: 'Local Engineer', email: 'local.engineer@example.invalid', loginName: 'local', isSiteAdmin: true }; }
  async readiness() { return { ready: true, checks: [] }; }
  async provision() { return []; }
  async load() { const data = clone(this.data); data.projects = data.projects.map((row) => ({ ...row, currentStageKey: normalizePhaseKey(row.currentStageKey), organization: normalizeOrganization(row.organization), ownerCanEdit: !!row.ownerCanEdit })); data.tasks = data.tasks.map((row) => ({ ...row, phaseKey: normalizePhaseKey(row.phaseKey), status: normalizeTaskStatus(row.status) })); data.users = normalizeDirectory(data.users, data.projects); data.tasks = normalizeTaskOrganizations(data.tasks, data.projects); return data; }
  async listTaskAttachments(task) { return localAttachments(task.id); }
  async downloadTaskAttachment(task, name) { const file = (await this.listTaskAttachments(task)).find((entry) => entry.name === name); if (!file) throw new Error('Document no longer exists.'); return file.blob; }
  async deleteTaskAttachment(task, name) { await localAttachments(task.id, undefined, name); }
  async addTaskAttachment(task, file) { validateAttachment(file); await localAttachments(task.id, file); return this.listTaskAttachments(task); }
  async renameTaskAttachment(task, name, nextName) {
    const cleanName = validateAttachmentName(nextName);
    const file = (await this.listTaskAttachments(task)).find((entry) => entry.name === name);
    if (!file) throw new Error('Document no longer exists.');
    if (file.name.toLowerCase() === cleanName.toLowerCase()) return this.listTaskAttachments(task);
    const renamed = new File([file.blob], cleanName, { type: file.blob.type, lastModified: Date.now() });
    validateAttachment(renamed);
    await localAttachments(task.id, renamed);
    await localAttachments(task.id, undefined, name);
    return this.listTaskAttachments(task);
  }
  async listReferenceEntries() { return clone(this.data.references); }
  async saveReferenceEntry(row, file) {
    const next = { ...row, id: row.id || uid('reference') };
    if (file) await localAttachments(next.id, file);
    return this.upsert('references', next, 'reference');
  }
  async deleteReferenceEntry(row) {
    this.data.references = this.data.references.filter((entry) => entry.id !== row.id);
    this.persist();
  }
  async downloadReferenceEntry(row) {
    const file = (await localAttachments(row.id)).find((entry) => entry.name === row.fileName);
    if (!file) throw new Error('This reference document is unavailable.');
    return file.blob;
  }
  async shareSiteAccess() { throw new Error('Site sharing is available only when the tracker is hosted in SharePoint.'); }
  async searchPeople(query) { const text = query.trim().toLowerCase(); return this.data.users.filter((row) => `${row.title} ${row.email}`.toLowerCase().includes(text)).map(({ title, email, loginName }) => ({ title, email, loginName: loginName || email })); }
  async resolvePerson(loginName) { const person = this.data.users.find((row) => (row.loginName || row.email) === loginName); if (!person) throw new Error('Preview mode can only select saved tracker users.'); return { title: person.title, email: person.email || '', loginName }; }
  async saveUser(row) { return this.upsert('users', row, 'user'); }
  async saveProject(row) { return this.upsert('projects', row, 'project'); }
  async saveTask(row) { return this.upsert('tasks', row, 'task'); }
  async saveUpdate(row) { return this.upsert('updates', row, 'update'); }
  async saveRisk(row) { return this.upsert('risks', row, 'risk'); }
  async saveAcronym(row) { return this.upsert('acronyms', row, 'acronym'); }
  async upsert(collection, row, prefix) {
    const record = { ...row, id: row.id || uid(prefix) };
    const index = this.data[collection].findIndex((item) => item.id === record.id);
    if (index >= 0) this.data[collection][index] = record; else this.data[collection].push(record);
    this.persist();
    return clone(record);
  }
  async recycle(collection, spId, id) {
    this.data[collection] = this.data[collection].filter((item) => item.id !== id && (spId == null || item.spId !== spId));
    this.persist();
  }
}

export function createRepository() {
  const config = window.MOD_TRACKER_CONFIG || {};
  const webUrl = config.webUrl || resolveWebUrl();
  // Forge loads this file in an iframe srcdoc. In that frame location.hostname
  // is empty even though resolveWebUrl() correctly finds the SharePoint parent,
  // so the resolved web URL — not the child frame location — is authoritative.
  const inSharePoint = /^https:\/\/[^/]*\.(?:sharepoint\.com|sharepoint\.us|sharepoint-mil\.us|sharepoint\.de|sharepoint\.cn)(?::\d+)?(?:\/|$)/i.test(webUrl);
  if ((inSharePoint || config.forceSharePoint === true) && config.forceLocal !== true) {
    return { mode: 'sharepoint', store: new SharePointStore({ webUrl, prefix: config.listPrefix || 'Modernization', hideLists: config.hideLists !== false }), config };
  }
  return { mode: 'local', store: new LocalStore(), config };
}

export function createStarterTasks(projectKey, owner = {}) {
  return [...workflowData.taskTemplates]
    .sort((a, b) => a.order - b.order)
    .map((task, index) => ({
      id: uid('task'),
      projectKey,
      title: task.title,
      phaseKey: task.phaseKey,
      order: index + 1,
      status: 'Not Started',
      startDate: '',
      finishDate: '',
      dueDate: '',
      ownerName: owner.ownerName?.trim() || 'Unassigned',
      ownerEmail: owner.ownerEmail || '',
      ownerKey: owner.ownerKey || '',
      notes: '',
      blockedReason: '',
      sourceStartLabel: '',
      dataIssue: '',
    }));
}

export { workflowData };
