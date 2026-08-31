import { seedData } from '../data/seed';
import { resolveWebUrl } from './spContext';
import { SharePointStore } from './spStore';

const STORAGE_KEY = 'modernization-project-tracker:v2';
const EMPTY_DATA = { projects: [], tasks: [], updates: [], risks: [] };
const clone = (value) => JSON.parse(JSON.stringify(value));
const uid = (prefix) => `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;

class LocalStore {
  constructor() {
    const stored = localStorage.getItem(STORAGE_KEY);
    this.data = stored ? JSON.parse(stored) : clone(EMPTY_DATA);
    this.data.updates ||= [];
    this.data.risks ||= [];
  }

  persist() { localStorage.setItem(STORAGE_KEY, JSON.stringify(this.data)); }
  async currentUser() { return { id: 1, title: 'Local Engineer', email: 'local.engineer@example.invalid', loginName: 'local' }; }
  async readiness() { return { ready: true, checks: [] }; }
  async provision() { return []; }
  async load() { return clone(this.data); }
  async saveProject(row) { return this.upsert('projects', row, 'project'); }
  async saveTask(row) { return this.upsert('tasks', row, 'task'); }
  async saveUpdate(row) { return this.upsert('updates', row, 'update'); }
  async saveRisk(row) { return this.upsert('risks', row, 'risk'); }
  async upsert(collection, row, prefix) {
    const record = { ...row, id: row.id || uid(prefix) };
    const index = this.data[collection].findIndex((item) => item.id === record.id);
    if (index >= 0) this.data[collection][index] = record; else this.data[collection].push(record);
    this.persist();
    return clone(record);
  }
  async recycle(collection, spId, id) {
    this.data[collection] = this.data[collection].filter((item) => item.id !== id && item.spId !== spId);
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
    return { mode: 'sharepoint', store: new SharePointStore({ webUrl, prefix: config.listPrefix || 'Modernization' }), config };
  }
  return { mode: 'local', store: new LocalStore(), config };
}

export function userIdentityKey(user) {
  if (!user) return '';
  const raw = user.loginName || user.email || (user.id ? `sharepoint-user:${user.id}` : '');
  return String(raw).trim().toLowerCase();
}

export function isOwnedByUser(record, user) {
  const identity = userIdentityKey(user);
  const recordKey = String(record?.ownerKey || '').trim().toLowerCase();
  if (identity && recordKey) return identity === recordKey;
  // Compatibility for records made before OwnerKey was introduced.
  const email = String(user?.email || '').trim().toLowerCase();
  return !!email && String(record?.ownerEmail || '').trim().toLowerCase() === email;
}

export function createStarterTasks(projectKey) {
  const templateProjectKey = seedData.projects[0]?.projectKey;
  return seedData.tasks
    .filter((task) => task.projectKey === templateProjectKey)
    .sort((a, b) => a.order - b.order)
    .map((task, index) => ({
      id: uid('task'),
      projectKey,
      wbs: task.wbs,
      title: task.title,
      phaseKey: task.phaseKey,
      order: index + 1,
      status: 'Not Started',
      startDate: '',
      finishDate: '',
      dueDate: '',
      ownerName: 'Unassigned',
      ownerEmail: '',
      ownerKey: '',
      notes: '',
      blockedReason: '',
      sourceStartLabel: '',
      dataIssue: '',
    }));
}

export { seedData };
