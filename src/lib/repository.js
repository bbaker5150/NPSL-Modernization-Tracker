import { seedData } from '../data/seed';
import { resolveWebUrl } from './spContext';
import { SharePointStore } from './spStore';

const STORAGE_KEY = 'npsl-modernization-tracker:v1';
const clone = (value) => JSON.parse(JSON.stringify(value));
const uid = (prefix) => `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;

class LocalStore {
  constructor() {
    const stored = localStorage.getItem(STORAGE_KEY);
    this.data = stored ? JSON.parse(stored) : clone(seedData);
    this.data.updates ||= [];
    this.data.risks ||= [];
  }

  persist() { localStorage.setItem(STORAGE_KEY, JSON.stringify(this.data)); }
  async currentUser() { return { id: 1, title: 'Demo Engineer', email: 'demo.engineer@navair.navy.mil', loginName: 'demo' }; }
  async readiness() { return { ready: true, checks: [] }; }
  async provision() { return []; }
  async load() { return clone(this.data); }
  async seed() { this.data = clone(seedData); this.persist(); return { count: this.data.projects.length + this.data.tasks.length }; }
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
  const inSharePoint = /^https?:/i.test(webUrl) && /sharepoint\.com/i.test(location.hostname);
  if (inSharePoint && config.forceLocal !== true) {
    return { mode: 'sharepoint', store: new SharePointStore({ webUrl, prefix: config.listPrefix || 'Modernization' }), config };
  }
  return { mode: 'demo', store: new LocalStore(), config };
}

export { seedData };
