import { workflowData, normalizePhaseKey } from '../data/workflow';
import { isOwnedByUser, userIdentityKey } from './repository';

export const DEFAULT_TEST_MANAGER_PASSWORD = 'admin123';

export function isManager(user, users = []) {
  return users.some((entry) => entry.role === 'Manager' && isOwnedByUser({ ownerKey: entry.loginName, ownerEmail: entry.email }, user));
}
export function canUpdateTask(task, user, projects) {
  return isOwnedByUser(task, user) || projects.some((project) => project.projectKey === task.projectKey && isOwnedByUser(project, user));
}
export function visibleData(data, user) {
  if (isManager(user, data.users)) return data;
  const tasks = data.tasks.filter((task) => canUpdateTask(task, user, data.projects));
  const keys = new Set(tasks.map((task) => task.projectKey));
  const projects = data.projects.filter((project) => isOwnedByUser(project, user) || keys.has(project.projectKey));
  // Related project history is restricted to projects owned by the user.
  const ownedKeys = new Set(projects.filter((project) => isOwnedByUser(project, user)).map((project) => project.projectKey));
  return { ...data, projects, tasks, updates: data.updates.filter((row) => ownedKeys.has(row.projectKey)), risks: data.risks.filter((row) => ownedKeys.has(row.projectKey)), users: data.users || [] };
}
export function validateTask(task) {
  if (!task.title?.trim()) throw new Error('Enter a task name.');
  if (!['Not Started', 'In Progress', 'In Progress – At Program Office', 'Blocked', 'Complete', 'Not Required', 'Not Applicable'].includes(task.status)) throw new Error('Select a valid task status.');
  for (const date of [task.dueDate, task.deferredDate]) {
    if (date && (!/^\d{4}-\d{2}-\d{2}$/.test(date) || new Date(`${date}T12:00:00Z`).toISOString().slice(0, 10) !== date)) throw new Error('Enter a valid date.');
  }
  if (task.deferredDate && (!task.dueDate || task.deferredDate <= task.dueDate)) throw new Error('The deferred date must be after the original due date.');
  if (task.deferredDate && !task.deferredJustification?.trim()) throw new Error('A deferred date requires justification.');
  if (['Not Required', 'Not Applicable'].includes(task.status) && !task.notRequiredJustification?.trim()) throw new Error('Marking a task not required requires justification.');
}

// Application authorization applies to every repository call, including exports.
// SharePoint ACLs remain the server-side security boundary (see deployment guide).
export function authorizedStore(raw, config = {}) {
  const context = async () => {
    const [user, data] = await Promise.all([raw.currentUser(), raw.load()]);
    return { user, data, manager: isManager(user, data.users) };
  };
  const requireManager = async () => {
    const ctx = await context();
    if (!ctx.manager) throw new Error('Only managers can perform this action.');
    return ctx;
  };
  return new Proxy(raw, {
    get(target, property) {
      if (property === 'registerCurrentUser' || property === 'activateTestingManager') return async (password) => {
        const { user, data } = await context();
        const key = userIdentityKey(user);
        if (!key) throw new Error('Your signed-in identity could not be resolved.');
        const existing = (data.users || []).find((entry) => isOwnedByUser({ ownerKey: entry.loginName, ownerEmail: entry.email }, user));
        if (property === 'activateTestingManager') {
          const expected = config.testingManagerPassword ?? DEFAULT_TEST_MANAGER_PASSWORD;
          if (!expected || password !== expected) throw new Error('Testing password is incorrect or testing access is disabled.');
        }
        const row = { ...existing, id: existing?.id || `user-${encodeURIComponent(key)}`, title: user.title || user.email || key, loginName: key, email: user.email || '', role: property === 'activateTestingManager' ? 'Manager' : existing?.role || 'User' };
        if (existing && Object.keys(row).every((field) => row[field] === existing[field])) return existing;
        return raw.saveUser(row);
      };
      if (property === 'load') return async () => { const { data, user } = await context(); return visibleData(data, user); };
      if (property === 'saveProgressMode') return async (id, mode) => {
        const { user, data, manager } = await context();
        const project = data.projects.find((row) => row.id === id);
        if (!project || (!manager && !isOwnedByUser(project, user))) throw new Error('Only the project owner or a manager can change progress calculation.');
        if (!['phases', 'tasks'].includes(mode)) throw new Error('Select a valid progress calculation.');
        return raw.saveProject({ ...project, progressMode: mode });
      };
      if (property === 'saveTask') return async (row) => {
        const { user, data, manager } = await context();
        const existing = data.tasks.find((task) => task.id === row.id);
        let next = { ...row };
        if (!manager) {
          if (!existing) {
            const project = data.projects.find((project) => project.projectKey === row.projectKey && isOwnedByUser(project, user));
            if (!project) throw new Error('You can only add tasks to your own project.');
            const phaseKey = normalizePhaseKey(row.phaseKey);
            if (!workflowData.phases.some((phase) => phase.key === phaseKey)) throw new Error('Select a valid phase.');
            next = { id: row.id, projectKey: project.projectKey, title: row.title, phaseKey, order: Math.max(0, ...data.tasks.filter((task) => task.projectKey === project.projectKey).map((task) => task.order || 0)) + 1, status: 'Not Started', ownerName: user.title || user.email, ownerKey: userIdentityKey(user), ownerEmail: user.email || '', dueDate: '', deferredDate: '', deferredJustification: '', notRequiredJustification: '' };
          } else {
            if (!canUpdateTask(existing, user, data.projects)) throw new Error('You can only update your assigned tasks.');
            next = { ...existing, status: row.status, deferredDate: row.deferredDate || '', deferredJustification: row.deferredJustification || '', notRequiredJustification: row.notRequiredJustification || '' };
          }
        }
        validateTask(next);
        next.finishDate = ['Complete', 'Not Required', 'Not Applicable'].includes(next.status) ? (existing?.finishDate || new Date().toISOString().slice(0, 10)) : '';
        return raw.saveTask(next);
      };
      if (property === 'saveTasks') return undefined;
      if (['saveProject', 'saveUpdate', 'saveRisk', 'saveAcronym', 'recycle', 'saveUser'].includes(property)) return async (...args) => {
        const { data, user } = await requireManager();
        if (property === 'saveUser') {
          const row = args[0];
          if (!row.title?.trim() || !userIdentityKey(row) || !['Manager', 'User'].includes(row.role)) throw new Error('Name, login/email and role are required.');
          if (data.users?.some((entry) => entry.id !== row.id && isOwnedByUser({ ownerKey: entry.loginName, ownerEmail: entry.email }, row))) throw new Error('This user is already in the directory.');
          if (row.role !== 'Manager' && isOwnedByUser({ ownerKey: row.loginName, ownerEmail: row.email }, user)) throw new Error('Ask another manager to change your role.');
        }
        return raw[property](...args);
      };
      const value = target[property];
      return typeof value === 'function' ? value.bind(target) : value;
    },
  });
}
