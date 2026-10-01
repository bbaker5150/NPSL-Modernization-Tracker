import { validateReference } from './referenceDocuments';
import { validateAttachment } from './taskAttachments';
import { normalizeOrganization, normalizeTaskStatus } from '../data/workflow';
import { isOwnedByUser, userIdentityKey } from './repository';

export const DEFAULT_TEST_MANAGER_PASSWORD = 'admin123';

export function isManager(user, users = []) {
  return users.some((entry) => entry.role === 'Manager' && isOwnedByUser({ ownerKey: entry.loginName, ownerEmail: entry.email }, user));
}
export function canViewPortfolio(user, users = []) { return isManager(user, users); }
export function canManageProject(project, user, users = []) { return isManager(user, users) || (!!project?.ownerCanEdit && isOwnedByUser(project, user)); }
export function canUpdateTask(task, user, projects) {
  return isOwnedByUser(task, user) || projects.some((project) => project.projectKey === task.projectKey && isOwnedByUser(project, user));
}
export function visibleData(data, user) {
  if (canViewPortfolio(user, data.users)) return data;
  const tasks = data.tasks.filter((task) => canUpdateTask(task, user, data.projects));
  const keys = new Set(tasks.map((task) => task.projectKey));
  const projects = data.projects.filter((project) => isOwnedByUser(project, user) || keys.has(project.projectKey));
  // Related project history is restricted to projects owned by the user.
  const ownedKeys = new Set(projects.filter((project) => isOwnedByUser(project, user)).map((project) => project.projectKey));
  return { ...data, projects, tasks, updates: data.updates.filter((row) => ownedKeys.has(row.projectKey)), risks: data.risks.filter((row) => ownedKeys.has(row.projectKey)), users: data.users || [] };
}
export function validateTask(task) {
  if (!task.title?.trim()) throw new Error('Enter a task name.');
  if (!['Not Started', 'In Progress', 'Blocked', 'Complete', 'Not Required', 'Not Applicable'].includes(task.status)) throw new Error('Select a valid task status.');
  if (task.estimatedHours !== '' && task.estimatedHours != null && (!Number.isFinite(Number(task.estimatedHours)) || Number(task.estimatedHours) < 0)) throw new Error('Estimated hours must be a nonnegative number or left blank.');
  for (const date of [task.dueDate, task.deferredDate, task.assignedDate]) {
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
      if (property === 'saveReferenceEntry') return async (row, file) => {
        await requireManager();
        const entries = await raw.listReferenceEntries();
        return raw.saveReferenceEntry(validateReference(entries, row, file), file);
      };
      if (property === 'deleteReferenceEntry') return async (id) => {
        await requireManager();
        const entries = await raw.listReferenceEntries();
        const row = entries.find((entry) => entry.id === id);
        if (!row) throw new Error('This reference item no longer exists.');
        if (row.kind === 'folder' && entries.some((entry) => entry.parentId === row.id)) throw new Error('Move or delete the contents before deleting this folder.');
        return raw.deleteReferenceEntry(row);
      };
      if (property === 'downloadReferenceEntry') return async (id) => {
        const row = (await raw.listReferenceEntries()).find((entry) => entry.id === id && entry.kind === 'file');
        if (!row) throw new Error('This reference document no longer exists.');
        return raw.downloadReferenceEntry(row);
      };
      if (property === 'shareSiteAccess') return async (person, role, appUrl) => {
        const { data } = await requireManager();
        const saved = data.users.find((entry) => entry.loginName === person.loginName && entry.role === role);
        if (!saved) throw new Error('Save the person and role in the tracker before granting site access.');
        return raw.shareSiteAccess(saved, role, appUrl);
      };
      if (['searchPeople', 'resolvePerson'].includes(property)) return async (value) => {
        await requireManager();
        return raw[property](value);
      };
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
        const { user, data } = await context();
        const project = data.projects.find((row) => row.id === id);
        if (!project || !canManageProject(project, user, data.users)) throw new Error('Only managers or authorized project owners can change progress calculation.');
        if (!['phases', 'tasks'].includes(mode)) throw new Error('Select a valid progress calculation.');
        return raw.saveProject({ ...project, progressMode: mode });
      };
      if (property === 'saveTask') return async (row) => {
        const { user, data, manager } = await context();
        const existing = data.tasks.find((task) => task.id === row.id);
        const project = data.projects.find((project) => project.projectKey === (existing?.projectKey || row.projectKey));
        let next = { ...row };
        if (!manager && !canManageProject(project, user, data.users)) {
          if (!existing) {
            throw new Error('Only managers or authorized project owners can add tasks.');
          } else {
            if (!canUpdateTask(existing, user, data.projects)) throw new Error('You can only update your assigned tasks.');
            next = { ...existing, status: row.status, deferredDate: row.deferredDate || '', deferredJustification: row.deferredJustification || '', notRequiredJustification: row.notRequiredJustification || '' };
          }
        }
        next.status = normalizeTaskStatus(next.status);
        if (!existing && !next.assignedDate) next.assignedDate = new Date().toISOString().slice(0, 10);
        validateTask(next);
        next.estimatedHours = next.estimatedHours === '' || next.estimatedHours == null ? null : Number(next.estimatedHours);
        next.finishDate = ['Complete', 'Not Required', 'Not Applicable'].includes(next.status) ? (existing?.finishDate || new Date().toISOString().slice(0, 10)) : '';
        return raw.saveTask(next);
      };
      if (property === 'listProjectAttachments') return async (projectKey) => {
        const { user, data } = await context();
        if (!visibleData(data, user).projects.some((project) => project.projectKey === projectKey)) throw new Error('You cannot access documents for this project.');
        const tasks = data.tasks.filter((task) => task.projectKey === projectKey);
        return (await Promise.all(tasks.map(async (task) => (await raw.listTaskAttachments(task)).map((file) => ({ ...file, taskId: task.id, taskTitle: task.title }))))).flat();
      };
      if (['listTaskAttachments', 'addTaskAttachment', 'deleteTaskAttachment', 'renameTaskAttachment', 'downloadTaskAttachment'].includes(property)) return async (taskId, file, nextName) => {
        const { user, data, manager } = await context();
        const task = data.tasks.find((row) => row.id === taskId);
        if (!task || (!canViewPortfolio(user, data.users) && !canUpdateTask(task, user, data.projects))) throw new Error('You cannot access documents for this task.');
        if (property === 'downloadTaskAttachment') return raw.downloadTaskAttachment(task, file);
        if (['deleteTaskAttachment', 'renameTaskAttachment'].includes(property)) {
          const project = data.projects.find((row) => row.projectKey === task.projectKey);
          if (!canManageProject(project, user, data.users)) throw new Error('Only managers or authorized project owners can change documents.');
          return raw[property](task, file, nextName);
        }
        if (property === 'addTaskAttachment') {
          if (!manager && !canUpdateTask(task, user, data.projects)) throw new Error('You cannot attach documents to this task.');
          validateAttachment(file);
          return raw.addTaskAttachment(task, file);
        }
        return raw.listTaskAttachments(task);
      };
      if (property === 'saveTasks') return undefined;
      if (['saveProject', 'saveUpdate', 'saveRisk', 'saveAcronym', 'recycle', 'saveUser'].includes(property)) return async (...args) => {
        const { data, user, manager } = await context();
        if (!manager && ['saveAcronym', 'saveUser'].includes(property)) throw new Error('Only managers can perform this action.');
        if (!manager && property === 'saveProject') {
          const current = data.projects.find((entry) => entry.id === args[0]?.id);
          if (!current || !canManageProject(current, user, data.users)) throw new Error('Only managers or authorized project owners can edit this project.');
          args[0] = { ...args[0], ownerName: current.ownerName, ownerEmail: current.ownerEmail, ownerKey: current.ownerKey, ownerCanEdit: current.ownerCanEdit, organization: current.organization };
        }
        if (!manager && ['saveUpdate', 'saveRisk'].includes(property)) {
          const project = data.projects.find((entry) => entry.projectKey === args[0]?.projectKey);
          if (!canManageProject(project, user, data.users)) throw new Error('Only managers or authorized project owners can update this project.');
        }
        if (!manager && property === 'recycle') {
          const [collection, spId, id] = args;
          const rows = data[collection] || [];
          const row = rows.find((entry) => entry.id === id || (spId != null && entry.spId === spId));
          const projectKey = collection === 'projects' ? row?.projectKey : row?.projectKey;
          const project = data.projects.find((entry) => entry.projectKey === projectKey);
          if (!row || !canManageProject(project, user, data.users)) throw new Error('Only managers or authorized project owners can delete project content.');
        }
        if (property === 'recycle' && args[0] === 'users') {
          if (!manager) throw new Error('Only managers can perform this action.');
          const row = data.users.find((entry) => args[2] ? entry.id === args[2] : entry.spId != null && entry.spId === args[1]);
          if (!row) throw new Error('This user is no longer in the directory.');
          if (isOwnedByUser({ ownerKey: row.loginName, ownerEmail: row.email }, user)) throw new Error('You cannot delete your own account. Ask another manager.');
          // Resolve the persisted record, never trust an edited identity or supplied item ID.
          return raw.recycle('users', row.spId, row.id);
        }
        if (property === 'saveUser') {
          const row = args[0];
          if (!row.title?.trim() || !userIdentityKey(row) || !['Manager', 'User'].includes(row.role)) throw new Error('Name, login/email and role are required.');
          if (data.users?.some((entry) => entry.id !== row.id && isOwnedByUser({ ownerKey: entry.loginName, ownerEmail: entry.email }, row))) throw new Error('This user is already in the directory.');
          if (row.role !== 'Manager' && isOwnedByUser({ ownerKey: row.loginName, ownerEmail: row.email }, user)) throw new Error('Ask another manager to change your role.');
        }
        if (property === 'saveProject') args[0] = { ...args[0], organization: normalizeOrganization(args[0].organization), ownerCanEdit: !!args[0].ownerCanEdit };
        return raw[property](...args);
      };
      const value = target[property];
      return typeof value === 'function' ? value.bind(target) : value;
    },
  });
}
