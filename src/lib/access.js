import { validateReference } from './referenceDocuments';
import { validateAttachment } from './taskAttachments';
import { METCAL_ORGANIZATION, METCAL_TEAMS, normalizeOrganization, normalizeTaskStatus, normalizeRole, ROLES } from '../data/workflow';
import { isOwnedByUser, userIdentityKey } from './repository';

export const DEFAULT_TEST_MANAGER_PASSWORD = 'admin123';

export function isManager(user, users = []) {
  return users.some((entry) => entry.role === 'Manager' && isOwnedByUser({ ownerKey: entry.loginName, ownerEmail: entry.email }, user));
}
export function userRole(user, users = []) {
  const entry = users.find((row) => isOwnedByUser({ ownerKey: row.loginName, ownerEmail: row.email }, user));
  return normalizeRole(entry?.role);
}
export function canViewPortfolio() { return true; }
export function canManageProject(project, user, users = []) {
  return isManager(user, users) || (userRole(user, users) === 'Project Engineer' && !!project && isOwnedByUser(project, user));
}
export function visibleData(data) {
  // All Tracker roles share Viewer access to the portfolio returned by
  // SharePoint. Project assignment restricts writes, not shared reads.
  // My work applies its own ownership filter without removing portfolio data.
  return data;
}

const engineerProjectFields = ['health', 'currentStageKey', 'status'];
const engineerTaskFields = [
  'status', 'organization', 'subAssignment', 'ownerName', 'ownerEmail', 'ownerKey',
  'estimatedHours', 'startDate', 'deferredDate', 'deferredJustification',
  'notRequiredJustification', 'blockedReason', 'notes',
];

function applyEditableFields(saved, draft, fields) {
  const next = { ...saved };
  for (const field of fields) {
    if (Object.hasOwn(draft, field)) next[field] = draft[field];
  }
  return next;
}

export function validateTask(task) {
  if (task.subAssignment && (normalizeOrganization(task.organization) !== METCAL_ORGANIZATION || !METCAL_TEAMS.includes(task.subAssignment))) throw new Error('Select a valid METCAL responsibility.');
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
        // On scoped sites the live groups determine who may look up an
        // assignee. Looking up an email shouldn't reload the whole portfolio.
        if (raw.permissions) {
          const { role } = await raw.permissions.currentRole(await raw.currentUser());
          if (role !== 'Manager') throw new Error('Only managers can perform this action.');
        } else await requireManager();
        return raw[property](value);
      };
      if (property === 'registerCurrentUser' || property === 'activateTestingManager') return async (password) => {
        const { user, data } = await context();
        const key = userIdentityKey(user);
        if (!key) throw new Error('Your signed-in identity could not be resolved.');
        const existing = (data.users || []).find((entry) => isOwnedByUser({ ownerKey: entry.loginName, ownerEmail: entry.email }, user));
        if (raw.scopedAccess) {
          if (property === 'activateTestingManager') throw new Error('Testing manager access is disabled. Ask the Tracker Managers group owner to assign your role.');
          // Read-only users must not write to the directory during startup.
          return existing;
        }
        if (property === 'activateTestingManager') {
          const expected = config.testingManagerPassword ?? DEFAULT_TEST_MANAGER_PASSWORD;
          if (!expected || password !== expected) throw new Error('Testing password is incorrect or testing access is disabled.');
        }
        const row = { ...existing, id: existing?.id || `user-${encodeURIComponent(key)}`, title: user.title || user.email || key, loginName: key, email: user.email || '', role: property === 'activateTestingManager' ? 'Manager' : (existing ? normalizeRole(existing.role) : data.projects.some((project) => isOwnedByUser(project, user)) ? 'Project Engineer' : 'Viewer') };
        delete row.roleMigrationPending;
        if (row.role === 'Manager') {
          // Persist the one-time migration while an authorized manager is present.
          // Explicit Viewer assignments are never promoted on subsequent loads.
          for (const entry of data.users.filter((entry) => entry.roleMigrationPending && entry.id !== row.id)) {
            const migrated = { ...entry }; delete migrated.roleMigrationPending;
            await raw.saveUser(migrated);
          }
        }
        if (existing && !existing.roleMigrationPending && Object.keys(row).every((field) => row[field] === existing[field])) return existing;
        return raw.saveUser(row);
      };
      if (property === 'load') return async () => { const { data, user } = await context(); return visibleData(data, user); };
      if (property === 'saveProgressMode') return async (id, mode) => {
        const { user, data } = await context();
        const project = data.projects.find((row) => row.id === id);
        if (!project || !canManageProject(project, user, data.users)) throw new Error('Only managers or assigned project engineers can change progress calculation.');
        if (!['phases', 'tasks'].includes(mode)) throw new Error('Select a valid progress calculation.');
        const next = { ...project, progressMode: mode };
        return raw.saveProjectTracking ? raw.saveProjectTracking(next, { progressModeOnly: true }) : raw.saveProject(next);
      };
      if (property === 'saveTask') return async (row) => {
        const { user, data, manager } = await context();
        const existing = row.id && data.tasks.find((task) => task.id === row.id);
        const project = data.projects.find((project) => project.projectKey === (existing?.projectKey || row.projectKey));
        if (!project || !canManageProject(project, user, data.users)) throw new Error('Only managers or assigned project engineers can edit tasks.');
        if (!manager && !existing) throw new Error('Only managers can create tasks.');
        const next = manager ? { ...row, projectKey: project.projectKey, spId: existing?.spId }
          : applyEditableFields(existing, row, engineerTaskFields);
        if (normalizeOrganization(next.organization || existing?.organization || project.organization) !== METCAL_ORGANIZATION) next.subAssignment = '';
        next.organization = normalizeOrganization(next.organization || existing?.organization || project.organization);
        next.status = normalizeTaskStatus(next.status);
        if (!existing && !next.assignedDate) next.assignedDate = new Date().toISOString().slice(0, 10);
        validateTask(next);
        next.estimatedHours = next.estimatedHours === '' || next.estimatedHours == null ? null : Number(next.estimatedHours);
        next.finishDate = ['Complete', 'Not Required', 'Not Applicable'].includes(next.status) ? (existing?.finishDate || new Date().toISOString().slice(0, 10)) : '';
        return manager ? raw.saveTask(next) : raw.saveTask(next, { trackingOnly: true });
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
        if (!task || !visibleData(data, user).tasks.some((row) => row.id === task.id)) throw new Error('You cannot access documents for this task.');
        if (property === 'downloadTaskAttachment') return raw.downloadTaskAttachment(task, file);
        if (['deleteTaskAttachment', 'renameTaskAttachment'].includes(property)) {
          const project = data.projects.find((row) => row.projectKey === task.projectKey);
          if (!canManageProject(project, user, data.users)) throw new Error('Only managers or assigned project engineers can change documents.');
          return raw[property](task, file, nextName);
        }
        if (property === 'addTaskAttachment') {
          if (!canManageProject(data.projects.find((row) => row.projectKey === task.projectKey), user, data.users)) throw new Error('You cannot attach documents to this task.');
          validateAttachment(file);
          return raw.addTaskAttachment(task, file);
        }
        return raw.listTaskAttachments(task);
      };
      if (['editUpdate', 'deleteUpdate', 'editRisk', 'deleteRisk'].includes(property)) return async (row) => {
        const { user, data } = await context();
        const risk = property === 'editRisk' || property === 'deleteRisk';
        const existing = (risk ? data.risks : data.updates).find(item => item.id === row.id);
        const project = data.projects.find(item => item.projectKey === existing?.projectKey);
        if (!existing || !project || !canManageProject(project, user, data.users)) throw new Error('Only managers or assigned project engineers can change an existing project update.');
        if (!row.version || row.version === '*' || row.version !== existing.version) throw new Error('This update changed. Copy your draft and reload the tracker before trying again.');
        if (property === 'deleteUpdate' || property === 'deleteRisk') return raw[property](existing);
        if (risk) {
          if (!row.title?.trim()) throw new Error('Enter a risk title.');
          return raw.editRisk({ ...existing, title: row.title.trim(), mitigation: row.mitigation?.trim() || '',
            severity: row.severity?.trim() || '', status: row.status?.trim() || 'Open' });
        }
        if (!row.summary?.trim()) throw new Error('Enter an update.');
        const task = row.taskId ? data.tasks.find(item => item.id === row.taskId && item.projectKey === existing.projectKey) : null;
        // A removed task may keep its historical reference if it was not changed.
        if (row.taskId && !task && row.taskId !== existing.taskId) throw new Error('Select a task from this project.');
        return raw.editUpdate({ ...existing, summary: row.summary.trim(), taskId: row.taskId || '', taskTitle: task?.title || (row.taskId ? existing.taskTitle : ''),
          editedAt: new Date().toISOString(), editorName: user.title || user.email || user.loginName, editorKey: userIdentityKey(user) });
      };
      if (property === 'saveUpdate') return async (row) => {
        const { user, data } = await context();
        const existing = data.updates.find(item => item.id === row.id);
        const projectKey = existing?.projectKey || row.projectKey;
        const project = data.projects.find(item => item.projectKey === projectKey);
        if (!project || !canManageProject(project, user, data.users)) throw new Error('Only managers or assigned project engineers can update this project.');
        if (existing) throw new Error('Use Edit update to change an existing update.');
        if (!row.summary?.trim()) throw new Error('Enter an update.');
        if (row.type && !['Status', 'Risk', 'Decision', 'Handoff'].includes(row.type)) throw new Error('Select a valid update type.');
        const task = row.taskId ? data.tasks.find(item => item.id === row.taskId && item.projectKey === projectKey) : null;
        if (row.taskId && !task) throw new Error('Select a task from this project or leave the reference blank.');
        return raw.saveUpdate({ ...row, type: row.type || 'Status', projectKey, spId: existing?.spId, summary: row.summary.trim(), taskId: task?.id || '', taskTitle: task?.title || '',
          postedAt: existing?.postedAt || new Date().toISOString(), entryDate: existing?.entryDate || new Date().toISOString().slice(0, 10),
          authorName: existing?.authorName || user.title || user.email || user.loginName, authorEmail: existing?.authorEmail || user.email || '', authorKey: existing?.authorKey || userIdentityKey(user) });
      };
      if (['listUpdateAttachments', 'addUpdateAttachment', 'downloadUpdateAttachment', 'deleteUpdateAttachment'].includes(property)) return async (updateId, file, version) => {
        const { user, data } = await context();
        const update = visibleData(data, user).updates.find(row => row.id === updateId);
        if (!update) throw new Error('You cannot access documents for this update.');
        if (property === 'addUpdateAttachment' || property === 'deleteUpdateAttachment') {
          const project = data.projects.find(row => row.projectKey === update.projectKey);
          if (!project || !canManageProject(project, user, data.users)) throw new Error('Only managers or assigned project engineers can change update documents.');
          if (property === 'addUpdateAttachment') validateAttachment(file);
          else if (!version || version === '*' || version !== update.version) throw new Error('This update changed. Refresh history before removing its document.');
        }
        return raw[property](update, file);
      };
      if (property === 'saveTasks' || property === 'saveProjectTracking') return undefined;
      if (['saveProject', 'saveRisk', 'saveAcronym', 'recycle', 'saveUser'].includes(property)) return async (...args) => {
        const { data, user, manager } = await context();
        if (!manager && property === 'saveUser') throw new Error('Only managers can perform this action.');
        if (property === 'saveAcronym') {
          if (!userIdentityKey(user)) throw new Error('Sign in to add an acronym.');
          const [row] = args;
          const existing = data.acronyms.find(entry => entry.id === row.id);
          if (!manager && (existing || row.spId != null)) throw new Error('Only managers can change existing acronyms.');
          const acronym = row.acronym?.trim().toUpperCase();
          const term = row.term?.trim();
          if (!acronym || !term) throw new Error('Enter an acronym and its full term.');
          if (data.acronyms.some(entry => entry !== existing && entry.acronym.trim().toUpperCase() === acronym)) throw new Error(`${acronym} is already in the glossary.`);
          const record = manager ? row : { id: row.id };
          return raw.saveAcronym({ ...record, acronym, term, definition: row.definition?.trim() || '' });
        }
        if (property === 'saveProject') {
          const [draft] = args;
          if (manager) {
            args[0] = { ...draft, organization: normalizeOrganization(draft.organization) };
            return raw.saveProject(...args);
          }
          const current = draft?.id && data.projects.find((entry) => entry.id === draft.id);
          if (!current || !canManageProject(current, user, data.users)) throw new Error('Only managers or assigned project engineers can edit this project.');
          // Only a manager's explicit assignment may trigger a role change.
          const next = applyEditableFields(current, draft, engineerProjectFields);
          return raw.saveProjectTracking ? raw.saveProjectTracking(next) : raw.saveProject(next);
        }
        if (!manager && property === 'saveRisk') {
          const current = data.risks.find((entry) => entry.id === args[0]?.id);
          args[0] = { ...args[0], spId: current?.spId, projectKey: current?.projectKey || args[0]?.projectKey };
          const project = data.projects.find((entry) => entry.projectKey === args[0]?.projectKey);
          if (!canManageProject(project, user, data.users)) throw new Error('Only managers or assigned project engineers can update this project.');
        }
        if (!manager && property === 'recycle') {
          const [collection, spId, id] = args;
          if (collection === 'acronyms') throw new Error('Only managers can remove acronyms.');
          if (collection === 'projects' || collection === 'tasks') throw new Error('Only managers can delete projects and tasks.');
          const rows = data[collection] || [];
          const row = rows.find((entry) => entry.id === id || (spId != null && entry.spId === spId));
          const project = data.projects.find((entry) => entry.projectKey === row?.projectKey);
          if (!row || !canManageProject(project, user, data.users)) throw new Error('Only managers or assigned project engineers can delete project content.');
          args[1] = row.spId; args[2] = row.id;
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
          if (!row.title?.trim() || !userIdentityKey(row) || !ROLES.includes(row.role)) throw new Error('Name, login/email and role are required.');
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
