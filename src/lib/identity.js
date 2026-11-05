import { normalizeRole, ROLES } from '../data/workflow';

export function userIdentityKey(user) {
  if (!user) return '';
  const raw = user.loginName || user.email || (user.id ? `sharepoint-user:${user.id}` : '');
  return String(raw).trim().toLowerCase();
}

function identityAliases(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (!normalized) return [];
  const aliases = new Set([normalized]);
  const claimsValue = normalized.split('|').pop();
  if (claimsValue) aliases.add(claimsValue);
  return [...aliases];
}

function hasIdentityAlias(record, user) {
  const userAliases = new Set([
    ...identityAliases(userIdentityKey(user)),
    ...identityAliases(user?.loginName),
    ...identityAliases(user?.email),
  ]);
  const recordAliases = [
    ...identityAliases(record?.ownerKey),
    ...identityAliases(record?.ownerEmail),
  ];
  return recordAliases.some((alias) => userAliases.has(alias));
}

export function isOwnedByUser(record, user) {
  const ownerLogin = String(record?.ownerKey || '').trim().toLowerCase();
  const login = String(user?.loginName || '').trim().toLowerCase();
  // Distinct resolved accounts may share an email address. Use aliases only
  // when one side has no claims login, as on older imported records.
  if (ownerLogin.includes('|') && login.includes('|')) return ownerLogin === login;
  return hasIdentityAlias(record, user);
}

// A display name or a shared mailbox cannot establish which account should
// receive permissions. Accept older records keyed by the same login or email.
export function matchesAssignmentLogin(login, person) {
  const key = String(login || '').trim().toLowerCase();
  return key === userIdentityKey(person) || (!key.includes('|') && isOwnedByUser({ ownerKey: key }, person));
}
export function assignmentDirectoryUser(users, person) {
  const matches = users.filter(row => hasIdentityAlias({ ownerKey: row.loginName, ownerEmail: row.email }, person));
  const entry = matches[0];
  if (matches.length > 1 || (entry?.loginName && !matchesAssignmentLogin(entry.loginName, person))) {
    throw new Error('The selected account matches conflicting Tracker directory identities. Ask a manager to check the logins before assigning this project.');
  }
  return entry;
}

// Only legacy roles migrate. An explicitly saved Viewer remains read-only,
// even while a historical project assignment still references that person.
export function normalizeDirectory(users, projects) {
  return users.map((user) => ({ ...user, ...(!ROLES.includes(user.role) ? { roleMigrationPending: true } : {}), role: !ROLES.includes(user.role) && projects.some((project) => isOwnedByUser(project, user)) ? 'Project Engineer' : normalizeRole(user.role) }));
}
