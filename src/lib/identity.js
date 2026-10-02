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

export function isOwnedByUser(record, user) {
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

// Only legacy roles migrate. An explicitly saved Viewer remains read-only,
// even while a historical project assignment still references that person.
export function normalizeDirectory(users, projects) {
  return users.map((user) => ({ ...user, ...(!ROLES.includes(user.role) ? { roleMigrationPending: true } : {}), role: !ROLES.includes(user.role) && projects.some((project) => isOwnedByUser(project, user)) ? 'Project Engineer' : normalizeRole(user.role) }));
}
