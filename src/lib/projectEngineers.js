// Project teams use resolved identities. Legacy Owner* values remain readable
// and mirror the first member for older exports and integrations.
function normalizeEngineers(entries) {
  if (!Array.isArray(entries)) throw new Error('The project engineer team is invalid. Ask a manager to repair its assignment.');
  const seen = new Set();
  return entries.flatMap(entry => {
    if (!entry || typeof entry !== 'object' || Array.isArray(entry) ||
        ['name', 'title', 'email', 'loginName'].some(key => entry[key] != null && typeof entry[key] !== 'string')) {
      throw new Error('The project engineer team contains an invalid account. Ask a manager to repair its assignment.');
    }
    const person = { name: (entry.name || entry.title || '').trim(), email: (entry.email || '').trim(), loginName: (entry.loginName || '').trim() };
    if (!person.name && !person.email && !person.loginName) throw new Error('The project engineer team contains an empty account.');
    const key = (person.loginName || person.email || `name:${person.name}`).toLowerCase();
    if (seen.has(key)) return [];
    seen.add(key);
    return [person];
  });
}

export function projectEngineers(project) {
  if (project?.engineers != null) return normalizeEngineers(project.engineers);
  const person = { name: project?.ownerName || '', email: project?.ownerEmail || '', loginName: project?.ownerKey || '' };
  if (!person.loginName && !person.email && (!person.name.trim() || person.name.trim().toLowerCase() === 'unassigned')) return [];
  return normalizeEngineers([person]);
}

export function engineerFields(entries) {
  const engineers = normalizeEngineers(entries);
  const first = engineers[0];
  return { engineers, ownerName: first?.name || first?.email || first?.loginName || 'Unassigned', ownerEmail: first?.email || '', ownerKey: first?.loginName || '' };
}

export function decodeProjectEngineers(value, legacyProject = {}) {
  if (value == null || value === '') return projectEngineers(legacyProject);
  try { return normalizeEngineers(JSON.parse(value)); }
  catch { throw new Error('The saved project engineer team is invalid. Ask a site owner to check EngineersJson before changing project access.'); }
}
