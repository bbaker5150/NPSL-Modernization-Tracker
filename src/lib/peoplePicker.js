export function parsePeopleResults(response) {
  const raw = response?.d?.ClientPeoplePickerSearchUser ?? response?.ClientPeoplePickerSearchUser ?? response?.value ?? response;
  const rows = typeof raw === 'string' ? JSON.parse(raw) : raw;
  if (!Array.isArray(rows)) throw new Error('SharePoint returned an unreadable people search result.');
  const seen = new Set();
  return rows.filter((row) => row.IsResolved === true && row.EntityType === 'User' && row.Key).map((row) => ({
    title: row.DisplayText || row.Description || row.Key,
    email: row.EntityData?.Email || '',
    loginName: row.Key,
    detail: [row.EntityData?.Department, row.EntityData?.JobTitle].filter(Boolean).join(' · '),
  })).filter((row) => { const key = row.loginName.toLowerCase(); if (seen.has(key)) return false; seen.add(key); return true; });
}

export function invitationUrl(config = {}, win = window) {
  const candidates = [config.appUrl];
  try { if (win.parent !== win) candidates.push(win.parent.location.href); } catch { /* cross-origin host */ }
  candidates.push(win.location.href);
  return candidates.find((value) => { try { return ['https:', 'http:'].includes(new URL(value).protocol); } catch { return false; } }) || '';
}

export function invitationMailto(person, role, url) {
  const parsed = new URL(url);
  if (!['https:', 'http:'].includes(parsed.protocol)) throw new Error('Enter a valid app link.');
  if (!person.email || /[\r\n]/.test(person.email)) throw new Error('This person has no email address. Copy the invitation link instead.');
  const subject = 'Invitation to the NPSL Modernization Tracker';
  const body = `Hello ${person.title},\n\nYou have been added to the NPSL Modernization Tracker as ${role}.\n\nOpen the tracker: ${parsed.href}\n\nSign in with your Flank Speed account. If SharePoint says access is denied, contact the site owner for access.\n`;
  return `mailto:${encodeURIComponent(person.email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
}
