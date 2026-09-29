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

export function normalizeInvitationUrl(value, webUrl) {
  try {
    if (!String(value || '').trim()) throw new Error();
    const url = webUrl ? new URL(String(value).trim(), `${webUrl.replace(/\/$/, '')}/`) : new URL(String(value).trim());
    if (!['http:', 'https:'].includes(url.protocol) || url.username || url.password) throw new Error();
    if (webUrl) {
      const site = new URL(webUrl);
      const path = site.pathname.replace(/\/$/, '');
      if (url.origin !== site.origin || !(url.pathname === path || url.pathname.startsWith(`${path}/`))) throw new Error();
    }
    return url.href;
  } catch { throw new Error('Enter the full published tracker page URL in App invitation link. It must be on this SharePoint site.'); }
}

export function trackerPageUrl(value, webUrl) {
  const url = new URL(normalizeInvitationUrl(value, webUrl));
  if (!/\.(aspx|html?)$/i.test(url.pathname) || /\/(Forms|_layouts)\//i.test(url.pathname)) throw new Error('Use the direct published tracker page URL, not the site address or Pages library view.');
  return url.href;
}

export function invitationUrl(config = {}, win = window) {
  const candidates = [config.appUrl];
  // This deployment's host can expose a Pages library view instead of the app.
  // Keep explicit configuration authoritative; sharing still checks the file.
  try {
    const site = new URL(config.webUrl);
    if (site.origin === 'https://flankspeed.sharepoint-mil.us' && site.pathname.replace(/\/$/, '').toLowerCase() === '/sites/iseameteng') {
      candidates.push(`${site.origin}${site.pathname.replace(/\/$/, '')}/SitePages/Modernization%20Tracker.aspx`);
    }
  } catch { /* no recognized deployment */ }
  try { if (win.parent !== win) candidates.push(win.parent.location.href); } catch { /* cross-origin host */ }
  candidates.push(win.location?.href, win.document?.referrer);
  // Never substitute a site/library URL for the actual tracker page.
  for (const candidate of candidates) {
    if (!candidate || /^(about|blob|data):/i.test(candidate)) continue;
    try { return trackerPageUrl(candidate, config.webUrl); } catch { /* next host hint */ }
  }
  return '';
}

export function invitationMailto(person, role, url) {
  const parsed = new URL(url);
  if (!['https:', 'http:'].includes(parsed.protocol)) throw new Error('Enter a valid app link.');
  if (!person.email || /[\r\n]/.test(person.email)) throw new Error('This person has no email address. Copy the invitation link instead.');
  const subject = 'Invitation to the NPSL Modernization Tracker';
  const body = `Hello ${person.title},\n\nYou have been added to the NPSL Modernization Tracker as ${role}.\n\nOpen the tracker: ${parsed.href}\n\nSign in with your Flank Speed account. If SharePoint says access is denied, contact the site owner for access.\n`;
  return `mailto:${encodeURIComponent(person.email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
}
