import { describe, it, expect } from 'vitest';
import { invitationUrl, normalizeInvitationUrl } from './peoplePicker';

describe('embedded invitation URLs', () => {
  const webUrl = 'https://tenant.sharepoint-mil.us/sites/mod';
  it('targets the known ISEA METENG tracker instead of the host library view', () => {
    const site = 'https://flankspeed.sharepoint-mil.us/sites/ISEAMETENG';
    const win = { location: { href: `${site}/SitePages/Forms/ByAuthor.aspx` } };
    expect(invitationUrl({ webUrl: site }, win)).toBe(`${site}/SitePages/Modernization%20Tracker.aspx`);
    expect(invitationUrl({ webUrl: `${site}/`, appUrl: 'SitePages/Renamed.aspx' }, win)).toBe(`${site}/SitePages/Renamed.aspx`);
    expect(invitationUrl({ webUrl: site.toLowerCase() }, {})).toBe(`${site.toLowerCase()}/SitePages/Modernization%20Tracker.aspx`);
    expect(invitationUrl({ webUrl: `${site}-other` }, {})).toBe('');
    expect(invitationUrl({ webUrl: 'https://other.sharepoint-mil.us/sites/ISEAMETENG' }, {})).toBe('');
  });
  it('uses the referring app page when srcdoc cannot read its parent', () => {
    const win = { location: { href: 'about:srcdoc' }, document: { referrer: `${webUrl}/SitePages/Tracker.aspx` } };
    Object.defineProperty(win, 'parent', { get() { throw new Error('cross origin'); } });
    expect(invitationUrl({ webUrl }, win)).toBe(`${webUrl}/SitePages/Tracker.aspx`);
  });
  it('requires a direct page when host details are unavailable', () => {
    expect(invitationUrl({ webUrl }, { location: { href: 'blob:https://host/abc' } })).toBe('');
  });
  it('resolves explicit server-relative links and rejects empty or external links', () => {
    expect(invitationUrl({ webUrl, appUrl: '/sites/mod/app.aspx' }, {})).toBe(`${webUrl}/app.aspx`);
    expect(normalizeInvitationUrl('SitePages/Tracker.aspx', webUrl)).toBe(`${webUrl}/SitePages/Tracker.aspx`);
    for (const url of ['', 'https://elsewhere.example/app', 'javascript:alert(1)', 'https://tenant.sharepoint-mil.us/sites/other/app']) expect(() => normalizeInvitationUrl(url, webUrl)).toThrow('Enter the full published tracker page URL');
  });
});
