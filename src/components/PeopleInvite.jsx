import React, { useEffect, useRef, useState } from 'react';
import { invitationMailto, invitationUrl, normalizeInvitationUrl } from '../lib/peoplePicker';

export function PeopleInvite({ store, config, onSave, localPreview }) {
  const [query, setQuery] = useState('');
  const [results, setResults] = useState([]);
  const [selected, setSelected] = useState(null);
  const [role, setRole] = useState('User');
  const [busy, setBusy] = useState(false);
  const [searching, setSearching] = useState(false);
  const [error, setError] = useState('');
  const [added, setAdded] = useState(null);
  const [link, setLink] = useState(() => invitationUrl({ ...config, webUrl: store.webUrl || config?.webUrl }));
  const [access, setAccess] = useState(null);
  const [copied, setCopied] = useState(false);
  const generation = useRef(0);
  useEffect(() => {
    const current = ++generation.current;
    setResults([]);
    if (selected || query.trim().length < 2) { setSearching(false); return; }
    setSearching(true);
    const timer = setTimeout(async () => {
      try {
        const people = await store.searchPeople(query.trim());
        if (generation.current === current) { setResults(people); setError(''); }
      } catch (err) { if (generation.current === current) setError(err.message); }
      finally { if (generation.current === current) setSearching(false); }
    }, 350);
    return () => { clearTimeout(timer); generation.current += 1; };
  }, [query, selected, store]);
  async function add() {
    if (!selected || busy) return;
    setBusy(true); setError('');
    try {
      const validLink = normalizeInvitationUrl(link, store.webUrl || config?.webUrl);
      setLink(validLink);
      const person = added || await store.resolvePerson(selected.loginName);
      const saved = added || await onSave({ ...person, role });
      setAdded(saved); setCopied(false);
      if (!localPreview) setAccess(await store.shareSiteAccess(saved, saved.role, validLink));
    } catch (err) { setError(err.message); }
    finally { setBusy(false); }
  }
  let mailto = '';
  try { if (added) mailto = invitationMailto(added, added.role, link); } catch { /* show copy option instead */ }
  return <section className="panel people-invite" aria-label="Invite people">
    <h2>Invite people</h2><p>Find a person, choose their role, then invite them and grant site access.</p>
    {localPreview && <p>Preview mode searches saved tracker users. Organization search is available in SharePoint.</p>}
    <label className="field"><span>Find a person</span><input type="search" autoComplete="off" value={query} disabled={busy} placeholder="Name or email (at least 2 characters)" onChange={(event) => { setQuery(event.target.value); setSelected(null); setAdded(null); setAccess(null); setError(''); }} /></label>
    {searching && <p role="status">Searching people…</p>}
    {!selected && query.trim().length >= 2 && !searching && !error && !results.length && <p>No people found. Try a full name or email address.</p>}
    {!!results.length && <ul className="people-results" aria-label="Matching people">{results.map((person) => <li key={person.loginName}><button type="button" onClick={() => { setSelected(person); setRole('User'); setResults([]); setQuery(person.title); setAdded(null); setAccess(null); }}><strong>{person.title}</strong><span>{person.email || person.loginName}</span>{person.detail && <small>{person.detail}</small>}</button></li>)}</ul>}
    {selected && <div className="selected-person"><strong>{selected.title}</strong><span>{selected.email || selected.loginName}</span></div>}
    <div className="invite-actions"><label className="field"><span>Invitation role</span><select aria-label="Invitation role" value={role} disabled={busy || !!added} onChange={(event) => setRole(event.target.value)}><option>User</option><option>SME</option><option>Manager</option></select></label><button type="button" className="button primary" disabled={!selected || busy || !!access || (localPreview && !!added)} onClick={add}>{busy ? 'Saving and granting access…' : localPreview ? 'Add to tracker' : added ? 'Retry site invitation' : 'Invite and grant access'}</button></div>
    <label className="field"><span>App invitation link</span><input type="url" disabled={busy || !!access} value={link} onChange={(event) => { setLink(event.target.value); setCopied(false); }} /></label><p className="field-hint">Check that the link opens the tracker page. If only the site address was detected, replace it with the published app page URL.</p><p className="field-hint">Grants {role === 'SME' ? 'Read' : 'Contribute'} access to this site and content that inherits its permissions. Requires permission to share the site. Separately secured pages/lists need site-owner access; existing broader permissions are not removed.</p>
    {added && <div className="invitation-result"><p role="status">{added.title} saved as {added.role}. {access ? `${access.access} site access verified. SharePoint invitation email requested.` : localPreview ? 'Preview only: no site access granted or email sent.' : 'Site invitation is not yet confirmed. Retry below if sharing failed.'}</p><div className="invite-actions">{mailto && <a className="button secondary" href={mailto}>Draft email invitation</a>}<button type="button" className="button secondary" onClick={async () => { try { if (!/^https?:\/\//i.test(link)) throw new Error('Enter a valid app link.'); await navigator.clipboard.writeText(link); setCopied(true); } catch (err) { setError('Could not copy. Select and copy the app invitation link above.'); } }}>Copy invitation link</button></div><p>{copied ? 'Invitation link copied.' : access ? 'SharePoint accepted the invitation request; email delivery is not confirmed. The draft option is available as a fallback.' : 'The email option opens a draft in your email app. No email has been sent by this draft option.'}</p></div>}
    {error && <p className="inline-error" role="alert">{error}</p>}
  </section>;
}
