import React, { useEffect, useRef, useState } from 'react';
import { invitationUrl, trackerPageUrl } from '../lib/peoplePicker';

export function PeopleInvite({ store, config, onSave, localPreview }) {
  const [query, setQuery] = useState('');
  const [results, setResults] = useState([]);
  const [selected, setSelected] = useState(null);
  const [role, setRole] = useState('User');
  const [busy, setBusy] = useState(false);
  const [searching, setSearching] = useState(false);
  const [error, setError] = useState('');
  const [added, setAdded] = useState(null);
  const link = invitationUrl({ ...config, webUrl: store.webUrl || config?.webUrl });
  const [access, setAccess] = useState(null);
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
      if (!localPreview && !link) throw new Error('The tracker invitation address is unavailable. Ask the site owner to configure the published tracker page address.');
      const validLink = localPreview ? '' : trackerPageUrl(link, store.webUrl || config?.webUrl);
      const person = added || await store.resolvePerson(selected.loginName);
      const saved = added || await onSave({ ...person, role });
      setAdded(saved);
      if (!localPreview) setAccess(await store.shareSiteAccess(saved, saved.role, validLink));
    } catch (err) { setError(err.message); }
    finally { setBusy(false); }
  }
  return <section className="panel people-invite" aria-label="Invite people">
    <h2>Invite people</h2><p>Find a person, choose their role, then invite them and grant site access.</p>
    {localPreview && <p>Preview mode searches saved tracker users. Organization search is available in SharePoint.</p>}
    <label className="field"><span>Find a person</span><input type="search" autoComplete="off" value={query} disabled={busy} placeholder="Name or email (at least 2 characters)" onChange={(event) => { setQuery(event.target.value); setSelected(null); setAdded(null); setAccess(null); setError(''); }} /></label>
    {searching && <p role="status">Searching people…</p>}
    {!selected && query.trim().length >= 2 && !searching && !error && !results.length && <p>No people found. Try a full name or email address.</p>}
    {!!results.length && <ul className="people-results" aria-label="Matching people">{results.map((person) => <li key={person.loginName}><button type="button" onClick={() => { setSelected(person); setRole('User'); setResults([]); setQuery(person.title); setAdded(null); setAccess(null); }}><strong>{person.title}</strong><span>{person.email || person.loginName}</span>{person.detail && <small>{person.detail}</small>}</button></li>)}</ul>}
    {selected && <div className="selected-person"><strong>{selected.title}</strong><span>{selected.email || selected.loginName}</span></div>}
    <div className="invite-actions"><label className="field"><span>Invitation role</span><select aria-label="Invitation role" value={role} disabled={busy || !!added} onChange={(event) => setRole(event.target.value)}><option>User</option><option>SME</option><option>Manager</option></select></label><button type="button" className="button primary" disabled={!selected || busy || !!access || (localPreview && !!added)} onClick={add}>{busy ? 'Saving and granting access…' : localPreview ? 'Add to tracker' : added ? 'Retry site invitation' : 'Invite and grant access'}</button></div>
    {added && <p className="invitation-result" role="status">{added.title} saved as {added.role}. {access ? 'Invitation sent to SharePoint for delivery.' : localPreview ? 'Preview only: no site access granted or email sent.' : 'Site invitation is not yet confirmed. Retry below if sharing failed.'}</p>}
    {error && <p className="inline-error" role="alert">{error}</p>}
  </section>;
}
