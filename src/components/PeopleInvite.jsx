import React, { useEffect, useRef, useState } from 'react';
import { invitationMailto, invitationUrl } from '../lib/peoplePicker';

export function PeopleInvite({ store, config, onSave, localPreview }) {
  const [query, setQuery] = useState('');
  const [results, setResults] = useState([]);
  const [selected, setSelected] = useState(null);
  const [role, setRole] = useState('User');
  const [busy, setBusy] = useState(false);
  const [searching, setSearching] = useState(false);
  const [error, setError] = useState('');
  const [added, setAdded] = useState(null);
  const [link, setLink] = useState(() => invitationUrl(config));
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
    setBusy(true); setError(''); setAdded(null);
    try {
      const person = await store.resolvePerson(selected.loginName);
      const saved = await onSave({ ...person, role });
      setAdded(saved); setCopied(false);
    } catch (err) { setError(err.message); }
    finally { setBusy(false); }
  }
  let mailto = '';
  try { if (added) mailto = invitationMailto(added, added.role, link); } catch { /* show copy option instead */ }
  return <section className="panel people-invite" aria-label="Invite people">
    <h2>Invite people</h2><p>Find a person, choose their role, then share an invitation.</p>
    {localPreview && <p>Preview mode searches saved tracker users. Organization search is available in SharePoint.</p>}
    <label className="field"><span>Find a person</span><input type="search" autoComplete="off" value={query} disabled={busy} placeholder="Name or email (at least 2 characters)" onChange={(event) => { setQuery(event.target.value); setSelected(null); setAdded(null); setError(''); }} /></label>
    {searching && <p role="status">Searching people…</p>}
    {!selected && query.trim().length >= 2 && !searching && !error && !results.length && <p>No people found. Try a full name or email address.</p>}
    {!!results.length && <ul className="people-results" aria-label="Matching people">{results.map((person) => <li key={person.loginName}><button type="button" onClick={() => { setSelected(person); setRole('User'); setResults([]); setQuery(person.title); setAdded(null); }}><strong>{person.title}</strong><span>{person.email || person.loginName}</span>{person.detail && <small>{person.detail}</small>}</button></li>)}</ul>}
    {selected && <div className="selected-person"><strong>{selected.title}</strong><span>{selected.email || selected.loginName}</span></div>}
    <div className="invite-actions"><label className="field"><span>Invitation role</span><select aria-label="Invitation role" value={role} disabled={busy || !!added} onChange={(event) => setRole(event.target.value)}><option>User</option><option>SME</option><option>Manager</option></select></label><button type="button" className="button primary" disabled={!selected || busy || !!added} onClick={add}>{busy ? 'Adding user…' : 'Add to tracker'}</button></div>
    <p className="field-hint">Tracker roles do not grant SharePoint site permissions. New users also need access to the app page and tracker lists.</p>
    {added && <div className="invitation-result"><p role="status">{added.title} added as {added.role}. Share the invitation below.</p><label className="field"><span>App invitation link</span><input type="url" value={link} onChange={(event) => { setLink(event.target.value); setCopied(false); }} /></label><div className="invite-actions">{mailto && <a className="button secondary" href={mailto}>Draft email invitation</a>}<button type="button" className="button secondary" onClick={async () => { try { if (!/^https?:\/\//i.test(link)) throw new Error('Enter a valid app link.'); await navigator.clipboard.writeText(link); setCopied(true); } catch (err) { setError('Could not copy. Select and copy the app invitation link above.'); } }}>Copy invitation link</button></div><p>{copied ? 'Invitation link copied.' : 'The email option opens a draft in your email app. No email has been sent.'}</p></div>}
    {error && <p className="inline-error" role="alert">{error}</p>}
  </section>;
}
