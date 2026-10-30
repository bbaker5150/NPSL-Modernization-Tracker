import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Icon } from './Icon';
import { createFeedback } from '../lib/feedback';

export function Feedback({ store }) {
  const api = useMemo(() => createFeedback(store), [store]);
  const dialog = useRef(null), trigger = useRef(null), active = useRef(false);
  const [open, setOpen] = useState(false), [owner, setOwner] = useState(false), [tab, setTab] = useState('submit');
  const [kind, setKind] = useState('Bug'), [title, setTitle] = useState(''), [details, setDetails] = useState('');
  const [busy, setBusy] = useState(false), [error, setError] = useState(''), [sent, setSent] = useState(false), [items, setItems] = useState(null);
  useEffect(() => { let mounted = true; api.isOwner().then(value => { if (mounted) setOwner(value); }).catch(() => {}); return () => { mounted = false; }; }, [api]);
  useEffect(() => {
    if (!open) return;
    dialog.current.showModal();
    return () => { dialog.current?.close(); trigger.current?.focus(); };
  }, [open]);
  async function submit(event) {
    event.preventDefault(); if (active.current) return;
    active.current = true; setBusy(true); setError('');
    try { await api.submit({ kind, title, details }); setSent(true); setTitle(''); setDetails(''); }
    catch (caught) { setError(`${caught.message} Your draft is still here. If access is denied or the list is missing, ask a site owner to configure feedback. If the request timed out, check with an owner before retrying.`); }
    finally { active.current = false; setBusy(false); }
  }
  async function inbox() {
    setTab('inbox'); setBusy(true); setError(''); setItems(null);
    try { setItems(await api.list()); } catch (caught) { setError(caught.message); }
    finally { setBusy(false); }
  }
  return <>
    <button ref={trigger} className="icon-button" aria-label="Report a bug or request a feature" title="Report a bug or request a feature" onClick={() => { setOpen(true); setTab('submit'); setError(''); setSent(false); }}><Icon name="bug" /></button>
    {open && <dialog ref={dialog} className="modal feedback-dialog" aria-labelledby="feedback-title" onCancel={event => { if (busy) event.preventDefault(); else setOpen(false); }} onClose={() => setOpen(false)}>
      <header><div><h2 id="feedback-title">Help improve the tracker</h2><p>Report an issue or share an idea for what comes next.</p></div><button className="icon-button" aria-label="Close feedback" disabled={busy} onClick={() => setOpen(false)}><Icon name="close" /></button></header>
      <div className="modal-body feedback-body">
        <p className="feedback-privacy"><Icon name="shield" size={16} /> {api.local ? 'Local preview · submissions disappear when the page reloads.' : 'Private feedback · only site owners can read submissions.'}</p>
        {owner && <div className="feedback-tabs"><button className={`button ${tab === 'submit' ? 'primary' : 'secondary'}`} disabled={busy} onClick={() => { setTab('submit'); setError(''); }}>Send feedback</button><button className={`button ${tab === 'inbox' ? 'primary' : 'secondary'}`} disabled={busy} onClick={inbox}>Owner inbox</button></div>}
        {error && <p role="alert" className="inline-error">{error}</p>}
        {tab === 'submit' ? sent ? <div className="feedback-success" role="status"><Icon name="check" size={32} /><h3>Thank you for the feedback</h3><p>{api.local ? 'Your preview submission is in the owner inbox.' : 'Your submission has been sent privately to the site owners.'}</p><button className="button secondary" onClick={() => setSent(false)}>Send another</button></div> : <form onSubmit={submit} className="feedback-form">
          <fieldset disabled={busy}><legend>What would you like to share?</legend><div className="feedback-types">{[['Bug', 'bug', 'Report a bug', 'Something isn’t working as expected.'], ['Feature', 'idea', 'Request a feature', 'An idea to make your work easier.']].map(([value, icon, label, hint]) => <label key={value} className={kind === value ? 'selected' : ''}><input type="radio" name="feedback-kind" value={value} checked={kind === value} onChange={() => setKind(value)} /><Icon name={icon} /><span><strong>{label}</strong><small>{hint}</small></span></label>)}</div></fieldset>
          <label className="field">Summary<input autoFocus required maxLength={200} value={title} disabled={busy} onChange={event => setTitle(event.target.value)} placeholder={kind === 'Bug' ? 'What went wrong?' : 'What would you like to do?'} /></label>
          <label className="field">{kind === 'Bug' ? 'What happened?' : 'Tell us about your idea'}<textarea required maxLength={10000} rows={7} value={details} disabled={busy} onChange={event => setDetails(event.target.value)} placeholder={kind === 'Bug' ? 'Include steps to reproduce, what you expected, and what happened instead.' : 'Describe the feature and how it would help your workflow.'} /></label>
          <div className="feedback-submit"><small>Please don’t include passwords or sensitive information.</small><button className="button primary" disabled={busy || !title.trim() || !details.trim()}>{busy ? 'Sending…' : 'Send feedback'}</button></div>
        </form> : <section aria-label="Feedback submissions" aria-busy={busy}>{busy ? <p role="status">Loading submissions…</p> : items?.length === 0 ? <div className="feedback-success"><Icon name="note" size={32} /><h3>No feedback yet</h3><p>Bug reports and feature requests will appear here.</p></div> : items?.map(item => <article key={item.id} className="feedback-item"><div><span className="feedback-kind"><Icon name={item.kind === 'Bug' ? 'bug' : 'idea'} size={15} />{item.kind === 'Bug' ? 'Bug report' : 'Feature request'}</span><time dateTime={item.created}>{new Date(item.created).toLocaleString()}</time></div><h3>{item.title}</h3><small>{item.author}</small><p>{item.details}</p></article>)}</section>}
      </div>
    </dialog>}
  </>;
}
