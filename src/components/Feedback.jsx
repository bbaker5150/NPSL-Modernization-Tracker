import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { Icon } from './Icon';
import { createFeedback, validateFeedback } from '../lib/feedback';
import { FeedbackSetup } from './FeedbackSetup';

export function Feedback({ store }) {
  const api = useMemo(() => createFeedback(store), [store]);
  const dialog = useRef(null), trigger = useRef(null), active = useRef(false);
  const [open, setOpen] = useState(false), [owner, setOwner] = useState(false), [tab, setTab] = useState('submit');
  const [kind, setKind] = useState('Bug'), [title, setTitle] = useState(''), [details, setDetails] = useState('');
  const [busy, setBusy] = useState(false), [error, setError] = useState(''), [sent, setSent] = useState(false), [items, setItems] = useState(null);
  const [clearIds, setClearIds] = useState(null), [inboxMessage, setInboxMessage] = useState('');
  const [ownerError, setOwnerError] = useState('');
  const [setupState, setSetupState] = useState(null), [setupError, setSetupError] = useState(''), [setupMessage, setSetupMessage] = useState('');
  useEffect(() => {
    setOwner(false); setOwnerError(''); setSetupState(null); setSetupError(''); setSetupMessage(''); setClearIds(null); setInboxMessage('');
    if (!open) return;
    let mounted = true;
    api.isOwner().then(async value => {
      if (!mounted) return;
      setOwner(value);
      if (value) {
        try { const status = await api.setup.status(); if (mounted) setSetupState(status); }
        catch (caught) { if (mounted) setSetupError(caught.message); }
      }
    }).catch(() => {
      if (mounted) setOwnerError('Site-owner access could not be verified. Close and reopen feedback to retry. You can still send feedback.');
    });
    return () => { mounted = false; };
  }, [api, open]);
  useEffect(() => {
    if (!open) return;
    dialog.current.showModal();
    return () => { dialog.current?.close(); trigger.current?.focus(); };
  }, [open]);
  async function submit(event) {
    event.preventDefault(); if (active.current || busy || setupBlocked) return;
    let report;
    try { report = validateFeedback({ kind, title, details }); }
    catch (caught) { setError(caught.message); return; }
    active.current = true; setBusy(true); setError('');
    try { await api.submit(report); setSent(true); setTitle(''); setDetails(''); }
    catch (caught) { setError(`${caught.message} Your draft is still here. If access is denied or the list is missing, ask a site owner to configure feedback. If the request timed out, check with an owner before retrying.`); }
    finally { active.current = false; setBusy(false); }
  }
  const setupBlocked = owner && (!setupState?.ready || !!setupError);
  async function inbox() {
    setTab('inbox'); setBusy(true); setError(''); setItems(null); setClearIds(null); setInboxMessage('');
    try { setItems(await api.list()); } catch (caught) { setError(caught.message); }
    finally { setBusy(false); }
  }
  async function clearInbox() {
    if (!owner || busy || active.current || !clearIds?.length) return;
    active.current = true; setBusy(true); setError(''); setInboxMessage('Clearing submissions…');
    try {
      const { clearedIds } = await api.clearInbox(clearIds, ({ cleared, total }) => setInboxMessage(`Clearing submissions… ${cleared} of ${total}`));
      setItems(value => value?.filter(item => !clearedIds.includes(item.id)));
      setInboxMessage(`${clearedIds.length} submission${clearedIds.length === 1 ? '' : 's'} cleared.${api.local ? '' : ' Items can be restored from the SharePoint recycle bin.'}`);
      try { setItems(await api.list()); }
      catch (caught) { setError(`The submissions were cleared, but the inbox could not refresh. Reopen it to see new arrivals. ${caught.message}`); }
    } catch (caught) {
      if (caught.clearedIds) setItems(value => value?.filter(item => !caught.clearedIds.includes(item.id)));
      setInboxMessage(''); setError(caught.message);
    } finally { setClearIds(null); active.current = false; setBusy(false); }
  }
  return <>
    <button ref={trigger} className="icon-button" aria-label="Report a bug or request a feature" title="Report a bug or request a feature" onClick={() => { setOpen(true); setTab('submit'); setError(''); setSent(false); }}><Icon name="bug" /></button>
    {open && createPortal(<dialog ref={dialog} className="modal feedback-dialog" aria-labelledby="feedback-title" onCancel={event => { if (busy) event.preventDefault(); else setOpen(false); }} onClose={() => setOpen(false)}>
      <header><div><h2 id="feedback-title">{tab === 'inbox' ? 'Owner inbox' : 'Help improve the tracker'}</h2><p>{tab === 'inbox' ? 'Private bug reports and feature requests for site owners.' : 'Report an issue or share an idea for what comes next.'}</p></div><button className="icon-button" aria-label="Close feedback" disabled={busy} onClick={() => setOpen(false)}><Icon name="close" /></button></header>
      <div className="modal-body feedback-body">
        <p className="feedback-privacy"><Icon name="shield" size={16} /> {api.local ? 'Local preview · submissions disappear when the page reloads.' : 'Private feedback · only site owners can read submissions.'}</p>
        {owner && !setupState && !setupError && <p role="status">Checking feedback storage and permissions…</p>}
        {ownerError && <p role="alert" className="inline-error">{ownerError}</p>}
        {setupMessage && <p role="status">{setupMessage}</p>}
        {setupError && <p role="alert" className="inline-error">Feedback setup could not be verified: {setupError}</p>}
        {owner && setupState && !setupState.ready && <FeedbackSetup setup={api.setup} state={setupState} busy={busy} onBusy={setBusy} onReady={state => { setSetupState(state); setError(''); setSetupMessage(`${state.title} is ready in Site Contents.`); }} />}
        {tab === 'inbox' && inboxMessage && <p role="status" aria-live="polite">{inboxMessage}</p>}
        {error && <p role="alert" className="inline-error">{error}</p>}
        {tab === 'submit' ? sent ? <div className="feedback-success" role="status"><Icon name="check" size={32} /><h3>Thank you for the feedback</h3><p>{api.local ? 'Your preview submission is in the owner inbox.' : 'Your submission has been sent privately to the site owners.'}</p></div> : <form noValidate onSubmit={submit} className="feedback-form">
          <fieldset disabled={busy}><legend>What would you like to share?</legend><div className="feedback-types">{[['Bug', 'bug', 'Report a bug', 'Something isn’t working as expected.'], ['Feature', 'idea', 'Request a feature', 'An idea to make your work easier.']].map(([value, icon, label, hint]) => <label key={value} className={kind === value ? 'selected' : ''}><input type="radio" name="feedback-kind" value={value} checked={kind === value} onChange={() => setKind(value)} /><Icon name={icon} /><span><strong>{label}</strong><small>{hint}</small></span></label>)}</div></fieldset>
          <label className="field">Summary<input autoFocus required maxLength={200} value={title} disabled={busy} onChange={event => setTitle(event.target.value)} placeholder={kind === 'Bug' ? 'What went wrong?' : 'What would you like to do?'} /></label>
          <label className="field">{kind === 'Bug' ? 'What happened?' : 'Tell us about your idea'}<textarea required maxLength={10000} rows={7} value={details} disabled={busy} onChange={event => setDetails(event.target.value)} placeholder={kind === 'Bug' ? 'Include steps to reproduce, what you expected, and what happened instead.' : 'Describe the feature and how it would help your workflow.'} /></label>
          <small className="feedback-hint">Please don’t include passwords or sensitive information.</small>
        </form> : clearIds ? <section className="feedback-clear-confirm" aria-labelledby="feedback-clear-title"><Icon name="trash" size={28} /><h3 id="feedback-clear-title">Clear {clearIds.length} submission{clearIds.length === 1 ? '' : 's'}?</h3><p>This clears the submissions currently shown for all site owners. New feedback received after this inbox was loaded will be kept.</p><p>{api.local ? 'Local preview submissions will be removed.' : 'Submissions will move to the SharePoint recycle bin, where a site owner can restore them.'}</p></section> : <section aria-label="Feedback submissions" aria-busy={busy}>{busy ? <p role="status">Loading submissions…</p> : items?.length === 0 ? <div className="feedback-success"><Icon name="note" size={32} /><h3>No feedback yet</h3><p>Bug reports and feature requests will appear here.</p></div> : items?.map(item => <article key={item.id} className="feedback-item"><div><span className="feedback-kind"><Icon name={item.kind === 'Bug' ? 'bug' : 'idea'} size={15} />{item.kind === 'Bug' ? 'Bug report' : 'Feature request'}</span><time dateTime={item.created}>{new Date(item.created).toLocaleString()}</time></div><h3>{item.title}</h3><small>{item.author}</small><p>{item.details}</p></article>)}</section>}
      </div>
      <footer className="feedback-actions">
        {tab === 'inbox' ? clearIds ? <>
          <button type="button" className="button secondary" disabled={busy} onClick={() => setClearIds(null)}>Cancel</button>
          <button type="button" className="button feedback-clear" disabled={busy} onClick={clearInbox}>{busy ? 'Clearing…' : 'Confirm clear'}</button>
        </> : <>
          {owner && <button type="button" className="button feedback-clear" disabled={busy || !items?.length} onClick={() => { setClearIds(items.map(item => item.id)); setError(''); setInboxMessage(''); }}><Icon name="trash" size={16} />Clear inbox</button>}
          <button type="button" className="button primary" disabled={busy} onClick={() => { setTab('submit'); setSent(false); setError(''); }}>Back to feedback</button>
        </> : <>
          {owner && <button type="button" className="button secondary" disabled={busy || setupBlocked} onClick={inbox}><Icon name="note" size={16} />Owner inbox</button>}
          {sent ? <button type="button" className="button primary" onClick={() => setSent(false)}>Send another</button>
            : <button type="button" onClick={submit} className="button primary" disabled={busy || setupBlocked}>{busy ? 'Sending…' : 'Send feedback'}</button>}
        </>}
      </footer>
    </dialog>, document.body)}
  </>;
}
