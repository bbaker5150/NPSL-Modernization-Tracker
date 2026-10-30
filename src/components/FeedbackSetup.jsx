import React, { useEffect, useState } from 'react';

export function FeedbackSetup({ setup, state, busy, onBusy, onReady }) {
  const [options, setOptions] = useState(null), [selected, setSelected] = useState([]);
  const [error, setError] = useState(''), [progress, setProgress] = useState('');
  useEffect(() => {
    let mounted = true;
    setup.options().then(value => { if (mounted) { setOptions(value); setSelected(value.selected); } }).catch(caught => { if (mounted) setError(caught.message); });
    return () => { mounted = false; };
  }, [setup]);
  async function create() {
    if (busy) return;
    onBusy(true); setError('');
    try { onReady(await setup.provision(selected, setProgress)); }
    catch (caught) { setError(`${caught.message} Setup may be partially complete. Close and reopen feedback to check before retrying.`); }
    finally { onBusy(false); setProgress(''); }
  }
  return <section className="feedback-setup" aria-label="Set up private feedback">
    <h3>{state.partial ? 'Finish setting up private feedback' : 'Set up private feedback'}</h3>
    <p>{state.title} {state.partial ? 'needs its setup completed' : 'has not been created on this site'}. Create the private inbox here before collecting feedback.</p>
    {options && <>
      <p><strong>{options.owners.Title}</strong> will be able to read submissions. Select every group containing tracker users; these groups can submit feedback without reading it.</p>
      <fieldset disabled={busy}><legend>Groups allowed to submit</legend><div className="feedback-setup-groups">{options.groups.map(group => <label key={group.Id}><input type="checkbox" checked={selected.includes(group.Id)} onChange={event => setSelected(value => event.target.checked ? [...value, group.Id] : value.filter(entry => entry !== group.Id))} />{group.Title}</label>)}</div></fieldset>
      <button type="button" className="button primary" disabled={busy || !selected.length} onClick={create}>{busy ? 'Setting up…' : state.partial ? 'Finish setup' : 'Create private feedback list'}</button>
    </>}
    {!options && !error && <p role="status">Loading SharePoint groups…</p>}
    {progress && <p role="status">{progress}</p>}
    {error && <p role="alert" className="inline-error">{error}</p>}
  </section>;
}
