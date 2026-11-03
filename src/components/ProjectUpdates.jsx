import React, { useEffect, useRef, useState } from 'react';
import { Icon } from './Icon';
import { displayName } from '../lib/displayName';
import { validateAttachment } from '../lib/taskAttachments';

export function updateTimestamp(item) {
  const value = item.postedAt || item.entryDate;
  if (!value) return 'Date not recorded';
  const hasTime = !!item.postedAt;
  const parsed = new Date(hasTime ? value : `${String(value).slice(0, 10)}T12:00:00`);
  return Number.isNaN(parsed.getTime()) ? 'Date not recorded' : parsed.toLocaleString(undefined, hasTime
    ? { year: 'numeric', month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit', timeZoneName: 'short' }
    : { year: 'numeric', month: 'short', day: 'numeric' });
}

function UpdateDocuments({ item, store, revision, canPost, onRefresh }) {
  const [files, setFiles] = useState([]), [error, setError] = useState(''), [busy, setBusy] = useState(false), [refresh, setRefresh] = useState(0);
  useEffect(() => {
    let active = true;
    store.listUpdateAttachments(item.id).then(rows => { if (active) { setFiles(rows); setError(''); } }).catch(err => { if (active) setError(err.message); });
    return () => { active = false; };
  }, [item.id, store, revision, refresh]);
  async function download(file) {
    setBusy(true); setError('');
    try {
      const blob = await store.downloadUpdateAttachment(item.id, file.name);
      const url = URL.createObjectURL(blob), link = document.createElement('a');
      link.href = url; link.download = file.name; document.body.append(link); link.click(); link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30000);
    } catch (err) { setError(err.message); } finally { setBusy(false); }
  }
  async function upload(event) {
    const selected = [...event.target.files]; event.target.value = ''; setBusy(true); setError('');
    const errors = [];
    for (const file of selected) {
      try { await store.addUpdateAttachment(item.id, file); } catch (err) { errors.push(`${file.name}: ${err.message}`); }
    }
    // Reload first; keep upload errors visible after a partial failure.
    try { setFiles(await store.listUpdateAttachments(item.id)); } catch (err) { errors.push(err.message); }
    await onRefresh?.();
    setError(errors.join('\n')); setBusy(false);
  }
  return <div className="update-documents">
    {!!files.length && <ul>{files.map(file => <li key={file.name}><button type="button" className="text-button" disabled={busy} onClick={() => download(file)}><Icon name="paperclip" size={14} />{file.name}</button></li>)}</ul>}
    {canPost && <details><summary>Attach documents to this update</summary><input className="document-upload" type="file" multiple aria-label={`Attach documents to update by ${item.authorName || 'Unknown author'}`} disabled={busy} onChange={upload} /><small>Up to 50 MB each.</small></details>}
    {busy && <p role="status">Working on update documents…</p>}
    {error && <p className="inline-error" role="alert">{error} <button type="button" className="text-button" disabled={busy} onClick={() => setRefresh(value => value + 1)}>Refresh documents</button></p>}
  </div>;
}

function UpdateCard({ item, tasks, store, revision, canPost, onEdit, onDelete, onTask, onRefresh }) {
  const [mode, setMode] = useState(''), [draft, setDraft] = useState(null), [busy, setBusy] = useState(false), [error, setError] = useState('');
  const active = useRef(false);
  const task = tasks.find(row => row.id === item.taskId);
  const canEdit = canPost && !item.legacyRisk;
  function begin(next) { setError(''); setDraft({ summary: item.summary, taskId: item.taskId || '', version: item.version }); setMode(next); }
  async function apply(event) {
    event.preventDefault(); if (active.current || !canEdit) return;
    active.current = true; setBusy(true); setError('');
    try {
      if (mode === 'delete') await onDelete({ ...item, version: draft.version });
      else { if (!draft.summary.trim()) throw new Error('Write an update before saving.'); await onEdit({ ...item, version: draft.version }, { summary: draft.summary, taskId: draft.taskId }); }
      setMode(''); setDraft(null);
    } catch (err) { setError(err.message); }
    finally { active.current = false; setBusy(false); }
  }
  return <article className="update-card">
    <header><span className="update-avatar" aria-hidden="true">{(displayName(item.authorName || '?')).slice(0, 1).toUpperCase()}</span><div><strong>{item.legacyRisk ? 'Risk owner · ' : ''}{displayName(item.authorName || 'Author not recorded')}</strong><time dateTime={item.postedAt || item.entryDate || undefined}>{updateTimestamp(item)}</time></div><span className={`update-kind ${item.type === 'Risk' ? 'update-risk' : ''}`}>{item.type || 'Status'}</span></header>
    {item.editedAt && <p className="update-edited">Edited by {displayName(item.editorName || 'Unknown editor')} · <time dateTime={item.editedAt}>{updateTimestamp({ postedAt: item.editedAt })}</time></p>}
    {item.taskId && (task ? <button type="button" className="update-task-reference" onClick={() => onTask(task)}><Icon name="task" size={14} />{task.title}</button> : <span className="update-task-reference">{item.taskTitle || 'Referenced task'} · Task removed</span>)}
    {item.legacyRisk && <small className="update-legacy">Previously recorded risk · {item.severity} · {item.riskStatus}</small>}
    {mode === 'edit' ? <form className="update-editor" onSubmit={apply} noValidate aria-label="Edit project update">
      <label className="field">Task reference <small>Optional</small><select aria-label="Edit task reference" disabled={busy} value={draft.taskId} onChange={event => setDraft(value => ({ ...value, taskId: event.target.value }))}><option value="">Project-wide update</option>{item.taskId && !task && <option value={item.taskId}>{item.taskTitle || 'Referenced task'} · Task removed</option>}{tasks.map(row => <option key={row.id} value={row.id}>{row.title}</option>)}</select></label>
      <label className="field">Update<textarea aria-label="Edit update" autoFocus rows={4} disabled={busy} value={draft.summary} onChange={event => setDraft(value => ({ ...value, summary: event.target.value }))} /></label>
      <div className="update-actions"><button type="button" className="button secondary" disabled={busy} onClick={() => { setMode(''); setError(''); }}>Cancel edit</button><button type="button" className="button primary" disabled={busy || !draft.summary.trim()} onClick={apply}>{busy ? 'Saving…' : 'Save changes'}</button></div>
    </form> : <p className="update-summary">{item.summary}</p>}
    {!item.legacyRisk && <UpdateDocuments item={item} store={store} revision={revision} canPost={canPost && !mode} onRefresh={onRefresh} />}
    {canEdit && !mode && <div className="update-actions update-actions-rest"><button type="button" className="text-button" disabled={!item.version} onClick={() => begin('edit')}>Edit update</button><button type="button" className="text-button update-delete" disabled={!item.version} onClick={() => begin('delete')}>Delete update</button>{!item.version && <small>Reload the tracker to manage this update.</small>}</div>}
    {mode === 'delete' && <div className="update-delete-confirm" role="group" aria-label="Delete this update?">
      <strong>Delete this update?</strong><p>This removes this update and its attached documents from the project history. Other updates stay in place.</p>
      <div className="update-actions"><button type="button" className="button secondary" disabled={busy} onClick={() => { setMode(''); setError(''); }}>Keep update</button><button type="button" className="button danger" disabled={busy} onClick={apply}>{busy ? 'Deleting…' : 'Confirm delete'}</button></div>
    </div>}
    {error && <p className="inline-error" role="alert">{error}</p>}
  </article>;
}

export function ProjectUpdates({ project, tasks, updates, risks = [], store, canPost, onPost, onEdit, onDelete, onTask, onRefresh }) {
  const [summary, setSummary] = useState(''), [taskId, setTaskId] = useState('');
  const [files, setFiles] = useState([]), [posted, setPosted] = useState(null), [busy, setBusy] = useState(false), [error, setError] = useState(''), [message, setMessage] = useState(''), [revision, setRevision] = useState(0);
  const active = useRef(false);
  function queue(event) {
    const selected = [...event.target.files]; event.target.value = '';
    const next = [...files], errors = [];
    for (const file of selected) {
      try { validateAttachment(file); if (next.some(row => row.name.toLowerCase() === file.name.toLowerCase())) throw new Error('Already selected.'); next.push(file); }
      catch (err) { errors.push(`${file.name}: ${err.message}`); }
    }
    setFiles(next); setError(errors.join('\n'));
  }
  function reset() { setSummary(''); setTaskId(''); setFiles([]); setPosted(null); }
  async function post(event) {
    event.preventDefault(); if (active.current || busy || !canPost) return;
    if (!summary.trim()) { setError('Write an update before posting.'); return; }
    active.current = true; setBusy(true); setError(''); setMessage('');
    try {
      const saved = posted || await onPost(project, { summary, type: 'Status', taskId });
      if (!saved?.id) throw new Error('The update was not confirmed. Your draft has been kept.');
      setPosted(saved);
      const failed = [], errors = [];
      for (const file of files) {
        try { await store.addUpdateAttachment(saved.id, file); }
        catch (err) { failed.push(file); errors.push(`${file.name}: ${err.message}`); }
      }
      setRevision(value => value + 1); setFiles(failed);
      await onRefresh?.();
      if (failed.length) setError(`Update posted. ${failed.length} document(s) could not upload. Retry only the remaining documents, or finish and attach them later.\n${errors.join('\n')}`);
      else { reset(); setMessage(files.length ? 'Update posted with all selected documents.' : 'Update posted.'); }
    } catch (err) { setError(err.message); }
    finally { active.current = false; setBusy(false); }
  }
  const history = [...updates, ...risks.map(risk => ({ ...risk, id: `risk-${risk.id}`, legacyRisk: true, type: 'Risk', summary: `${risk.title}\n${risk.mitigation || 'No mitigation recorded.'}`, authorName: risk.ownerName, riskStatus: risk.status, severity: risk.severity }))]
    .sort((a, b) => String(b.postedAt || b.entryDate || '').localeCompare(String(a.postedAt || a.entryDate || '')));
  return <div className="project-updates">
    {canPost && <form className="update-composer" noValidate onSubmit={post} aria-label="New project update">
      <div className="update-options"><label className="field">Task reference <small>Optional</small><select aria-label="Task reference" disabled={busy || !!posted} value={taskId} onChange={event => setTaskId(event.target.value)}><option value="">Project-wide update</option>{tasks.map(task => <option value={task.id} key={task.id}>{task.title}</option>)}</select></label></div>
      <label className="field">Your update<textarea aria-label="Your update" rows={4} disabled={busy || !!posted} value={summary} onChange={event => setSummary(event.target.value)} placeholder="What changed, what needs attention, or what was decided?" /></label>
      {!posted && <label className="update-attachment-picker"><Icon name="paperclip" size={16} /><span>Attach update documents <small>Optional · up to 50 MB each</small></span><input aria-label="Update documents" className="document-upload" type="file" multiple disabled={busy} onChange={queue} /></label>}
      {!!files.length && <ul className="update-pending-files">{files.map(file => <li key={file.name}><span>{file.name}</span><button type="button" className="text-button" disabled={busy} onClick={() => setFiles(rows => rows.filter(row => row !== file))}>Remove</button></li>)}</ul>}
      <div className="update-composer-footer"><div>{posted && <button type="button" className="button secondary" disabled={busy} onClick={() => { reset(); setError(''); setMessage('Update kept. Remaining documents were not attached.'); }}>Finish without remaining files</button>}<button type="button" className="button primary" disabled={busy || !summary.trim()} onClick={post}>{busy ? 'Posting…' : posted ? 'Retry documents' : 'Post update'}</button></div></div>
    </form>}
    {error && <p className="inline-error" role="alert">{error}</p>}{message && <p role="status">{message}</p>}
    <section className="update-history" aria-label="Project update history">
      {!history.length && <div className="empty-state"><h3>No updates yet</h3><p>Progress, decisions and risks will appear here with their author and date.</p></div>}
      {history.map(item => <UpdateCard key={item.id} item={item} tasks={tasks} store={store} revision={revision} canPost={canPost} onEdit={onEdit} onDelete={onDelete} onTask={onTask} onRefresh={onRefresh} />)}
    </section>
  </div>;
}
