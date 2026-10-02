import React, { useEffect, useRef, useState } from 'react';

function DocumentActions({ file, busy, canEdit, onDownload, onRename, onDelete }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return;
    const dismiss = (event) => { if (event.type === 'keydown' ? event.key === 'Escape' : !ref.current?.contains(event.target)) setOpen(false); };
    document.addEventListener('pointerdown', dismiss); document.addEventListener('keydown', dismiss);
    return () => { document.removeEventListener('pointerdown', dismiss); document.removeEventListener('keydown', dismiss); };
  }, [open]);
  const action = (callback) => { setOpen(false); callback(file); };
  return <div className="reference-actions" ref={ref}>
    <button type="button" className="button secondary" disabled={busy} aria-label={`Edit ${file.name}`} aria-expanded={open} onClick={() => setOpen(!open)}>Edit</button>
    {open && <div className="reference-action-menu" aria-label={`Actions for ${file.name}`}>
      <button type="button" onClick={() => action(onDownload)}>Download</button>
      {canEdit && <><button type="button" onClick={() => action(onRename)}>Rename</button><button type="button" onClick={() => action(onDelete)}>Delete</button></>}
    </div>}
  </div>;
}

export function TaskDocuments({ taskId, projectKey, store, readOnly, canDelete = false, refreshKey, onChange, compact = false }) {
  const [files, setFiles] = useState([]);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [message, setMessage] = useState('');
  const [renaming, setRenaming] = useState(null);
  const [refresh, setRefresh] = useState(0);
  useEffect(() => {
    let active = true;
    const urls = [];
    setLoading(true);
    setError('');
    setFiles([]);
    (projectKey ? store.listProjectAttachments(projectKey) : store.listTaskAttachments(taskId)).then((rows) => {
      if (!active) return;
      setFiles(rows.map((row) => {
        if (!row.blob) return row;
        const url = URL.createObjectURL(row.blob); urls.push(url);
        return { ...row, url };
      }));
    }).catch((err) => { if (active) setError(err.message); }).finally(() => { if (active) setLoading(false); });
    return () => { active = false; urls.forEach((url) => URL.revokeObjectURL(url)); };
  }, [taskId, projectKey, store, refresh, refreshKey]);
  async function upload(event) {
    const selected = [...event.target.files];
    event.target.value = '';
    if (!selected.length) return;
    setBusy(true); setError(''); setMessage('');
    let uploaded = 0;
    const failures = [];
    for (const file of selected) {
      try { await store.addTaskAttachment(taskId, file); uploaded += 1; }
      catch (err) { failures.push(`${file.name}: ${err.message}`); }
    }
    setMessage(`${uploaded} of ${selected.length} documents uploaded.`);
    if (uploaded) onChange?.();
    setError(failures.join('\n')); setBusy(false); setRefresh((value) => value + 1);
  }
  async function download(file) {
    setBusy(true); setError('');
    try {
      const blob = await store.downloadTaskAttachment(file.taskId || taskId, file.name);
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url; link.download = file.name;
      document.body.appendChild(link); link.click(); link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30000);
    } catch (err) { setError(err.message); }
    finally { setBusy(false); }
  }
  async function remove(file) {
    setBusy(true); setError(''); setMessage('');
    try { await store.deleteTaskAttachment(file.taskId || taskId, file.name); setMessage(`${file.name} deleted.`); onChange?.(); }
    catch (err) { setError(err.message); }
    finally { setBusy(false); setRefresh((value) => value + 1); }
  }
  async function rename() {
    if (!renaming?.name.trim()) return;
    setBusy(true); setError(''); setMessage('');
    try {
      await store.renameTaskAttachment(renaming.file.taskId || taskId, renaming.file.name, renaming.name.trim());
      setMessage(`${renaming.file.name} renamed.`); setRenaming(null);
      onChange?.();
    } catch (err) { setError(err.message); }
    finally { setBusy(false); setRefresh((value) => value + 1); }
  }
  if (compact && !loading && !files.length && !error) return null;
  return <section className={`task-documents${compact ? ' task-documents-compact' : ''}`}><h3>Documents</h3>
    {renaming && <div className="reference-editor document-rename"><h4>Rename document</h4><label className="field"><span>Document name</span><input autoFocus value={renaming.name} disabled={busy} onChange={(event) => setRenaming({ ...renaming, name: event.target.value })} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); rename(); } }} /></label><div className="invite-actions"><button type="button" className="button primary" disabled={busy || !renaming.name.trim()} onClick={rename}>Save</button><button type="button" className="button secondary" disabled={busy} onClick={() => setRenaming(null)}>Cancel</button></div></div>}
    {loading && <p role="status">Loading documents…</p>}
    {!loading && !files.length && <p>No documents attached.</p>}
    <ul className="document-list">{files.map((file) => <li key={`${file.taskId || taskId}/${file.name}`}><div className="document-name"><button type="button" className="text-button document-file-name" disabled={busy} onClick={() => download(file)}>{file.name}</button>{file.taskTitle && <small>{file.taskTitle}</small>}</div>{!compact && <DocumentActions file={file} busy={busy || loading} canEdit={canDelete} onDownload={download} onRename={(row) => setRenaming({ file: row, name: row.name })} onDelete={remove} />}</li>)}</ul>
    {!readOnly && <label className="field"><span>{busy ? 'Uploading documents…' : 'Attach documents (up to 50 MB each)'}</span><input aria-label="Attach documents" className="document-upload" type="file" multiple disabled={busy || loading} onChange={upload} /></label>}
    {message && <p role="status">{message}</p>}{error && <p role="alert" className="inline-error">{error}</p>}
  </section>;
}
