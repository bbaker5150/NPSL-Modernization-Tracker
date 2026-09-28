import React, { useEffect, useState } from 'react';

export function TaskDocuments({ taskId, projectKey, store, readOnly, canDelete = false, refreshKey }) {
  const [files, setFiles] = useState([]);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [message, setMessage] = useState('');
  const [refresh, setRefresh] = useState(0);
  useEffect(() => {
    let active = true;
    const urls = [];
    setLoading(true);
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
    setError(failures.join('\n')); setBusy(false); setRefresh((value) => value + 1);
  }
  async function remove(file) {
    if (!window.confirm(`Delete “${file.name}”? This cannot be undone.`)) return;
    setBusy(true); setError(''); setMessage('');
    try { await store.deleteTaskAttachment(file.taskId || taskId, file.name); setMessage(`${file.name} deleted.`); }
    catch (err) { setError(err.message); }
    finally { setBusy(false); setRefresh((value) => value + 1); }
  }
  return <section className="task-documents"><h3>Documents</h3>
    {loading && <p role="status">Loading documents…</p>}
    {!loading && !files.length && <p>No documents attached.</p>}
    <ul className="document-list">{files.map((file) => <li key={`${file.taskId || taskId}/${file.name}`}><div className="document-name"><a href={file.url} target="_blank" rel="noopener noreferrer">{file.name}</a>{file.taskTitle && <small>{file.taskTitle}</small>}</div><div className="document-actions"><a className="button secondary" href={file.url} download={file.name}>Download</a>{canDelete && <button type="button" className="button secondary" disabled={busy || loading} onClick={() => remove(file)} aria-label={`Delete ${file.name}`}>Delete</button>}</div></li>)}</ul>
    {!readOnly && <label className="field"><span>{busy ? 'Uploading documents…' : 'Attach documents (up to 20 MB each)'}</span><input aria-label="Attach documents" className="document-upload" type="file" multiple disabled={busy || loading} onChange={upload} /></label>}
    {message && <p role="status">{message}</p>}{error && <p role="alert" className="inline-error">{error}</p>}
    {!busy && <button type="button" className="text-button" onClick={() => { setError(''); setRefresh((value) => value + 1); }}>Refresh documents</button>}
  </section>;
}
