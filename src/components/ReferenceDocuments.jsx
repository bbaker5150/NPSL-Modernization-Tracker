import React, { useEffect, useRef, useState } from 'react';
import { Icon } from './Icon';
import { referencePath } from '../lib/referenceDocuments';

function ReferenceActions({ row, busy, readOnly, onDownload, onEdit, onDelete }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return;
    const dismiss = (event) => { if (event.type === 'keydown' ? event.key === 'Escape' : !ref.current?.contains(event.target)) setOpen(false); };
    document.addEventListener('pointerdown', dismiss);
    document.addEventListener('keydown', dismiss);
    return () => { document.removeEventListener('pointerdown', dismiss); document.removeEventListener('keydown', dismiss); };
  }, [open]);
  const action = (callback) => { setOpen(false); callback(); };
  return <div className="reference-actions" ref={ref}>
    <button type="button" className="button secondary" disabled={busy} aria-label={`Edit ${row.name}`} aria-expanded={open} onClick={() => setOpen(!open)}>Edit</button>
    {open && <div className="reference-action-menu" aria-label={`Actions for ${row.name}`}>
      {row.kind === 'file' && <button type="button" onClick={() => action(() => onDownload(row))}>Download</button>}
      {!readOnly && <><button type="button" onClick={() => action(() => onEdit(row, 'rename'))}>Rename</button><button type="button" onClick={() => action(() => onEdit(row, 'move'))}>Move</button><button type="button" onClick={() => action(() => onDelete(row))}>Delete</button></>}
    </div>}
  </div>;
}

export function ReferenceDocuments({ store, readOnly }) {
  const [entries, setEntries] = useState([]);
  const [folder, setFolder] = useState('');
  const [query, setQuery] = useState('');
  const [draft, setDraft] = useState(null);
  const [busy, setBusy] = useState(false);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [message, setMessage] = useState('');
  useEffect(() => {
    let active = true;
    store.listReferenceEntries().then((rows) => { if (active) setEntries(rows); })
      .catch((err) => { if (active) setError(err.message); })
      .finally(() => { if (active) setLoading(false); });
    return () => { active = false; };
  }, [store]);
  async function refresh() { setEntries(await store.listReferenceEntries()); }
  async function save() {
    if (busy) return;
    setBusy(true); setError(''); setMessage('');
    try { await store.saveReferenceEntry(draft.id ? draft : { ...draft, parentId: folder }); await refresh(); setDraft(null); setMessage('Reference library updated.'); }
    catch (err) { setError(err.message); }
    finally { setBusy(false); }
  }
  async function upload(event) {
    const files = [...event.target.files];
    event.target.value = '';
    if (!files.length || busy) return;
    setBusy(true); setError(''); setMessage('');
    const failures = [];
    let count = 0;
    for (const file of files) {
      try { await store.saveReferenceEntry({ name: file.name, parentId: folder }, file); count++; }
      catch (err) { failures.push(`${file.name}: ${err.message}`); }
    }
    try { await refresh(); } catch (err) { failures.push(err.message); }
    setMessage(`${count} of ${files.length} documents uploaded.`);
    setError(failures.join('\n')); setBusy(false);
  }
  async function download(row) {
    setBusy(true); setError('');
    try {
      const blob = await store.downloadReferenceEntry(row.id);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = row.name;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30000);
    } catch (err) { setError(err.message); }
    finally { setBusy(false); }
  }
  async function remove(row) {
    if (busy) return;
    setBusy(true); setError(''); setMessage('');
    try { await store.deleteReferenceEntry(row.id); await refresh(); setDraft(null); setMessage(`${row.name} deleted.`); }
    catch (err) { setError(err.message); }
    finally { setBusy(false); }
  }
  function navigate(id) { setFolder(id); setQuery(''); setDraft(null); setError(''); }
  const search = query.trim().toLowerCase();
  const rows = entries.filter((row) => search ? row.name.toLowerCase().includes(search) : row.parentId === folder)
    .sort((a, b) => (a.kind === b.kind ? 0 : a.kind === 'folder' ? -1 : 1) || a.name.localeCompare(b.name));
  const folders = entries.filter((row) => row.kind === 'folder' && !referencePath(entries, row.id).some((part) => part.id === draft?.id));
  return <section className="page-stack reference-library">
    <section className="page-heading"><div><span className="section-kicker">Shared resources</span><h1>Reference Documents</h1><p>Find modernization templates and guidance, organized by type.</p></div></section>
    <section className="panel reference-panel">
      <div className="reference-toolbar"><input type="search" aria-label="Search reference documents" placeholder="Search all folders and documents" value={query} onChange={(event) => setQuery(event.target.value)} />
        {!readOnly && <button type="button" className="button secondary" disabled={busy || loading} onClick={() => { setDraft({ name: '', parentId: folder }); setError(''); }}><Icon name="plus" />New folder</button>}
      </div>
      <nav className="reference-breadcrumbs" aria-label="Reference document folders">
        <button type="button" disabled={busy} onClick={() => navigate('')}>All references</button>
        {referencePath(entries, folder).map((row) => <React.Fragment key={row.id}><Icon name="chevron" size={14} /><button type="button" disabled={busy} onClick={() => navigate(row.id)}>{row.name}</button></React.Fragment>)}
      </nav>
      {draft && <div className="reference-editor">
        <h2>{draft.id ? draft.mode === 'move' ? 'Move to folder' : 'Rename' : 'New folder'}</h2>
        {draft.mode !== 'move' && <label className="field"><span>{draft.kind === 'file' ? 'Document name' : 'Folder name'}</span><input autoFocus value={draft.name} disabled={busy} onChange={(event) => setDraft({ ...draft, name: event.target.value })} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); save(); } }} /></label>}
        {draft.id && draft.mode === 'move' && <label className="field"><span>Folder location</span><select aria-label="Folder location" value={draft.parentId} disabled={busy} onChange={(event) => setDraft({ ...draft, parentId: event.target.value })}><option value="">All references</option>{folders.map((row) => <option key={row.id} value={row.id}>{referencePath(entries, row.id).map((part) => part.name).join(' / ')}</option>)}</select></label>}
        <div className="invite-actions"><button type="button" className="button primary" disabled={busy || !draft.name.trim()} onClick={save}>{busy ? 'Saving…' : draft.id ? 'Save reference' : 'Create folder'}</button><button type="button" className="button secondary" disabled={busy} onClick={() => setDraft(null)}>Cancel</button></div>
      </div>}
      {loading && <p role="status">Loading references…</p>}
      {!loading && !rows.length && <p className="empty-state">{search ? 'No matching references.' : 'No references in this folder yet.'}</p>}
      <ul className="reference-list">{rows.map((row) => <li key={row.id}>
        <Icon name={row.kind === 'folder' ? 'folder' : 'note'} size={22} />
        <div className="reference-name"><button type="button" className="text-button document-file-name" disabled={busy} onClick={() => row.kind === 'folder' ? navigate(row.id) : download(row)}>{row.name}</button>
          <small>{row.kind === 'folder' ? `${entries.filter((entry) => entry.parentId === row.id).length} items` : `${Math.max(1, Math.round(row.size / 1024)).toLocaleString()} KB`}{search && ` · ${referencePath(entries, row.parentId).map((part) => part.name).join(' / ') || 'All references'}`}</small>
        </div>
        {(!readOnly || row.kind === 'file') && <ReferenceActions row={row} busy={busy} readOnly={readOnly} onDownload={download} onEdit={(entry, mode) => { setDraft({ ...entry, mode }); setError(''); }} onDelete={remove} />}

      </li>)}</ul>
      {!readOnly && <label className="field"><span>Upload to {entries.find((row) => row.id === folder)?.name || 'All references'} (up to 20 MB each)</span><input type="file" multiple className="document-upload" aria-label="Upload reference documents" disabled={busy || loading} onChange={upload} /></label>}
      {message && <p role="status">{message}</p>}{error && <p role="alert" className="inline-error">{error}</p>}
    </section>
  </section>;
}
