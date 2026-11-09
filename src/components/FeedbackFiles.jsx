import React, { useEffect, useRef, useState } from 'react';
import { FilePicker } from './FilePicker';
import { Icon } from './Icon';
import { attachmentSize, FEEDBACK_FILE_ACCEPT } from '../lib/feedbackAttachments';

function FileThumbnail({ file }) {
  const [url, setUrl] = useState('');
  useEffect(() => {
    if (!/\.(png|jpe?g|gif|webp)$/i.test(file.name) || !URL.createObjectURL) return;
    const next = URL.createObjectURL(file); setUrl(next);
    return () => URL.revokeObjectURL(next);
  }, [file]);
  return <span className="feedback-attachment-preview">{url ? <img src={url} alt="" /> : <Icon name="note" size={24} />}</span>;
}
export function FeedbackFileQueue({ files, busy, onAdd, onRemove }) {
  const [dragging, setDragging] = useState(false);
  return <section className={`feedback-attachment-dropzone${dragging ? ' dragging' : ''}`} aria-label="Feedback attachments" onDragOver={event => { if (!busy && event.dataTransfer.types.includes('Files')) { event.preventDefault(); setDragging(true); } }} onDragLeave={event => { if (!event.currentTarget.contains(event.relatedTarget)) setDragging(false); }} onDrop={event => { event.preventDefault(); setDragging(false); if (!busy) onAdd(Array.from(event.dataTransfer.files)); }}>
    <FilePicker label="Screenshots and attachments" hint="Optional · choose or drop files, or paste a screenshot into this form. Up to 5 files, 1 MB each." ariaLabel="Choose feedback attachments" accept={FEEDBACK_FILE_ACCEPT} disabled={busy} onChange={event => { onAdd(Array.from(event.target.files)); event.target.value = ''; }} />
    {files.length > 0 && <ul className="feedback-attachment-list">{files.map((file, index) => <li key={`${file.name}-${file.lastModified}-${index}`}>
      <FileThumbnail file={file} /><span className="feedback-attachment-info"><strong>{file.name}</strong><small>{attachmentSize(file.size)}</small></span>
      <button type="button" className="icon-button" aria-label={`Remove ${file.name}`} title="Remove attachment" disabled={busy} onClick={() => onRemove(index)}><Icon name="close" size={16} /></button>
    </li>)}</ul>}
  </section>;
}

export function FeedbackSavedFiles({ api, item }) {
  const [loading, setLoading] = useState(''), [error, setError] = useState(''), [previews, setPreviews] = useState({});
  const urls = useRef([]), mounted = useRef(true);
  useEffect(() => { mounted.current = true; return () => { mounted.current = false; urls.current.forEach(url => URL.revokeObjectURL(url)); }; }, []);
  async function load(file, preview) {
    if (loading) return;
    setLoading(file.key); setError('');
    try {
      const { blob } = await api.downloadAttachment(item.id, file.key);
      if (!mounted.current) return;
      const url = URL.createObjectURL(blob); urls.current.push(url);
      if (preview) setPreviews(value => ({ ...value, [file.key]: url }));
      else {
        const link = document.createElement('a'); link.href = url; link.download = file.name; document.body.append(link); link.click(); link.remove();
      }
    } catch (caught) { if (mounted.current) setError(caught.message); }
    finally { if (mounted.current) setLoading(''); }
  }
  if (!item.attachments?.length) return null;
  return <section className="feedback-saved-attachments" aria-label={`Attachments for ${item.title}`}>
    <h4><Icon name="paperclip" size={15} />Attachments <span>({item.attachments.length})</span></h4>
    <ul className="feedback-attachment-list">{item.attachments.map(file => <li key={file.key}>
      <span className="feedback-attachment-preview">{previews[file.key] ? <img src={previews[file.key]} alt={`Screenshot: ${file.name}`} /> : <Icon name="note" size={24} />}</span>
      <span className="feedback-attachment-info"><strong>{file.name}</strong><small>{attachmentSize(file.size)}</small></span>
      <div className="feedback-attachment-actions">
        {file.type.startsWith('image/') && !previews[file.key] && <button type="button" className="button secondary" disabled={!!loading} onClick={() => load(file, true)}>Preview</button>}
        <button type="button" className="button secondary" aria-label={`Download ${file.name}`} disabled={!!loading} onClick={() => load(file, false)}><Icon name="download" size={15} />{loading === file.key ? 'Loading…' : 'Download'}</button>
      </div>
      {previews[file.key] && <a className="feedback-screenshot" href={previews[file.key]} download={file.name}><img src={previews[file.key]} alt={file.name} /></a>}
    </li>)}</ul>
    {error && <p className="inline-error" role="alert">{error}</p>}
  </section>;
}
