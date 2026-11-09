import React, { useId } from 'react';
import { Icon } from './Icon';

/** A native, keyboard-accessible file input with the tracker's shared control styling. */
export function FilePicker({ label = 'Attach documents', hint, ariaLabel, multiple = true, disabled = false, accept, onChange, className = '' }) {
  const id = useId();
  return <label className={`file-picker${disabled ? ' is-disabled' : ''}${className ? ` ${className}` : ''}`} htmlFor={id}>
    <input id={id} className="file-picker-input" type="file" multiple={multiple} disabled={disabled} accept={accept}
      aria-label={ariaLabel || label} aria-describedby={hint ? `${id}-hint` : undefined} onChange={onChange} />
    <span className="file-picker-icon" aria-hidden="true"><Icon name="paperclip" size={20} /></span>
    <span className="file-picker-copy"><strong>{label}</strong>{hint && <small id={`${id}-hint`}>{hint}</small>}</span>
    <span className="button secondary file-picker-button" aria-hidden="true">{multiple ? 'Choose files' : 'Choose file'}</span>
  </label>;
}
