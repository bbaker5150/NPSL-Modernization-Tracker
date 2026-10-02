import React, { useId, useState } from 'react';
import { displayName } from '../lib/displayName';
import { userIdentityKey } from '../lib/repository';

export function SavedUserPicker({ label, users, value, disabled, onQuery, onSelect }) {
  const id = useId();
  const [open, setOpen] = useState(false);
  const [active, setActive] = useState(-1);
  const terms = String(value || '').toLowerCase().trim().split(/\s+/).filter(Boolean);
  const matches = users.filter((entry) => terms.every((term) => [entry.title, displayName(entry.title), entry.email, entry.loginName].join(' ').toLowerCase().includes(term))).slice(0, 8);
  const choose = (entry) => { onSelect(entry); setOpen(false); setActive(-1); };
  return <div className="field saved-user-field" onBlur={(event) => { if (!event.currentTarget.contains(event.relatedTarget)) setOpen(false); }}>
    <label htmlFor={id}>{label}</label>
    <div className="owner-autocomplete">
      <input id={id} role="combobox" aria-autocomplete="list" aria-expanded={open && !disabled} aria-controls={`${id}-options`} aria-activedescendant={open && matches[active] ? `${id}-${active}` : undefined} disabled={disabled} value={value === 'Unassigned' ? '' : value || ''} autoComplete="off" placeholder="Type a saved user's name…" onFocus={() => setOpen(true)} onChange={(event) => { onQuery(event.target.value); setOpen(true); setActive(-1); }} onKeyDown={(event) => {
        if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); setOpen(false); }
        if (event.key === 'ArrowDown' || event.key === 'ArrowUp') { event.preventDefault(); setOpen(true); setActive((index) => !matches.length ? -1 : index < 0 ? (event.key === 'ArrowDown' ? 0 : matches.length - 1) : (index + (event.key === 'ArrowDown' ? 1 : -1) + matches.length) % matches.length); }
        if (event.key === 'Enter' && open) { event.preventDefault(); if (matches[active]) choose(matches[active]); }
      }} />
      {open && !disabled && <div id={`${id}-options`} role="listbox" aria-label={`Saved users for ${label}`} className="owner-suggestions">
        {matches.map((entry, index) => <button key={entry.id || userIdentityKey(entry)} id={`${id}-${index}`} type="button" role="option" aria-selected={active === index} onMouseDown={(event) => event.preventDefault()} onClick={() => choose(entry)}><strong>{displayName(entry.title)}</strong><small>{entry.email || entry.loginName} · {entry.role}</small></button>)}
        {!matches.length && <span className="field-hint">No matching saved users.</span>}
      </div>}
    </div>
  </div>;
}
