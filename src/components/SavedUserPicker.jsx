import React, { useId, useState } from 'react';
import { displayName } from '../lib/displayName';
import { userIdentityKey } from '../lib/repository';

const optionKey = entry => entry.id || userIdentityKey(entry);

export function PeopleCombobox({
  id, value, disabled, placeholder, matches = [], listLabel,
  emptyMessage, onQuery, onSelect, renderOption, action,
}) {
  const [open, setOpen] = useState(false);
  const [activeKey, setActiveKey] = useState(null);
  const expanded = open && !disabled && Boolean(matches.length || emptyMessage);
  const active = matches.findIndex(entry => optionKey(entry) === activeKey);
  const listId = `${id}-options`;

  function close() {
    setOpen(false);
    setActiveKey(null);
  }

  function choose(entry) {
    close();
    onSelect(entry);
  }

  function handleKeyDown(event) {
    if (event.key === 'Escape' && open) {
      event.preventDefault();
      event.stopPropagation();
      close();
    } else if (matches.length && (event.key === 'ArrowDown' || event.key === 'ArrowUp')) {
      event.preventDefault();
      setOpen(true);
      const direction = event.key === 'ArrowDown' ? 1 : -1;
      const start = active < 0 && direction === -1 ? 0 : active;
      const next = (start + direction + matches.length) % matches.length;
      setActiveKey(optionKey(matches[next]));
      document.getElementById(`${id}-${next}`)?.scrollIntoView?.({ block: 'nearest' });
    } else if (event.key === 'Enter' && expanded && active >= 0) {
      event.preventDefault();
      choose(matches[active]);
    }
  }

  const input = <input
    id={id} type="text" role="combobox" aria-autocomplete="list" aria-expanded={expanded}
    aria-controls={expanded ? listId : undefined}
    aria-activedescendant={expanded && active >= 0 ? `${id}-${active}` : undefined}
    disabled={disabled} value={value || ''} autoComplete="off" placeholder={placeholder}
    onFocus={() => setOpen(true)}
    onChange={event => { setOpen(true); setActiveKey(null); onQuery(event.target.value); }}
    onKeyDown={handleKeyDown}
  />;
  return <div className="owner-autocomplete" onBlur={event => { if (!event.currentTarget.contains(event.relatedTarget)) close(); }}>
    {action ? <div className="input-action">{input}{action}</div> : input}
    {expanded && <div id={listId} role="listbox" aria-label={listLabel} className="owner-suggestions">
      {matches.map((entry, index) => <button
        key={optionKey(entry)} id={`${id}-${index}`} type="button" role="option" tabIndex={-1}
        aria-selected={active === index} onMouseDown={event => event.preventDefault()} onClick={() => choose(entry)}
      >{renderOption(entry)}</button>)}
      {!matches.length && <span className="field-hint">{emptyMessage}</span>}
    </div>}
  </div>;
}

export function SavedUserPicker({ label, users, value, disabled, onQuery, onSelect }) {
  const id = useId();
  const terms = String(value || '').toLowerCase().trim().split(/\s+/).filter(Boolean);
  const matches = users.filter((entry) => terms.every((term) => [entry.title, displayName(entry.title), entry.email, entry.loginName].join(' ').toLowerCase().includes(term))).slice(0, 8);
  return <div className="field saved-user-field">
    <label htmlFor={id}>{label}</label>
    <PeopleCombobox
      id={id} value={value === 'Unassigned' ? '' : value} disabled={disabled} placeholder="Type a saved user's name…"
      matches={matches} listLabel={`Saved users for ${label}`} emptyMessage="No matching saved users."
      onQuery={onQuery} onSelect={onSelect}
      renderOption={entry => <><strong>{displayName(entry.title)}</strong><small>{entry.email || entry.loginName} · {entry.role}</small></>}
    />
  </div>;
}
