import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, it, vi } from 'vitest';
import { ProjectUpdates, updateTimestamp } from './ProjectUpdates';
globalThis.IS_REACT_ACT_ENVIRONMENT = true;
const button = label => [...document.querySelectorAll('button')].find(node => node.textContent === label);
const change = (node, value) => {
  Object.getOwnPropertyDescriptor(node instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLTextAreaElement.prototype, 'value').set.call(node, value);
  node.dispatchEvent(new Event(node instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true }));
};
it('keeps failed update files for retry without posting the update twice', async () => {
  const store = { addUpdateAttachment: vi.fn().mockResolvedValueOnce().mockRejectedValueOnce(new Error('Upload failed')).mockResolvedValueOnce(), listUpdateAttachments: vi.fn().mockResolvedValue([]) };
  const onPost = vi.fn().mockResolvedValue({ id: 'u' });
  const host = document.createElement('div'); document.body.append(host); const root = createRoot(host);
  const cancelNative = event => { if (event.target.closest('button')) event.preventDefault(); };
  document.addEventListener('click', cancelNative, true);
  try {
    await act(async () => root.render(<ProjectUpdates project={{ projectKey: 'p' }} tasks={[{ id: 't', title: 'Review' }]} updates={[]} store={store} canPost onPost={onPost} />));
    await act(async () => change(document.querySelector('textarea'), 'Decision recorded'));
    const input = document.querySelector('input[type=file]');
    Object.defineProperty(input, 'files', { value: [new File(['one'], 'one.txt'), new File(['two'], 'two.txt')] });
    await act(async () => input.dispatchEvent(new Event('change', { bubbles: true })));
    await act(async () => button('Post update').click());
    expect(document.querySelector('[role=alert]').textContent).toContain('Update posted');
    expect(document.querySelector('.update-pending-files').textContent).toContain('two.txt');
    expect(document.querySelector('.update-pending-files').textContent).not.toContain('one.txt');
    await act(async () => button('Retry documents').click());
    expect(onPost).toHaveBeenCalledTimes(1);
    expect(onPost).toHaveBeenCalledWith({ projectKey: 'p' }, { summary: 'Decision recorded', type: 'Status', taskId: '' });
    expect(store.addUpdateAttachment.mock.calls.map(([, file]) => file.name)).toEqual(['one.txt', 'two.txt', 'two.txt']);
    expect(document.querySelector('textarea').value).toBe('');
  } finally { document.removeEventListener('click', cancelNative, true); await act(async () => root.unmount()); host.remove(); }
});
it('renders author, time, optional task links and legacy risks for readers without posting controls', async () => {
  const store = { listUpdateAttachments: vi.fn().mockResolvedValue([]) }, onTask = vi.fn();
  const host = document.createElement('div'); document.body.append(host); const root = createRoot(host);
  try {
    await act(async () => root.render(<ProjectUpdates project={{}} tasks={[{ id: 't', title: 'Review requirement' }]} updates={[{ id: 'u', summary: '<script>plain text</script>', authorName: 'Alex Engineer', postedAt: '2026-11-02T18:05:00Z', taskId: 't' }, { id: 'u2', summary: 'General update', authorName: 'Sam', entryDate: '2026-10-01' }]} risks={[{ id: 'r', title: 'Legacy funding risk', ownerName: 'Jordan', severity: 'High', status: 'Open' }]} store={store} canPost={false} onTask={onTask} />));
    expect(document.querySelector('form')).toBeNull();
    expect(document.body.textContent).toContain('Alex Engineer');
    expect(document.querySelector('time').dateTime).toBe('2026-11-02T18:05:00Z');
    expect(document.body.textContent).toContain('Legacy funding risk');
    expect(document.querySelector('script')).toBeNull();
    await act(async () => button('Review requirement').click());
    expect(onTask).toHaveBeenCalledWith({ id: 't', title: 'Review requirement' });
    expect(store.listUpdateAttachments).toHaveBeenCalledTimes(2);
  } finally { await act(async () => root.unmount()); host.remove(); }
});
it('does not invent timestamps for old date-only updates', () => {
  expect(updateTimestamp({ entryDate: '2026-10-01' })).not.toContain(':');
  expect(updateTimestamp({})).toBe('Date not recorded');
});

it('edits and deletes only the selected update with draft retention, cancellation and host-safe buttons', async () => {
  const store = { listUpdateAttachments: vi.fn().mockResolvedValue([{ name: 'decision.txt' }]) };
  const onEdit = vi.fn().mockRejectedValueOnce(new Error('Conflict: keep your draft')).mockResolvedValueOnce();
  const onDelete = vi.fn().mockRejectedValueOnce(new Error('Delete denied')).mockResolvedValueOnce();
  const original = { id: 'u', version: '1', summary: 'Original summary', authorName: 'Original author', postedAt: '2026-11-02T18:05:00Z', taskId: 't' };
  function Harness() {
    const [updates, setUpdates] = React.useState([original, { ...original, id: 'other', summary: 'Other update' }]);
    return <ProjectUpdates project={{ projectKey: 'p' }} tasks={[{ id: 't', title: 'Review' }]} updates={updates} store={store} canPost onEdit={async (item, changes) => { await onEdit(item, changes); setUpdates(rows => rows.map(row => row.id === item.id ? { ...row, ...changes, editorName: 'Manager', editedAt: '2026-11-03T18:05:00Z', version: '2' } : row)); }} onDelete={async item => { await onDelete(item); setUpdates(rows => rows.filter(row => row.id !== item.id)); }} />;
  }
  const host = document.createElement('div'); document.body.append(host); const root = createRoot(host);
  const cancelNative = event => { if (event.target.closest('button')) event.preventDefault(); };
  document.addEventListener('click', cancelNative, true);
  try {
    await act(async () => root.render(<Harness />));
    await act(async () => button('Edit update').click());
    await act(async () => change(document.querySelector('[aria-label="Edit update"]'), 'Discard this'));
    await act(async () => button('Cancel edit').click());
    expect(onEdit).not.toHaveBeenCalled();
    await act(async () => button('Edit update').click());
    expect(document.querySelector('[aria-label="Edit update"]').value).toBe('Original summary');
    await act(async () => change(document.querySelector('[aria-label="Edit update"]'), 'Revised summary'));
    await act(async () => change(document.querySelector('[aria-label="Edit task reference"]'), ''));
    await act(async () => button('Save changes').click());
    expect(document.querySelector('[role=alert]').textContent).toContain('keep your draft');
    expect(document.querySelector('[aria-label="Edit update"]').value).toBe('Revised summary');
    await act(async () => button('Save changes').click());
    expect(onEdit).toHaveBeenLastCalledWith(original, { summary: 'Revised summary', taskId: '' });
    const card = document.querySelector('.update-card');
    expect(card.querySelector('header strong').textContent).toBe('Original author');
    expect(card.querySelector('header time').dateTime).toBe(original.postedAt);
    expect(card.querySelector('.update-edited').textContent).toContain('Edited by Manager');
    expect(card.querySelector('.update-task-reference')).toBeNull();
    expect(card.textContent).toContain('decision.txt');
    await act(async () => button('Delete update').click());
    await act(async () => button('Keep update').click());
    expect(onDelete).not.toHaveBeenCalled();
    await act(async () => button('Delete update').click());
    await act(async () => button('Confirm delete').click());
    expect(document.querySelector('[role=alert]').textContent).toBe('Delete denied');
    expect(document.querySelectorAll('.update-card')).toHaveLength(2);
    await act(async () => button('Confirm delete').click());
    expect(document.querySelectorAll('.update-card')).toHaveLength(1);
    expect(document.querySelector('.update-card').textContent).toContain('Other update');
    expect(onDelete).toHaveBeenLastCalledWith(expect.objectContaining({ id: 'u', version: '2' }));
  } finally { document.removeEventListener('click', cancelNative, true); await act(async () => root.unmount()); host.remove(); }
});
it.each([false, true])('allows risk changes and attachment removal only for project writers (canPost=%s)', async canPost => {
  const host = document.createElement('div'); document.body.append(host); const root = createRoot(host);
  const onEdit = vi.fn(), onDelete = vi.fn();
  const store = { listUpdateAttachments: async () => [{ name: 'report.pdf' }], deleteUpdateAttachment: vi.fn() };
  try {
    await act(async () => root.render(<ProjectUpdates project={{}} tasks={[]} updates={[{ id: 'u', version: '1', summary: 'By another user' }]} risks={[{ id: 'r', version: '2', title: 'Old risk', ownerName: 'Original owner' }]} store={store} canPost={canPost} onEdit={onEdit} onDelete={onDelete} />));
    const remove = () => document.querySelector('[aria-label="Remove report.pdf"]');
    if (!canPost) {
      for (const action of ['Edit update', 'Delete update', 'Edit risk', 'Delete risk']) expect(button(action)).toBeUndefined();
      expect(remove()).toBeNull();
      return;
    }
    await act(async () => button('Edit risk').click());
    await act(async () => change(document.querySelector('.update-editor textarea'), 'Mitigation after review'));
    await act(async () => button('Save changes').click());
    expect(onEdit).toHaveBeenCalledWith(expect.objectContaining({ id: 'r', legacyRisk: true, version: '2' }), expect.objectContaining({ title: 'Old risk', mitigation: 'Mitigation after review' }));
    await act(async () => button('Delete risk').click());
    await act(async () => button('Confirm delete').click());
    expect(onDelete).toHaveBeenCalledWith(expect.objectContaining({ id: 'r', version: '2', legacyRisk: true }));
    await act(async () => remove().click());
    expect(store.deleteUpdateAttachment).not.toHaveBeenCalled();
    await act(async () => document.querySelector('[aria-label="Confirm remove report.pdf"]').click());
    expect(store.deleteUpdateAttachment).toHaveBeenCalledExactlyOnceWith('u', 'report.pdf', '1');
  } finally { await act(async () => root.unmount()); host.remove(); }
});
it('explains a pending editing upgrade without blocking history or individual deletion', async () => {
  const store = { updateEditingAvailable: false, listUpdateAttachments: async () => [] };
  const host = document.createElement('div'); document.body.append(host); const root = createRoot(host);
  try {
    await act(async () => root.render(<ProjectUpdates project={{}} tasks={[]} updates={[{ id: 'u', version: '1', summary: 'Existing history' }]} store={store} canPost />));
    expect(document.body.textContent).toContain('Existing history');
    expect(button('Edit update').disabled).toBe(true);
    expect(button('Delete update').disabled).toBe(false);
    expect(document.body.textContent).toContain('a site owner needs to open the latest tracker');
    expect(document.querySelector('[aria-label="Your update"]')).not.toBeNull();
  } finally { await act(async () => root.unmount()); host.remove(); }
});
