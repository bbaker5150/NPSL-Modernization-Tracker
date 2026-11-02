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
