import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, it, vi } from 'vitest';
import { Feedback } from './Feedback';
globalThis.IS_REACT_ACT_ENVIRONMENT = true;
HTMLDialogElement.prototype.showModal = function () { this.open = true; };
HTMLDialogElement.prototype.close = function () { this.open = false; };
const change = (node, value) => { Object.getOwnPropertyDescriptor(node.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype, 'value').set.call(node, value); node.dispatchEvent(new Event('input', { bubbles: true })); };
it('offers feedback to a viewer, hides inbox, and preserves a failed draft', async () => {
  const store = { webUrl: 'https://example.test', currentUser: async () => ({}), get: async () => ({ value: [] }), post: vi.fn().mockRejectedValue(new Error('Denied')) };
  const node = document.createElement('div'); document.body.append(node); const root = createRoot(node);
  try {
    await act(async () => root.render(<Feedback store={store} />));
    await act(async () => document.querySelector('button').click());
    expect(document.body.textContent).not.toContain('Owner inbox');
    await act(async () => { change(document.querySelector('input[maxlength]'), 'Broken button'); change(document.querySelector('textarea'), 'Clicking it does nothing'); });
    await act(async () => document.querySelector('form').dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })));
    expect(document.querySelector('[role="alert"]').textContent).toContain('Denied');
    expect(document.querySelector('textarea').value).toBe('Clicking it does nothing');
    store.post.mockResolvedValue({});
    await act(async () => document.querySelector('form').dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })));
    expect(document.querySelector('[role="status"]').textContent).toContain('Thank you');
  } finally { await act(async () => root.unmount()); node.remove(); }
});
it('shows the private inbox for owners and renders reports as text', async () => {
  const store = { webUrl: 'https://example.test', currentUser: async () => ({ isSiteAdmin: true }), get: vi.fn().mockResolvedValue({ value: [{ Id: 1, Title: '<script>bad</script>', FeedbackKind: 'Bug', FeedbackDetails: '<img src=x>', Created: '2026-10-30T00:00:00Z', Author: { Title: 'Reporter' } }] }) };
  const node = document.createElement('div'); document.body.append(node); const root = createRoot(node);
  try {
    await act(async () => root.render(<Feedback store={store} />));
    expect(store.get).not.toHaveBeenCalled();
    await act(async () => document.querySelector('button').click());
    const send = [...document.querySelectorAll('button')].filter(button => button.textContent === 'Send feedback');
    expect(send).toHaveLength(1);
    expect(send[0].closest('footer')).not.toBeNull();
    expect(send[0].previousElementSibling.textContent).toBe('Owner inbox');
    expect(document.querySelector('header').textContent).not.toContain('Owner inbox');
    await act(async () => [...document.querySelectorAll('button')].find(button => button.textContent === 'Owner inbox').click());
    expect(document.body.textContent).toContain('<script>bad</script>'); expect(document.querySelector('script,img')).toBeNull();
  } finally { await act(async () => root.unmount()); node.remove(); }
});

it('rechecks signed-in owner access on reopen and explains verification failures', async () => {
  const store = { webUrl: 'https://example.test', currentUser: async () => ({}), get: vi.fn().mockRejectedValue(new Error('Network failure')) };
  const node = document.createElement('div'); document.body.append(node); const root = createRoot(node);
  try {
    await act(async () => root.render(<Feedback store={store} />));
    await act(async () => node.querySelector('button').click());
    expect(document.querySelector('[role="alert"]').textContent).toContain('could not be verified');
    expect(document.body.textContent).not.toContain('Owner inbox');
    await act(async () => document.querySelector('[aria-label="Close feedback"]').click());
    store.get.mockResolvedValue({ Low: '4294967295', High: '2147483647' });
    await act(async () => node.querySelector('button').click());
    expect(document.querySelector('footer').textContent).toContain('Owner inbox');
    expect(document.querySelector('[role="alert"]')).toBeNull();
    await act(async () => document.querySelector('[aria-label="Close feedback"]').click());
    store.get.mockResolvedValue({ Low: '0', High: '0' });
    await act(async () => node.querySelector('button').click());
    expect(document.body.textContent).not.toContain('Owner inbox');
  } finally { await act(async () => root.unmount()); node.remove(); }
});
