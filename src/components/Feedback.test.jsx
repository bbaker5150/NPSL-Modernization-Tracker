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
  const cancelNativeSubmit = event => { if (event.target.closest('.feedback-actions button')) event.preventDefault(); };
  document.addEventListener('click', cancelNativeSubmit, true);
  try {
    await act(async () => root.render(<Feedback store={store} />));
    await act(async () => document.querySelector('button').click());
    expect(document.body.textContent).not.toContain('Owner inbox');
    await act(async () => { change(document.querySelector('input[maxlength]'), 'Broken button'); change(document.querySelector('textarea'), 'Clicking it does nothing'); });
    await act(async () => [...document.querySelectorAll('button')].find(button => button.textContent === 'Send feedback').click());
    expect(document.querySelector('[role="alert"]').textContent).toContain('Denied');
    expect(document.querySelector('textarea').value).toBe('Clicking it does nothing');
    store.post.mockResolvedValue({});
    await act(async () => [...document.querySelectorAll('button')].find(button => button.textContent === 'Send feedback').click());
    expect(document.querySelector('[role="status"]').textContent).toContain('Thank you');
  } finally { document.removeEventListener('click', cancelNativeSubmit, true); await act(async () => root.unmount()); node.remove(); }
});
it('shows the private inbox for owners and renders reports as text', async () => {
  const store = { webUrl: 'https://example.test', currentUser: async () => ({ isSiteAdmin: true }), get: vi.fn().mockImplementation(async path => path.includes('/lists?') ? { value: [{ Id: '11111111-1111-1111-1111-111111111111', Title: 'Modernization-Tracker - Feedback', BaseTemplate: 100 }] } : { Id: 7, Title: 'Modernization-Tracker - Feedback', HasUniqueRoleAssignments: true, value: [{ PrincipalId: 7, RoleDefinitionBindings: [], InternalName: 'FeedbackKind' }, { PrincipalId: 7, RoleDefinitionBindings: [], InternalName: 'FeedbackDetails' }, { PrincipalId: 7, RoleDefinitionBindings: [], Id: 1, Title: '<script>bad</script>', FeedbackKind: 'Bug', FeedbackDetails: '<img src=x>', Created: '2026-10-30T00:00:00Z', Author: { Title: 'Reporter' } }] }) };
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
    store.get.mockImplementation(async path => path.includes('/lists?') ? { value: [{ Id: '11111111-1111-1111-1111-111111111111', Title: 'Modernization-Tracker - Feedback', BaseTemplate: 100 }] } : { Low: '4294967295', High: '2147483647', Id: 7, Title: 'Modernization-Tracker - Feedback', HasUniqueRoleAssignments: true, value: [{ PrincipalId: 7, RoleDefinitionBindings: [], InternalName: 'FeedbackKind' }, { PrincipalId: 7, RoleDefinitionBindings: [], InternalName: 'FeedbackDetails' }] });
    await act(async () => node.querySelector('button').click());
    expect(document.querySelector('footer').textContent).toContain('Owner inbox');
    expect(document.querySelector('[role="alert"]')).toBeNull();
    await act(async () => document.querySelector('[aria-label="Close feedback"]').click());
    store.get.mockResolvedValue({ Low: '0', High: '0' });
    await act(async () => node.querySelector('button').click());
    expect(document.body.textContent).not.toContain('Owner inbox');
  } finally { await act(async () => root.unmount()); node.remove(); }
});

it('automatically offers setup to an owner when SharePoint storage is missing', async () => {
  const store = { webUrl: 'https://example.test', scopedAccess: true, currentUser: async () => ({ id: 8, isSiteAdmin: true }), post: vi.fn(), get: vi.fn(async path => {
    if (path.includes('/associatedownergroup')) return { Id: 10, Title: 'Site Owners' };
    if (path.includes('/lists?')) return { value: [] };
    if (path.includes('/sitegroups?')) return { value: [{ Id: 20, Title: 'Metrology App User' }] };
    throw Object.assign(new Error('List not found'), { status: 404 });
  }) };
  const node = document.createElement('div'); document.body.append(node); const root = createRoot(node);
  try {
    await act(async () => root.render(<Feedback store={store} />));
    await act(async () => node.querySelector('button').click());
    expect(document.querySelector('[aria-label="Set up private feedback"]').textContent).toContain('Modernization-Tracker - Feedback has not been created');
    expect([...document.querySelectorAll('button')].find(button => button.textContent === 'Create private feedback list')).toBeDefined();
    expect(document.querySelector('footer button').disabled).toBe(true);
    expect(store.post).not.toHaveBeenCalled();
  } finally { await act(async () => root.unmount()); node.remove(); }
});

it('shows validation feedback when Send is clicked with empty fields', async () => {
  const store = { currentUser: async () => ({ isSiteAdmin: true }) };
  const node = document.createElement('div'); document.body.append(node); const root = createRoot(node);
  try {
    await act(async () => root.render(<Feedback store={store} />));
    await act(async () => node.querySelector('button').click());
    const send = [...document.querySelectorAll('button')].find(button => button.textContent === 'Send feedback');
    expect(send.disabled).toBe(false);
    await act(async () => send.click());
    expect(document.querySelector('[role="alert"]').textContent).toContain('Enter a summary');
  } finally { await act(async () => root.unmount()); node.remove(); }
});

it('hides the inbox navigation inside the inbox and confirms or cancels clearing', async () => {
  const store = { currentUser: async () => ({ isSiteAdmin: true }) };
  const node = document.createElement('div'); document.body.append(node); const root = createRoot(node);
  const button = label => [...document.querySelectorAll('button')].find(button => button.textContent === label);
  try {
    await act(async () => root.render(<Feedback store={store} />));
    await act(async () => node.querySelector('button').click());
    await act(async () => { change(document.querySelector('input[maxlength]'), 'Review me'); change(document.querySelector('textarea'), 'Details'); });
    await act(async () => button('Send feedback').click());
    await act(async () => button('Owner inbox').click());
    expect(button('Owner inbox')).toBeUndefined();
    expect(document.querySelector('#feedback-title').textContent).toBe('Owner inbox');
    await act(async () => button('Clear inbox').click());
    expect(document.querySelector('.feedback-clear-confirm').textContent).toContain('all site owners');
    await act(async () => button('Cancel').click());
    expect(document.querySelector('.feedback-item').textContent).toContain('Review me');
    await act(async () => button('Clear inbox').click());
    await act(async () => button('Confirm clear').click());
    expect(document.querySelector('.feedback-item')).toBeNull();
    expect(document.body.textContent).toContain('1 submission cleared');
    expect(button('Clear inbox').disabled).toBe(true);
    await act(async () => button('Back to feedback').click());
    expect(button('Owner inbox')).toBeDefined();
  } finally { await act(async () => root.unmount()); node.remove(); }
});

it('edits, cancels and deletes individual reports while preserving other submissions under host click cancellation', async () => {
  const store = { currentUser: async () => ({ isSiteAdmin: true }) };
  const node = document.createElement('div'); document.body.append(node); const root = createRoot(node);
  const button = (label, scope = document) => [...scope.querySelectorAll('button')].find(button => button.textContent === label);
  const cancelNative = event => { if (event.target.closest('.feedback-dialog button')) event.preventDefault(); };
  document.addEventListener('click', cancelNative, true);
  try {
    await act(async () => root.render(<Feedback store={store} />));
    await act(async () => node.querySelector('button').click());
    for (const title of ['Keep me', 'Edit me']) {
      await act(async () => { change(document.querySelector('input[maxlength]'), title); change(document.querySelector('textarea'), 'Original details'); });
      await act(async () => button('Send feedback').click());
      if (title === 'Keep me') await act(async () => button('Send another').click());
    }
    await act(async () => button('Owner inbox').click());
    const selected = () => document.querySelectorAll('article')[0];
    await act(async () => button('Edit', selected()).click());
    await act(async () => change(document.querySelector('input[maxlength]'), 'Discard this'));
    await act(async () => button('Cancel edit').click());
    expect(selected().querySelector('h3').textContent).toBe('Edit me');
    await act(async () => button('Edit', selected()).click());
    await act(async () => {
      change(document.querySelector('input[maxlength]'), 'Updated report');
      change(document.querySelector('textarea'), 'Updated details');
      const status = document.querySelectorAll('.feedback-edit select')[1];
      status.value = 'Resolved'; status.dispatchEvent(new Event('change', { bubbles: true }));
    });
    await act(async () => button('Save changes').click());
    expect(selected().textContent).toContain('Updated report');
    expect(selected().textContent).toContain('Updated details');
    expect(selected().querySelector('.feedback-status').textContent).toBe('Resolved');
    expect(document.querySelectorAll('article')[1].textContent).toContain('Keep me');
    await act(async () => button('Delete', selected()).click());
    expect(document.querySelector('.feedback-clear-confirm').textContent).toContain('Updated report');
    await act(async () => button('Cancel').click());
    expect(document.querySelectorAll('article')).toHaveLength(2);
    await act(async () => button('Delete', selected()).click());
    await act(async () => button('Confirm delete').click());
    expect(document.querySelectorAll('article')).toHaveLength(1);
    expect(document.querySelector('article').textContent).toContain('Keep me');
    expect(document.body.textContent).toContain('Submission deleted.');
  } finally { document.removeEventListener('click', cancelNative, true); await act(async () => root.unmount()); node.remove(); }
});

it('keeps the owner edit draft and original record on a failed save', async () => {
  const store = { webUrl: 'https://example.test', currentUser: async () => ({ isSiteAdmin: true }), post: vi.fn().mockRejectedValue(new Error('Access denied')), get: vi.fn(async path => {
    if (path.includes('/lists?')) return { value: [{ Id: '11111111-1111-1111-1111-111111111111', Title: 'Modernization-Tracker - Feedback', BaseTemplate: 100 }] };
    if (path.includes('/fields?$filter')) return { value: [{ InternalName: 'FeedbackStatus', TypeAsString: 'Choice', Choices: ['New', 'In progress', 'Resolved', 'Closed'] }] };
    if (path.includes('/items?')) return { value: [{ Id: 1, Title: 'Original report', FeedbackKind: 'Bug', FeedbackDetails: 'Details', 'odata.etag': '"2"', Created: '2026-10-30T00:00:00Z' }] };
    return { Id: 7, Title: 'Modernization-Tracker - Feedback', HasUniqueRoleAssignments: true, value: [{ PrincipalId: 7, RoleDefinitionBindings: [], InternalName: 'FeedbackKind' }, { PrincipalId: 7, RoleDefinitionBindings: [], InternalName: 'FeedbackDetails' }] };
  }) };
  const node = document.createElement('div'); document.body.append(node); const root = createRoot(node);
  const button = label => [...document.querySelectorAll('button')].find(button => button.textContent === label);
  try {
    await act(async () => root.render(<Feedback store={store} />));
    await act(async () => node.querySelector('button').click());
    await act(async () => button('Owner inbox').click());
    await act(async () => button('Edit').click());
    await act(async () => change(document.querySelector('input[maxlength]'), 'My draft'));
    await act(async () => button('Save changes').click());
    expect(document.querySelector('[role="alert"]').textContent).toContain('Access denied');
    expect(document.querySelector('input[maxlength]').value).toBe('My draft');
    expect(document.querySelector('article h3').textContent).toBe('Original report');
    expect(store.post).toHaveBeenCalledTimes(1);
  } finally { await act(async () => root.unmount()); node.remove(); }
});
