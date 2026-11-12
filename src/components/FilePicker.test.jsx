import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, it, vi } from 'vitest';
import { FilePicker } from './FilePicker';

globalThis.IS_REACT_ACT_ENVIRONMENT = true;

it('keeps a labeled native input and forwards selected files so callers can reset and reselect them', async () => {
  const host = document.createElement('div'); document.body.append(host); const root = createRoot(host);
  const onChange = vi.fn(event => { expect(event.target.files[0].name).toBe('screenshot.png'); event.target.value = ''; });
  try {
    await act(async () => root.render(<FilePicker label="Screenshots and attachments" ariaLabel="Feedback attachments" hint="Optional · up to 50 MB each" accept="image/*,.pdf" onChange={onChange} />));
    const input = host.querySelector('input[type=file]');
    expect(input.getAttribute('aria-label')).toBe('Feedback attachments');
    expect(document.getElementById(input.getAttribute('aria-describedby')).textContent).toBe('Optional · up to 50 MB each');
    expect(input.labels[0].textContent).toContain('Screenshots and attachments');
    expect(input.multiple).toBe(true);
    expect(input.accept).toBe('image/*,.pdf');
    expect(input.tabIndex).toBe(0);
    Object.defineProperty(input, 'files', { value: [new File(['image'], 'screenshot.png', { type: 'image/png' })] });
    await act(async () => { input.focus(); input.dispatchEvent(new Event('change', { bubbles: true })); });
    expect(document.activeElement).toBe(input);
    await act(async () => input.dispatchEvent(new Event('change', { bubbles: true })));
    expect(onChange).toHaveBeenCalledTimes(2);
    await act(async () => root.render(<FilePicker label="Attach document" multiple={false} disabled onChange={onChange} />));
    expect(input.disabled).toBe(true);
    expect(input.multiple).toBe(false);
    expect(input.getAttribute('aria-label')).toBe('Attach document');
    expect(input.hasAttribute('aria-describedby')).toBe(false);
  } finally { await act(async () => root.unmount()); host.remove(); }
});
