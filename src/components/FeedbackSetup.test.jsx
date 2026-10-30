import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, it, vi } from 'vitest';
import { FeedbackSetup } from './FeedbackSetup';
globalThis.IS_REACT_ACT_ENVIRONMENT = true;
it('shows the missing list, lets the owner choose groups and provisions only on click', async () => {
  const setup = { options: vi.fn().mockResolvedValue({ owners: { Title: 'Site Owners' }, groups: [{ Id: 20, Title: 'Metrology App User' }, { Id: 21, Title: 'Managers' }], selected: [20] }), provision: vi.fn().mockResolvedValue({ ready: true }) };
  const ready = vi.fn(), busy = vi.fn(), node = document.createElement('div'), root = createRoot(node);
  try {
    await act(async () => root.render(<FeedbackSetup setup={setup} state={{ title: 'ModernizationFeedback', ready: false }} onBusy={busy} onReady={ready} />));
    expect(node.textContent).toContain('has not been created'); expect(setup.provision).not.toHaveBeenCalled();
    await act(async () => node.querySelectorAll('input')[1].click());
    await act(async () => node.querySelector('button').click());
    expect(setup.provision).toHaveBeenCalledWith([20, 21], expect.any(Function)); expect(ready).toHaveBeenCalledWith({ ready: true });
  } finally { await act(async () => root.unmount()); }
});
