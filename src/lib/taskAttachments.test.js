import { describe, expect, it } from 'vitest';
import { MAX_ATTACHMENT_BYTES, validateAttachment } from './taskAttachments';
import { validateReference } from './referenceDocuments';

describe('50 MB upload boundary', () => {
  it.each([1, MAX_ATTACHMENT_BYTES])('accepts %i bytes for task and reference uploads', (size) => {
    const file = { name: 'report.pdf', size };
    expect(() => validateAttachment(file)).not.toThrow();
    expect(validateReference([], { name: file.name }, file)).toMatchObject({ size, fileName: file.name });
  });
  it.each([0, -1, MAX_ATTACHMENT_BYTES + 1, Infinity, NaN])('rejects invalid size %s for both upload paths', (size) => {
    const file = { name: 'report.pdf', size };
    expect(() => validateAttachment(file)).toThrow('50 MB');
    expect(() => validateReference([], { name: file.name }, file)).toThrow('50 MB');
  });
});
