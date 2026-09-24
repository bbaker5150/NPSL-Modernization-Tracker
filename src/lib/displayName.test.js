import { expect, it } from 'vitest';
import { displayName } from './displayName';
it('shortens standard SharePoint names without changing plain names', () => {
  expect(displayName('Nguyen, Anh N CIV USN FLTREADCEN SW SAN CA (USA)')).toBe('Anh N Nguyen');
  expect(displayName('Baker, Barry T CIV (USA)')).toBe('Barry T Baker');
  expect(displayName('Smith-Jones, Mary Ann CTR (USA)')).toBe('Mary Ann Smith-Jones');
  expect(displayName('Alex Smith')).toBe('Alex Smith');
  expect(displayName('')).toBe('');
});
