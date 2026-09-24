// Display only: preserve directory identities and source names in storage.
export function displayName(value) {
  const name = String(value || '').trim();
  if (!name.includes(',')) return name;
  const [last, given] = name.split(',', 2);
  const personal = given.replace(/\s*\([^)]*\)/g, '').split(/\s+(?:CIV|CTR|MIL|USN|USMC|USA|USAF|USCG|FLTREADCEN)\b/i)[0].trim();
  return personal ? `${personal} ${last.trim()}` : name;
}
