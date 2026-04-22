/**
 * jobRef.js
 * Detects bracketed job-reference patterns:
 *   Initials: [IIXXXXX]  — 2-letter initials + 5 Crockford Base32 digits (e.g. [AW00001])
 *   HQ:       [HQXXXXXX] — "HQ" prefix + 6 alphanumeric chars
 *   Legacy:   [XXXXXX]   — exactly 6 alphanumeric chars
 * Initials format listed first; HQ before legacy so 8-char matches aren't truncated.
 */

export const JOB_REF_REGEX = /\[(?:[A-Z]{2}[0-9A-HJKMNP-TV-Z]{5}|HQ[A-Z0-9]{6}|[A-Z0-9]{6})\]/gi;

/** Return all unique job references found in text, upper-cased. */
export function findJobRefs(text = '') {
  const matches = text.match(JOB_REF_REGEX) || [];
  return [...new Set(matches.map(m => m.toUpperCase()))];
}

/** Returns true if text contains at least one valid job reference. */
export function hasJobRef(text = '') {
  JOB_REF_REGEX.lastIndex = 0;
  return JOB_REF_REGEX.test(text);
}
