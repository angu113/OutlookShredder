/**
 * jobRef.js
 * Detects bracketed job-reference patterns:
 *   New:    [HQXXXXXX]  — "HQ" prefix + 6 alphanumeric chars
 *   Legacy: [XXXXXX]    — exactly 6 alphanumeric chars
 * The HQ-prefixed alt is listed first so 8-char matches aren't truncated to 6.
 */

export const JOB_REF_REGEX = /\[(?:HQ[A-Z0-9]{6}|[A-Z0-9]{6})\]/gi;

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
