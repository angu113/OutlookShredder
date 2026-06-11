/**
 * jobRef.js
 * Detects bracketed job-reference patterns:
 *   [IIXXXX] — 2-letter user initials + 4 Crockford Base32 chars, 6 chars total (e.g. [AW0001])
 * The first two characters are always letters, which excludes a supplier's own job number
 * such as "J06601" (single letter then digits). The retired generic [A-Z0-9]{6} rule used to
 * mis-capture those; HQ+6 is likewise retired.
 */

export const JOB_REF_REGEX = /\[[A-Z]{2}[A-Z0-9]{4}\]/gi;

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
