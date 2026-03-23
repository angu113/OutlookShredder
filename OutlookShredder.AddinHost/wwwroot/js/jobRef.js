/**
 * jobRef.js
 * Detects [XXXXXX] job reference patterns — exactly 6 alphanumeric characters.
 */

export const JOB_REF_REGEX = /\[[A-Z0-9]{6}\]/gi;

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
