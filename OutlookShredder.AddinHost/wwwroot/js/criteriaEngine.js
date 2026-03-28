/**
 * criteriaEngine.js
 * Orchestrates the full RFQ processing flow:
 *   1. Scan subject + body + attachment names for [XXXXXX] job reference.
 *   2. Fork: attachment path or body path.
 *   3. POST to C# proxy (/api/extract) which calls Claude + writes SharePoint.
 */

import { findJobRefs }                                        from './jobRef.js';
import { hasProcessableAttachment, buildAttachmentPayloads } from './attachmentReader.js';
import { callProxy }                                         from './graphClient.js';

// ── Office.js helpers ────────────────────────────────────────────────────────

function readSubject(item) {
  return new Promise(r =>
    typeof item.subject === 'string' ? r(item.subject)
    : item.subject.getAsync(res => r(res.value || '')));
}

function readFrom(item) {
  return new Promise(r => {
    try { r(item.from?.emailAddress || item.sender?.emailAddress || ''); }
    catch { r(''); }
  });
}

function readBody(item) {
  return new Promise((resolve, reject) =>
    item.body.getAsync(Office.CoercionType.Text, {}, r =>
      r.status === Office.AsyncResultStatus.Succeeded
        ? resolve(r.value || '')
        : reject(new Error(r.error?.message || 'Body read failed'))));
}

// ── Main ─────────────────────────────────────────────────────────────────────

/**
 * processItem(item)
 * Returns { matched, jobRefs, path, results, message }
 */
export async function processItem(item) {
  // 1. Read email
  let subject = '', body = '';
  try { [subject, body] = await Promise.all([readSubject(item), readBody(item)]); }
  catch (e) { return { matched: false, error: 'Could not read email: ' + e.message }; }

  const from       = await readFrom(item);
  const receivedAt = item.dateTimeCreated?.toISOString() ?? new Date().toISOString();

  // 2. GATE — scan subject + body + attachment filenames
  const attNames = (item.attachments || []).map(a => a.name || '').join(' ');
  const jobRefs  = findJobRefs(`${subject} ${body} ${attNames}`);

  if (!jobRefs.length) {
    return { matched: false, jobRefs: [], path: null, results: [],
             message: 'No [XXXXXX] job reference found in subject, body, or attachments.' };
  }

  const emailMeta = { emailSubject: subject, emailFrom: from, receivedAt,
                      hasAttachment: false, jobRefs };

  // 3. FORK
  const withAtts = hasProcessableAttachment(item);
  emailMeta.hasAttachment = withAtts;

  let results = [];
  let path    = 'body';

  if (withAtts) {
    path = 'attachment';
    let payloads = [];
    try { payloads = await buildAttachmentPayloads(item); }
    catch (e) { return { matched: true, jobRefs, path, results: [],
                         message: 'Attachment read failed: ' + e.message }; }

    if (!payloads.length) {
      path = 'body'; // all inline images — fall through to body
    } else {
      for (const payload of payloads) {
        try {
          const res = await callProxy('extract', {
            content:     '',                         // attachment path — content in base64Data
            sourceType:  'attachment',
            fileName:    payload.name,
            base64Data:  payload.content,            // null if server must fetch via EWS
            contentType: payload.contentType,
            jobRefs,
            bodyContext: body.substring(0, 2000),
            emailBody:   body,
            ...emailMeta,
            // EWS fallback fields (proxy uses these if base64Data is null)
            ewsToken: payload.ewsToken,
            ewsUrl:   payload.ewsUrl,
            itemId:   payload.itemId,
            attachId: payload.id,
          });
          results.push({ source: 'attachment', file: payload.name, ...res });
        } catch (e) {
          results.push({ source: 'attachment', file: payload.name,
                         success: false, error: e.message });
        }
      }
    }
  }

  // Body path (or fallback)
  if (path === 'body') {
    try {
      const res = await callProxy('extract', {
        content:    body.substring(0, 12000),
        sourceType: 'body',
        emailBody:  body,
        jobRefs,
        ...emailMeta,
      });
      results.push({ source: 'body', ...res });
    } catch (e) {
      return { matched: true, jobRefs, path, results: [],
               message: 'Body extraction failed: ' + e.message };
    }
  }

  const saved  = results.filter(r => r.success).length;
  const errors = results.filter(r => !r.success).length;

  return {
    matched: true,
    jobRefs,
    path,
    results,
    message: `${results.length} source(s) processed. ` +
             `${results.flatMap(r => r.rows || []).filter(r => r.success).length} product rows saved.` +
             (errors ? ` ${errors} error(s).` : ''),
  };
}
