/**
 * attachmentReader.js
 * Reads attachment content from an Outlook item via Office.js.
 * Tries getAttachmentContentAsync() first (Office 1.8+).
 * Falls back to passing the EWS token + attachment ID to the proxy.
 */

const SUPPORTED_TYPES = new Set([
  'application/pdf',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/msword', 'application/vnd.ms-excel',
  'text/plain', 'text/csv', 'text/html', 'application/rtf', 'text/rtf',
  'application/octet-stream',
]);

export function hasProcessableAttachment(item) {
  return (item.attachments || []).some(a => !a.isInline && a.size > 0);
}

/** Try to get base64 content directly from Office.js (requires 1.8+). */
function tryGetContentDirect(attachmentId) {
  return new Promise(resolve => {
    if (!Office.context.mailbox.item.getAttachmentContentAsync) return resolve(null);
    Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, {}, result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) return resolve(null);
      const { content, format } = result.value;
      resolve(format === Office.MailboxEnums.AttachmentContentFormat.Base64 ? content : null);
    });
  });
}

/** Get EWS callback token (for server-side fetch fallback). */
function getEwsToken() {
  return new Promise(resolve => {
    try {
      Office.context.mailbox.getCallbackTokenAsync({}, r =>
        resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null));
    } catch { resolve(null); }
  });
}

export async function buildAttachmentPayloads(item) {
  const attachments = (item.attachments || []).filter(a => {
    if (a.isInline || a.size === 0) return false;
    const ct = (a.contentType || '').toLowerCase().split(';')[0].trim();
    return SUPPORTED_TYPES.has(ct) || ct.startsWith('text/');
  });

  if (!attachments.length) return [];

  const ewsToken = await getEwsToken();

  return Promise.all(attachments.map(async a => ({
    id:          a.id,
    name:        a.name,
    contentType: (a.contentType || 'application/octet-stream').toLowerCase(),
    size:        a.size,
    content:     await tryGetContentDirect(a.id),
    ewsToken,
    ewsUrl:      Office.context.mailbox.ewsUrl || null,
    itemId:      item.itemId || null,
  })));
}
