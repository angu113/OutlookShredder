/**
 * graphClient.js
 * Client-side API helpers for the OutlookShredder add-in and dashboard.
 *
 * All SharePoint reads and writes are routed through the C# proxy so that
 * credentials never need to be present in the browser.  The proxy uses
 * its own app-only (client credential) auth against Microsoft Graph.
 */

import { CONFIG } from './config.js';

/** POST to a proxy endpoint (used by criteriaEngine to trigger extraction). */
export async function callProxy(endpoint, body) {
  const res = await fetch(`${CONFIG.PROXY_URL}/api/${endpoint}`, {
    method:  'POST',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Proxy ${endpoint} returned ${res.status}`);
  return res.json();
}

/**
 * Read items from the RFQLineItems SharePoint list for the dashboard.
 * Calls GET /api/items on the proxy — no browser auth token required.
 * Returns an array of { fields: {...} } objects matching the SharePoint column names.
 */
export async function readSpItems(top = 500) {
  const res = await fetch(`${CONFIG.PROXY_URL}/api/items?top=${top}`);
  if (!res.ok) throw new Error(`Items fetch failed (${res.status})`);
  // Proxy returns a flat array of field dictionaries; wrap each so the dashboard
  // can access values as item.fields.ColumnName — consistent with the Graph API shape.
  const fields = await res.json();
  return fields.map(f => ({ fields: f }));
}
