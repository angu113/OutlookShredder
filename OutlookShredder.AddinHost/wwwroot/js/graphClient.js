/**
 * graphClient.js
 * Thin wrapper around the Microsoft Graph API using the MSAL token
 * acquired by Office.auth.getAccessToken().
 *
 * This handles SharePoint reads (for the dashboard) only.
 * SharePoint WRITES go through the C# proxy to keep credentials server-side.
 */

import { CONFIG } from './config.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';
// Derive the Graph host:path selector from the configured SharePoint site URL.
const _spUrl  = new URL(CONFIG.SP_SITE_URL);
const SP_SITE = `${_spUrl.host}:${_spUrl.pathname}`;   // e.g. 'contoso.sharepoint.com:/sites/mysite'
const SP_LIST = CONFIG.SP_LIST_NAME;

let _accessToken = null;

/** Get an access token from Office SSO (promise-based, requires require-resouce-access in manifest). */
async function getToken() {
  if (_accessToken) return _accessToken;
  try {
    _accessToken = await Office.auth.getAccessToken({ allowSignInPrompt: true });
    return _accessToken;
  } catch (e) {
    console.warn('[graphClient] Office SSO failed, dashboard reads will be unauthenticated:', e.message);
    return null;
  }
}

/** Call the C# proxy (which holds the Anthropic API key server-side). */
export async function callProxy(endpoint, body) {
  const res = await fetch(`${CONFIG.PROXY_URL}/api/${endpoint}`, {
    method:  'POST',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Proxy ${endpoint} returned ${res.status}`);
  return res.json();
}

/** Read items from the RFQLineItems SharePoint list (for dashboard). */
export async function readSpItems(top = 200, filter = '') {
  const token = await getToken();
  const headers = token ? { Authorization: `Bearer ${token}` } : {};

  const params = new URLSearchParams({ '$top': top, '$expand': 'fields', '$orderby': 'fields/ReceivedAt desc' });
  if (filter) params.set('$filter', filter);

  const siteRes = await fetch(`${GRAPH}/sites/${SP_SITE}`, { headers });
  if (!siteRes.ok) throw new Error(`Graph site lookup failed (${siteRes.status})`);
  const siteData = await siteRes.json();
  const siteId   = siteData.id;

  const listRes = await fetch(
    `${GRAPH}/sites/${siteId}/lists?$filter=displayName eq '${SP_LIST}'`, { headers });
  if (!listRes.ok) throw new Error(`Graph list lookup failed (${listRes.status})`);
  const listData = await listRes.json();
  const listId   = listData.value?.[0]?.id;
  if (!listId) throw new Error(`List "${SP_LIST}" not found`);

  const itemsRes = await fetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items?${params}`, { headers });
  if (!itemsRes.ok) throw new Error(`Graph items fetch failed (${itemsRes.status})`);
  const itemsData = await itemsRes.json();
  return itemsData.value || [];
}
