// config.js
// All runtime configuration in one place.
// Update PROXY_URL before deploying to production.

export const CONFIG = {
    // URL of OutlookShredder.Proxy — the C# API that holds the Anthropic key
    // Development: IIS Express on port 7001
    // Production:  your Azure App Service URL
    PROXY_URL: 'https://localhost:7001',

    // SharePoint
    SP_SITE_URL:   'https://metalsupermarkets.sharepoint.com/sites/hackensack',
    SP_LIST_NAME:  'RFQLineItems',

    // Azure AD app registration (used for Graph API — SharePoint writes)
    // The Proxy uses its OWN client credentials for Claude.
    // The add-in uses MSAL interactive / SSO for Graph API.
    AAD_CLIENT_ID: 'REPLACE_WITH_YOUR_CLIENT_ID',
    AAD_TENANT_ID: 'REPLACE_WITH_YOUR_TENANT_ID',

    // Graph API scopes needed by the add-in
    GRAPH_SCOPES: ['https://graph.microsoft.com/Sites.ReadWrite.All'],

    // Max body characters sent to the proxy for extraction
    MAX_BODY_LENGTH: 12000,
};
