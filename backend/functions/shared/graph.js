const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const fetch = require('node-fetch');

const TENANT_ID     = process.env.TENANT_ID;
const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SP_SITE_ID    = process.env.SP_SITE_ID;   // SharePoint site id
const SP_SITE_URL   = process.env.SP_SITE_URL;   // e.g. https://uabstrongpoint.sharepoint.com/sites/BCRoadmap

// ── Auth ──────────────────────────────────────────────────────────────────────
function getCredential() {
  return new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
}

function getGraphClient() {
  const credential = getCredential();
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default'],
  });
  return Client.initWithMiddleware({ authProvider });
}

async function getBearerToken(scope) {
  const credential = getCredential();
  const token = await credential.getToken(scope);
  return token.token;
}

// ── SharePoint ─────────────────────────────────────────────────────────────────
async function getSharePointItems(listName) {
  const client = getGraphClient();
  const res = await client
    .api(`/sites/${SP_SITE_ID}/lists/${listName}/items?expand=fields&$top=999`)
    .get();
  return (res.value || []).map(i => ({ id: i.id, ...i.fields }));
}

async function createSharePointItem(listName, fields) {
  const client = getGraphClient();
  const res = await client
    .api(`/sites/${SP_SITE_ID}/lists/${listName}/items`)
    .post({ fields });
  return res;
}

async function updateSharePointItem(listName, itemId, fields) {
  const client = getGraphClient();
  await client
    .api(`/sites/${SP_SITE_ID}/lists/${listName}/items/${itemId}/fields`)
    .patch(fields);
}

async function deleteSharePointItem(listName, itemId) {
  const client = getGraphClient();
  await client
    .api(`/sites/${SP_SITE_ID}/lists/${listName}/items/${itemId}`)
    .delete();
}

// ── BC Admin API ───────────────────────────────────────────────────────────────
// Returns list of customers with their BC environments and versions
async function getMicrosoftClients() {
  try {
    const token = await getBearerToken('https://api.businesscentral.dynamics.com/.default');
    // Partner endpoint: list all customers
    const resp = await fetch('https://api.businesscentral.dynamics.com/admin/v2.21/applications/environments', {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!resp.ok) {
      console.warn('BC Admin API response:', resp.status, await resp.text());
      return [];
    }
    const data = await resp.json();
    // data.value is array of environments across all tenants (CSP/MCSP)
    // Group by tenant
    const tenantMap = {};
    for (const env of data.value || []) {
      const tid = env.aadTenantId;
      if (!tenantMap[tid]) {
        tenantMap[tid] = {
          tenantId: tid,
          displayName: env.applicationFamily || tid,
          environments: [],
        };
      }
      tenantMap[tid].environments.push({
        name: env.environmentName,
        type: env.environmentType, // Production / Sandbox
        version: env.applicationVersion,
        status: env.status,
        updateWindow: env.ringName,
      });
      // Use production version as "current BC version"
      if (env.environmentType === 'Production') {
        tenantMap[tid].bcVersion = env.applicationVersion;
      }
    }
    return Object.values(tenantMap);
  } catch (err) {
    console.error('getMicrosoftClients error:', err.message);
    return [];
  }
}

// ── App Registrations secrets ──────────────────────────────────────────────────
async function getAppRegistrationSecrets() {
  const client = getGraphClient();
  const apps = await client.api('/applications?$top=999&$select=id,displayName,passwordCredentials,appId').get();
  const results = [];
  for (const app of apps.value || []) {
    for (const cred of app.passwordCredentials || []) {
      results.push({
        appId: app.appId,
        appName: app.displayName,
        secretId: cred.keyId,
        secretName: cred.displayName || '(unnamed)',
        expiresAt: cred.endDateTime,
      });
    }
  }
  return results;
}

// ── GDAP / Admin Relationships ─────────────────────────────────────────────────
async function getAdminRelationships() {
  const client = getGraphClient();
  try {
    const res = await client
      .api('/tenantRelationships/delegatedAdminRelationships?$top=999')
      .get();
    return (res.value || []).map(r => ({
      id: r.id,
      displayName: r.displayName,
      customerId: r.customer?.tenantId,
      customerName: r.customer?.displayName,
      status: r.status,
      activatedAt: r.activatedDateTime,
      expiresAt: r.endDateTime,
    }));
  } catch (err) {
    console.warn('getAdminRelationships error:', err.message);
    return [];
  }
}

module.exports = {
  getSharePointItems,
  createSharePointItem,
  updateSharePointItem,
  deleteSharePointItem,
  getMicrosoftClients,
  getAppRegistrationSecrets,
  getAdminRelationships,
};
