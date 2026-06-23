const { getSharePointItems, getAppRegistrationSecrets, getAdminRelationships } = require('../shared/graph');

module.exports = async function (context, req) {
  try {
    // Auto: App Registration secrets from Graph
    const appSecrets = await getAppRegistrationSecrets();

    // Auto: GDAP relationships
    const gdap = await getAdminRelationships();

    // Manual: extra entries from SharePoint
    const spSecrets = await getSharePointItems('BC_Secrets');

    const results = [];

    // Map app registration secrets
    for (const s of appSecrets) {
      const sp = spSecrets.find(x => x.SecretId === s.secretId) || {};
      results.push({
        id: s.secretId,
        spItemId: sp.id || null,
        source: 'appreg',
        type: 'appreg',
        clientId: sp.ClientId || '',
        name: s.appName + (s.secretName ? ` — ${s.secretName}` : ''),
        expires: s.expiresAt ? s.expiresAt.slice(0, 10) : '',
        warn: sp.WarnDays || 30,
        notes: sp.Notes || '',
      });
    }

    // Map GDAP relationships
    for (const r of gdap) {
      const sp = spSecrets.find(x => x.RelationshipId === r.id) || {};
      results.push({
        id: r.id,
        spItemId: sp.id || null,
        source: 'gdap',
        type: 'admin',
        clientId: r.customerId || sp.ClientId || '',
        name: r.displayName || r.customerName || r.id,
        expires: r.expiresAt ? r.expiresAt.slice(0, 10) : '',
        status: r.status,
        warn: sp.WarnDays || 60,
        notes: sp.Notes || '',
      });
    }

    // Manual-only SP entries (no auto source)
    for (const sp of spSecrets) {
      if (sp.Source === 'manual') {
        results.push({
          id: 'sp-' + sp.id,
          spItemId: sp.id,
          source: 'manual',
          type: sp.Type || 'appreg',
          clientId: sp.ClientId || '',
          name: sp.Name || '',
          expires: sp.ExpiresAt || '',
          warn: sp.WarnDays || 30,
          notes: sp.Notes || '',
        });
      }
    }

    context.res = { status: 200, body: results };
  } catch (err) {
    context.log.error('getSecrets error:', err);
    context.res = { status: 500, body: { error: err.message } };
  }
};
