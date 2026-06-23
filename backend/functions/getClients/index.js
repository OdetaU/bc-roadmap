const { getSharePointItems, getMicrosoftClients } = require('../shared/graph');

module.exports = async function (context, req) {
  try {
    // Get clients from Partner Center / BC Admin API
    const apiClients = await getMicrosoftClients();

    // Get extra info stored in SharePoint (contacts, notes etc.)
    const spClients = await getSharePointItems('BC_Clients');

    // Merge: API data is base, SharePoint adds manual fields
    const merged = apiClients.map(c => {
      const sp = spClients.find(s => s.TenantId === c.tenantId) || {};
      return {
        id: c.tenantId,
        name: c.displayName,
        tenantId: c.tenantId,
        bcVersion: c.bcVersion || sp.BcVersion || '',
        contact: sp.Contact || '',
        email: sp.Email || '',
        notes: sp.Notes || '',
        spItemId: sp.id || null,
      };
    });

    context.res = { status: 200, body: merged };
  } catch (err) {
    context.log.error('getClients error:', err);
    context.res = { status: 500, body: { error: err.message } };
  }
};
