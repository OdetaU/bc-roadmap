const { getMicrosoftClients } = require('../shared/graph');

module.exports = async function (context, req) {
  try {
    const clients = await getMicrosoftClients();
    // Return flat list of all environments with versions
    const envs = [];
    for (const c of clients) {
      for (const env of c.environments || []) {
        envs.push({
          tenantId: c.tenantId,
          clientName: c.displayName,
          envName: env.name,
          envType: env.type,
          version: env.version,
          status: env.status,
          updateWindow: env.updateWindow,
        });
      }
    }
    context.res = { status: 200, body: envs };
  } catch (err) {
    context.log.error('getVersions error:', err);
    context.res = { status: 500, body: { error: err.message } };
  }
};
