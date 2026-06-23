const { createSharePointItem, updateSharePointItem } = require('../shared/graph');

module.exports = async function (context, req) {
  try {
    const body = req.body || {};
    const fields = {
      ClientId:       body.clientId || '',
      Type:           body.type || 'appreg',
      Name:           body.name || '',
      ExpiresAt:      body.expires || null,
      WarnDays:       body.warn || 30,
      Notes:          body.notes || '',
      Source:         body.source || 'manual',
      SecretId:       body.secretId || '',
      RelationshipId: body.relationshipId || '',
    };

    if (body.spItemId) {
      await updateSharePointItem('BC_Secrets', body.spItemId, fields);
      context.res = { status: 200, body: { ok: true } };
    } else {
      const res = await createSharePointItem('BC_Secrets', fields);
      context.res = { status: 200, body: { ok: true, spItemId: res.id } };
    }
  } catch (err) {
    context.log.error('saveSecret error:', err);
    context.res = { status: 500, body: { error: err.message } };
  }
};
