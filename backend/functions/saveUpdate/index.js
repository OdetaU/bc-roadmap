const { createSharePointItem, updateSharePointItem } = require('../shared/graph');

module.exports = async function (context, req) {
  try {
    const body = req.body || {};
    const fields = {
      ClientId:       body.clientId || '',
      CurrentVersion: body.curVer || '',
      NewVersion:     body.newVer || '',
      Environment:    body.env || 'Sandbox + Prod',
      DateTest:       body.dTest || null,
      DateTestEnd:    body.dTestEnd || null,
      DateProd:       body.dProd || null,
      DateNotify:     body.dNotify || null,
      Status:         body.status || 'planned',
      Notes:          body.notes || '',
    };

    if (body.spItemId) {
      // Update existing
      await updateSharePointItem('BC_Updates', body.spItemId, fields);
      context.res = { status: 200, body: { ok: true } };
    } else {
      // Create new
      const res = await createSharePointItem('BC_Updates', fields);
      context.res = { status: 200, body: { ok: true, spItemId: res.id } };
    }
  } catch (err) {
    context.log.error('saveUpdate error:', err);
    context.res = { status: 500, body: { error: err.message } };
  }
};
