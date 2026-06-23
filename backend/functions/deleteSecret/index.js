const { deleteSharePointItem } = require('../shared/graph');

module.exports = async function (context, req) {
  try {
    const spItemId = req.params.id || req.query.id;
    if (!spItemId) { context.res = { status: 400, body: { error: 'Missing id' } }; return; }
    await deleteSharePointItem('BC_Secrets', spItemId);
    context.res = { status: 200, body: { ok: true } };
  } catch (err) {
    context.log.error('deleteSecret error:', err);
    context.res = { status: 500, body: { error: err.message } };
  }
};
