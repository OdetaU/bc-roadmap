const { getSharePointItems } = require('../shared/graph');

module.exports = async function (context, req) {
  try {
    const items = await getSharePointItems('BC_Updates');
    const updates = items.map(i => ({
      id: i.id,
      clientId: i.ClientId,
      curVer: i.CurrentVersion || '',
      newVer: i.NewVersion || '',
      env: i.Environment || 'Sandbox + Prod',
      dTest: i.DateTest || '',
      dTestEnd: i.DateTestEnd || '',
      dProd: i.DateProd || '',
      dNotify: i.DateNotify || '',
      status: i.Status || 'planned',
      notes: i.Notes || '',
    }));
    context.res = { status: 200, body: updates };
  } catch (err) {
    context.log.error('getUpdates error:', err);
    context.res = { status: 500, body: { error: err.message } };
  }
};
