const { json, db } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');
const ExcelJS = require('exceljs');

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: require('./_common').cors() };
  if (event.httpMethod !== 'POST') return json(405, { error: 'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;
  const p = requirePartner(user);
  if (p.error) return p.error;

  const { clientId, fromYmd, toYmd } = JSON.parse(event.body || '{}');
  if (!clientId || !fromYmd || !toYmd) return json(400, { error: 'clientId,fromYmd,toYmd required' });

  const clientSnap = await db().collection('clients').doc(clientId).get();
  if (!clientSnap.exists) return json(404, { error: 'Client not found' });
  const client = clientSnap.data();

  const tasksSnap = await db().collection('tasks')
    .where('clientId', '==', clientId)
    .where('dueDateYmd', '>=', fromYmd)
    .where('dueDateYmd', '<=', toYmd)
    .get();

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('History');
  ws.addRow(['Client', 'Title', 'Category', 'Type', 'Due Date', 'Start Date', 'Status', 'DelayReason']);

  for (const tDoc of tasksSnap.docs) {
    const t = tDoc.data();
    ws.addRow([
      client.name,
      t.title,
      t.category,
      t.type,
      t.dueDateYmd,
      t.startDateYmd,
      t.status,
      t.delayReason || ''
    ]);
  }

  const buf = await wb.xlsx.writeBuffer();
  return {
    statusCode: 200,
    headers: {
      ...require('./_common').cors(),
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      ok: true,
      fileName: `${client.name}_history_${fromYmd}_to_${toYmd}.xlsx`,
      base64: Buffer.from(buf).toString('base64')
    })
  };
};
