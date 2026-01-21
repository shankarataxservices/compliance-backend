const { withCors, json, db } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');
const ExcelJS = require('exceljs');

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const { clientId, fromYmd, toYmd } = JSON.parse(event.body || '{}');
  if (!clientId || !fromYmd || !toYmd) return json(event, 400, { ok:false, error:'clientId,fromYmd,toYmd required' });

  const clientSnap = await db().collection('clients').doc(clientId).get();
  if (!clientSnap.exists) return json(event, 404, { ok:false, error:'Client not found' });
  const client = clientSnap.data();

  const tasksSnap = await db().collection('tasks')
    .where('clientId', '==', clientId)
    .where('dueDateYmd', '>=', fromYmd)
    .where('dueDateYmd', '<=', toYmd)
    .get();

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('History');

  ws.addRow([
    'Client',
    'Title',
    'Category',
    'Type',
    'Start Date',
    'Due Date',
    'Status',
    'Status Note',
    'Client Start Mail Sent'
  ]);

  for (const tDoc of tasksSnap.docs) {
    const t = tDoc.data();
    ws.addRow([
      client.name,
      t.title,
      t.category,
      t.type,
      t.startDateYmd,
      t.dueDateYmd,
      t.status,
      t.statusNote || '',
      t.clientStartMailSent ? 'YES' : 'NO'
    ]);
  }

  const buf = await wb.xlsx.writeBuffer();
  return json(event, 200, {
    ok: true,
    fileName: `${client.name}_history_${fromYmd}_to_${toYmd}.xlsx`,
    base64: Buffer.from(buf).toString('base64')
  });
});