// netlify/functions/reports_clientHistoryPdf.js
const PDFDocument = require('pdfkit');
const { withCors, json, db, ymdToDmy, IST_TZ } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

function tsToIstString(ts) {
  if (!ts || !ts.toDate) return '';
  return ts.toDate().toLocaleString('en-IN', { timeZone: IST_TZ });
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;

  const { user } = authRes;
  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');
  const clientId = String(body.clientId || '').trim();
  const fromYmd = String(body.fromYmd || '').trim();
  const toYmd = String(body.toYmd || '').trim();

  if (!clientId || !fromYmd || !toYmd) {
    return json(event, 400, { ok:false, error:'clientId, fromYmd, toYmd required' });
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fromYmd) || !/^\d{4}-\d{2}-\d{2}$/.test(toYmd)) {
    return json(event, 400, { ok:false, error:'fromYmd/toYmd must be YYYY-MM-DD' });
  }

  const cSnap = await db().collection('clients').doc(clientId).get();
  if (!cSnap.exists) return json(event, 404, { ok:false, error:'Client not found' });
  const client = cSnap.data() || {};

  const tSnap = await db().collection('tasks')
    .where('clientId', '==', clientId)
    .where('dueDateYmd', '>=', fromYmd)
    .where('dueDateYmd', '<=', toYmd)
    .limit(1200)
    .get();

  const tasks = tSnap.docs.map(d => ({ id: d.id, ...d.data() }))
    .sort((a,b)=>String(a.dueDateYmd||'').localeCompare(String(b.dueDateYmd||'')));

  // Build PDF
  const doc = new PDFDocument({ size: 'A4', margin: 40 });
  const chunks = [];
  doc.on('data', (c) => chunks.push(c));
  const done = new Promise((resolve) => doc.on('end', resolve));

  doc.fontSize(16).text('Client History Report', { align: 'left' });
  doc.moveDown(0.2);
  doc.fontSize(11).text(client.name || clientId);
  doc.fontSize(9).fillColor('#444')
    .text(`Range: ${ymdToDmy(fromYmd)} to ${ymdToDmy(toYmd)}   Generated: ${new Date().toLocaleString('en-IN', { timeZone: IST_TZ })}`);
  doc.moveDown(0.7);
  doc.fillColor('#000');

  if (!tasks.length) {
    doc.fontSize(12).text('No Tasks To Display');
    doc.end();
    await done;

    return json(event, 200, {
      ok: true,
      fileName: `client_history_${(client.name || 'client').replace(/[^\w\- ]+/g, '').slice(0, 40)}_${ymdToDmy(fromYmd)}_to_${ymdToDmy(toYmd)}.pdf`,
      base64: Buffer.concat(chunks).toString('base64'),
      mime: 'application/pdf'
    });
  }

  doc.fontSize(11).text(`Tasks: ${tasks.length}`);
  doc.moveDown(0.4);

  for (const t of tasks) {
    const occ = t.seriesId ? `${t.occurrenceIndex || ''}/${t.occurrenceTotal || ''}` : '';
    const line1 = `${t.title || ''}${occ ? ` (Series ${occ})` : ''}`;
    const line2 = `Start: ${ymdToDmy(t.startDateYmd)} | Due: ${ymdToDmy(t.dueDateYmd)} | Status: ${t.status || ''} | Assignee: ${t.assignedToEmail || ''}`;
    const note = (t.statusNote || '').trim();

    doc.fontSize(11).text(line1);
    doc.fontSize(9).fillColor('#444').text(line2);
    if (note) doc.fontSize(9).fillColor('#666').text(`Note: ${note}`);
    doc.fillColor('#000');
    doc.moveDown(0.5);
    if (doc.y > 760) doc.addPage();
  }

  doc.end();
  await done;

  return json(event, 200, {
    ok: true,
    fileName: `client_history_${(client.name || 'client').replace(/[^\w\- ]+/g, '').slice(0, 40)}_${ymdToDmy(fromYmd)}_to_${ymdToDmy(toYmd)}.pdf`,
    base64: Buffer.concat(chunks).toString('base64'),
    mime: 'application/pdf'
  });
});