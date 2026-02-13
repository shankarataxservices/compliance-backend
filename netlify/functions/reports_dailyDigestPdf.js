// netlify/functions/reports_dailyDigestPdf.js
// Optional: firm daily digest as PDF (Partner-only)
const PDFDocument = require('pdfkit');
const { withCors, json, db, ymdIST, ymdToDmy, addDays, dateFromYmdIST, IST_TZ } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

async function loadClientsMap(clientIds) {
  const ids = [...new Set((clientIds || []).filter(Boolean))];
  const map = new Map();
  if (!ids.length) return map;
  const refs = ids.map(id => db().collection('clients').doc(id));
  const snaps = await db().getAll(...refs);
  snaps.forEach(s => { if (s.exists) map.set(s.id, s.data()); });
  return map;
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;

  const { user } = authRes;
  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');
  const todayYmd = ymdIST(new Date());

  // Use same settings doc as email digest
  const settingsSnap = await db().collection('settings').doc('notifications').get();
  const settings = settingsSnap.exists ? settingsSnap.data() : { dailyWindowDays: 30 };
  const windowDays = Number(settings.dailyWindowDays || 30);
  const endYmd = ymdIST(addDays(dateFromYmdIST(todayYmd), windowDays));

  const activeStatuses = ['PENDING','IN_PROGRESS','CLIENT_PENDING','APPROVAL_PENDING'];

  const dueSoonSnap = await db().collection('tasks')
    .where('status', 'in', activeStatuses)
    .where('dueDateYmd', '>=', todayYmd)
    .where('dueDateYmd', '<=', endYmd)
    .limit(1500)
    .get();

  const overdueSnap = await db().collection('tasks')
    .where('status', 'in', activeStatuses)
    .where('dueDateYmd', '<', todayYmd)
    .limit(1500)
    .get();

  const tasks = [
    ...dueSoonSnap.docs.map(d => ({ id:d.id, ...d.data() })),
    ...overdueSnap.docs.map(d => ({ id:d.id, ...d.data() })),
  ].sort((a,b)=>String(a.dueDateYmd||'').localeCompare(String(b.dueDateYmd||'')));

  const clientsMap = await loadClientsMap(tasks.map(t => t.clientId));

  const doc = new PDFDocument({ size:'A4', margin:40 });
  const chunks = [];
  doc.on('data', c => chunks.push(c));
  const done = new Promise(resolve => doc.on('end', resolve));

  doc.fontSize(16).text('Daily Digest (PDF)');
  doc.moveDown(0.2);
  doc.fontSize(10).fillColor('#444')
    .text(`Date: ${ymdToDmy(todayYmd)}   Window: next ${windowDays} days + overdue`)
    .text(`Generated: ${new Date().toLocaleString('en-IN',{ timeZone: IST_TZ })}`);
  doc.fillColor('#000');
  doc.moveDown(0.8);

  if (!tasks.length) {
    doc.fontSize(12).text('No Tasks To Display');
    doc.end();
    await done;
    return json(event, 200, {
      ok:true,
      fileName: `daily_digest_${ymdToDmy(todayYmd)}.pdf`,
      base64: Buffer.concat(chunks).toString('base64'),
      mime: 'application/pdf'
    });
  }

  doc.fontSize(11).text(`Tasks: ${tasks.length}`);
  doc.moveDown(0.4);

  for (const t of tasks) {
    const c = clientsMap.get(t.clientId) || {};
    doc.fontSize(11).text(`${c.name || ''} — ${t.title || ''}`);
    doc.fontSize(9).fillColor('#444')
      .text(`Due: ${ymdToDmy(t.dueDateYmd)} | Status: ${t.status || ''} | Assignee: ${t.assignedToEmail || ''}`);
    const note = (t.statusNote || '').trim();
    if (note) doc.fontSize(9).fillColor('#666').text(`Note: ${note}`);
    doc.fillColor('#000');
    doc.moveDown(0.4);
    if (doc.y > 760) doc.addPage();
  }

  doc.end();
  await done;

  return json(event, 200, {
    ok:true,
    fileName: `daily_digest_${ymdToDmy(todayYmd)}.pdf`,
    base64: Buffer.concat(chunks).toString('base64'),
    mime: 'application/pdf'
  });
});