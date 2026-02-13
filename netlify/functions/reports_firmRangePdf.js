// netlify/functions/reports_firmRangePdf.js
const PDFDocument = require('pdfkit');
const { withCors, json, db, dmyToYmd, ymdToDmy, IST_TZ } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

function tsToIstString(ts) {
  if (!ts || !ts.toDate) return '';
  return ts.toDate().toLocaleString('en-IN', { timeZone: IST_TZ });
}

async function loadClientsMap(clientIds) {
  const ids = [...new Set((clientIds || []).filter(Boolean))];
  const map = new Map();
  if (!ids.length) return map;
  const refs = ids.map(id => db().collection('clients').doc(id));
  const snaps = await db().getAll(...refs);
  snaps.forEach(s => { if (s.exists) map.set(s.id, s.data()); });
  return map;
}

function mustYmdFromInput({ fromDmy, toDmy, fromYmd, toYmd }) {
  const pick = (x) => (x == null ? '' : String(x).trim());
  if (pick(fromDmy) && pick(toDmy)) {
    return { fromY: dmyToYmd(fromDmy), toY: dmyToYmd(toDmy) };
  }
  if (pick(fromYmd) && pick(toYmd)) {
    return { fromY: pick(fromYmd), toY: pick(toYmd) };
  }
  throw new Error('Provide fromDmy & toDmy (DD-MM-YYYY) or fromYmd & toYmd (YYYY-MM-DD)');
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;

  const { user } = authRes;
  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');
  let fromY, toY;
  try { ({ fromY, toY } = mustYmdFromInput(body)); }
  catch (e) { return json(event, 400, { ok:false, error: e.message }); }

  // optional filters
  const clientId = body.clientId ? String(body.clientId) : null;
  const status = body.status ? String(body.status) : null;
  const assignedToEmail = body.assignedToEmail ? String(body.assignedToEmail) : null;

  let q = db().collection('tasks')
    .where('dueDateYmd', '>=', fromY)
    .where('dueDateYmd', '<=', toY);

  if (clientId) q = q.where('clientId', '==', clientId);
  if (status) q = q.where('status', '==', status);
  if (assignedToEmail) q = q.where('assignedToEmail', '==', assignedToEmail);

  const tasksSnap = await q.limit(800).get();
  const tasks = tasksSnap.docs.map(d => ({ id: d.id, ...d.data() }))
    .sort((a,b)=>String(a.dueDateYmd||'').localeCompare(String(b.dueDateYmd||'')));

  const clientsMap = await loadClientsMap(tasks.map(t => t.clientId));

  // Build PDF
  const doc = new PDFDocument({ size: 'A4', margin: 40 });
  const chunks = [];
  doc.on('data', (c) => chunks.push(c));
  const done = new Promise((resolve) => doc.on('end', resolve));

  doc.fontSize(16).text('Firm Tasks Report', { align: 'left' });
  doc.moveDown(0.25);
  doc.fontSize(10).fillColor('#444')
    .text(`Range: ${ymdToDmy(fromY)} to ${ymdToDmy(toY)}   Generated: ${new Date().toLocaleString('en-IN', { timeZone: IST_TZ })}`);
  doc.moveDown(0.8);

  doc.fillColor('#000');

  if (!tasks.length) {
    doc.fontSize(12).text('No Tasks To Display');
    doc.end();
    await done;

    return json(event, 200, {
      ok: true,
      fileName: `firm_tasks_${ymdToDmy(fromY)}_to_${ymdToDmy(toY)}.pdf`,
      base64: Buffer.concat(chunks).toString('base64'),
      mime: 'application/pdf'
    });
  }

  // Simple list layout (PDFKit-friendly)
  doc.fontSize(11).text(`Tasks: ${tasks.length}`);
  doc.moveDown(0.4);

  for (const t of tasks) {
    const c = clientsMap.get(t.clientId) || {};
    const line1 = `${c.name || ''} — ${t.title || ''}`;
    const line2 = `Due: ${ymdToDmy(t.dueDateYmd)} | Start: ${ymdToDmy(t.startDateYmd)} | Status: ${t.status || ''} | Assignee: ${t.assignedToEmail || ''}`;
    const note = (t.statusNote || '').trim();

    doc.fontSize(11).text(line1);
    doc.fontSize(9).fillColor('#444').text(line2);
    if (note) doc.fontSize(9).fillColor('#666').text(`Note: ${note}`);
    doc.fillColor('#000');
    doc.moveDown(0.5);

    // page break safeguard
    if (doc.y > 760) doc.addPage();
  }

  doc.end();
  await done;

  return json(event, 200, {
    ok: true,
    fileName: `firm_tasks_${ymdToDmy(fromY)}_to_${ymdToDmy(toY)}.pdf`,
    base64: Buffer.concat(chunks).toString('base64'),
    mime: 'application/pdf'
  });
});