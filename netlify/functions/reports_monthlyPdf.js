// netlify/functions/reports_monthlyPdf.js
// Optional monthly PDF summary (firm-wide). Partner-only.
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
  // monthYmd: any date within month (YYYY-MM-DD). Default: today
  const monthYmd = String(body.monthYmd || ymdIST(new Date())).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(monthYmd)) return json(event, 400, { ok:false, error:'monthYmd must be YYYY-MM-DD' });

  const dt = dateFromYmdIST(monthYmd);
  const y = dt.getUTCFullYear();
  const m = dt.getUTCMonth(); // 0-based in UTC
  const monthStart = new Date(Date.UTC(y, m, 1));
  const monthEnd = new Date(Date.UTC(y, m + 1, 0)); // last day
  const fromYmd = ymdIST(monthStart);
  const toYmd = ymdIST(monthEnd);

  const activeStatuses = ['PENDING','IN_PROGRESS','CLIENT_PENDING','APPROVAL_PENDING'];

  const snap = await db().collection('tasks')
    .where('status', 'in', activeStatuses)
    .where('dueDateYmd', '>=', fromYmd)
    .where('dueDateYmd', '<=', toYmd)
    .limit(1500)
    .get();

  const tasks = snap.docs.map(d => ({ id:d.id, ...d.data() }))
    .sort((a,b)=>String(a.dueDateYmd||'').localeCompare(String(b.dueDateYmd||'')));

  const clientsMap = await loadClientsMap(tasks.map(t => t.clientId));

  const doc = new PDFDocument({ size:'A4', margin:40 });
  const chunks = [];
  doc.on('data', c => chunks.push(c));
  const done = new Promise(resolve => doc.on('end', resolve));

  doc.fontSize(16).text('Monthly Compliance Summary');
  doc.moveDown(0.2);
  doc.fontSize(10).fillColor('#444')
    .text(`Month: ${fromYmd.slice(0,7)}   Range: ${ymdToDmy(fromYmd)} to ${ymdToDmy(toYmd)}`)
    .text(`Generated: ${new Date().toLocaleString('en-IN',{ timeZone: IST_TZ })}`);
  doc.fillColor('#000');
  doc.moveDown(0.8);

  if (!tasks.length) {
    doc.fontSize(12).text('No Tasks To Display');
    doc.end();
    await done;
    return json(event, 200, {
      ok:true,
      fileName: `monthly_summary_${fromYmd.slice(0,7)}.pdf`,
      base64: Buffer.concat(chunks).toString('base64'),
      mime: 'application/pdf'
    });
  }

  doc.fontSize(11).text(`Open items due this month: ${tasks.length}`);
  doc.moveDown(0.5);

  for (const t of tasks) {
    const c = clientsMap.get(t.clientId) || {};
    doc.fontSize(11).text(`${c.name || ''} — ${t.title || ''}`);
    doc.fontSize(9).fillColor('#444')
      .text(`Due: ${ymdToDmy(t.dueDateYmd)} | Status: ${t.status || ''} | Assignee: ${t.assignedToEmail || ''}`);
    doc.fillColor('#000');
    doc.moveDown(0.4);
    if (doc.y > 760) doc.addPage();
  }

  doc.end();
  await done;

  return json(event, 200, {
    ok:true,
    fileName: `monthly_summary_${fromYmd.slice(0,7)}.pdf`,
    base64: Buffer.concat(chunks).toString('base64'),
    mime: 'application/pdf'
  });
});