const { withCors, json, db, dmyToYmd, ymdToDmy, IST_TZ } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');
const ExcelJS = require('exceljs');

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

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');

  let fromY, toY;
  try {
    ({ fromY, toY } = mustYmdFromInput(body));
  } catch (e) {
    return json(event, 400, { ok:false, error: e.message });
  }

  const limitTasks = Math.min(500, Math.max(10, Number(body.limitTasks || 300))); // safety
  const includeAudit = body.includeAudit !== false;

  // Filter options (optional)
  const clientId = body.clientId ? String(body.clientId) : null;
  const status = body.status ? String(body.status) : null;
  const assignedToEmail = body.assignedToEmail ? String(body.assignedToEmail) : null;

  let q = db().collection('tasks')
    .where('dueDateYmd', '>=', fromY)
    .where('dueDateYmd', '<=', toY);

  if (clientId) q = q.where('clientId', '==', clientId);
  if (status) q = q.where('status', '==', status);
  if (assignedToEmail) q = q.where('assignedToEmail', '==', assignedToEmail);

  const tasksSnap = await q.limit(limitTasks).get();
  const tasks = tasksSnap.docs.map(d => ({ id: d.id, ...d.data() }));

  const clientsMap = await loadClientsMap(tasks.map(t => t.clientId));

  // Workbook
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Compliance Management';
  wb.created = new Date();

  const wsT = wb.addWorksheet('Tasks');
  wsT.views = [{ state: 'frozen', ySplit: 1 }];

  wsT.addRow([
    'TaskId','Client','Title','Category','Type','Recurrence',
    'SeriesId','Occurrence',
    'Start (DD-MM-YYYY)','Due (DD-MM-YYYY)',
    'Status','Assignee',
    'StatusNote','DelayReason','DelayNotes',
    'CompletedRequestedAt (IST)','CompletedAt (IST)',
    'ClientStartMailSentAt (IST)','ClientStartThreadId',
    'SendClientCompletionMail',
    'CalendarEventId',
    'AttachmentLinks',
    'CreatedAt (IST)','UpdatedAt (IST)'
  ]);
  wsT.getRow(1).font = { bold: true };
  wsT.columns.forEach(c => { c.width = 22; });
  wsT.getColumn(2).width = 30;
  wsT.getColumn(3).width = 40;
  wsT.getColumn(23).width = 55;

  for (const t of tasks) {
    const c = clientsMap.get(t.clientId) || {};
    const occ = t.seriesId ? `${t.occurrenceIndex || ''}/${t.occurrenceTotal || ''}` : '';
    const attachLinks = Array.isArray(t.attachments)
      ? t.attachments.map(a => a.driveWebViewLink || '').filter(Boolean).join(' | ')
      : '';

    wsT.addRow([
      t.id,
      c.name || t.clientId || '',
      t.title || '',
      t.category || '',
      t.type || '',
      t.recurrence || '',
      t.seriesId || '',
      occ,
      ymdToDmy(t.startDateYmd),
      ymdToDmy(t.dueDateYmd),
      t.status || '',
      t.assignedToEmail || '',
      t.statusNote || '',
      t.delayReason || '',
      t.delayNotes || '',
      tsToIstString(t.completedRequestedAt),
      tsToIstString(t.completedAt),
      tsToIstString(t.clientStartMailSentAt),
      t.clientStartGmailThreadId || '',
      (t.sendClientCompletionMail === false) ? 'false' : 'true',
      t.calendarEventId || t.calendarStartEventId || '',
      attachLinks,
      tsToIstString(t.createdAt),
      tsToIstString(t.updatedAt)
    ]);
  }

  // Audit sheet
  const wsA = wb.addWorksheet('AuditLogs');
  wsA.views = [{ state: 'frozen', ySplit: 1 }];
  wsA.addRow(['Time (IST)','Action','TaskId','ActorEmail','Details']);
  wsA.getRow(1).font = { bold: true };
  wsA.getColumn(1).width = 22;
  wsA.getColumn(2).width = 22;
  wsA.getColumn(3).width = 26;
  wsA.getColumn(4).width = 26;
  wsA.getColumn(5).width = 80;

  let auditRows = 0;
  const maxAuditRows = Math.min(5000, Math.max(200, Number(body.maxAuditRows || 2000)));

  if (includeAudit) {
    // Best-effort approach (small-firm safe):
    // fetch audit logs per task (no orderBy to avoid composite index requirement),
    // then sort in JS.
    for (const t of tasks) {
      if (auditRows >= maxAuditRows) break;

      const aSnap = await db().collection('auditLogs')
        .where('taskId', '==', t.id)
        .get();

      const logs = aSnap.docs.map(d => d.data());
      logs.sort((x,y) => {
        const a = x.timestamp?.toMillis?.() || 0;
        const b = y.timestamp?.toMillis?.() || 0;
        return a - b;
      });

      for (const a of logs) {
        if (auditRows >= maxAuditRows) break;
        wsA.addRow([
          a.timestamp?.toDate?.() ? a.timestamp.toDate().toLocaleString('en-IN', { timeZone: IST_TZ }) : '',
          a.action || '',
          a.taskId || '',
          a.actorEmail || '',
          JSON.stringify(a.details || {})
        ]);
        auditRows++;
      }
    }
  }

  const buf = await wb.xlsx.writeBuffer();

  return json(event, 200, {
    ok: true,
    fileName: `firm_tasks_${ymdToDmy(fromY)}_to_${ymdToDmy(toY)}.xlsx`,
    meta: { tasks: tasks.length, auditRows, truncatedAudit: includeAudit && auditRows >= maxAuditRows },
    base64: Buffer.from(buf).toString('base64')
  });
});