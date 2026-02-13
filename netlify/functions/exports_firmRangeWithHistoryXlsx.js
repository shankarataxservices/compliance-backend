// netlify/functions/exports_firmRangeWithHistoryXlsx.js
const { withCors, json, db, dmyToYmd, ymdToDmy, IST_TZ } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');
const ExcelJS = require('exceljs');

function mustYmdFromInput({ fromDmy, toDmy, fromYmd, toYmd }) {
  const pick = (x) => (x == null ? '' : String(x).trim());
  if (pick(fromDmy) && pick(toDmy)) return { fromY: dmyToYmd(fromDmy), toY: dmyToYmd(toDmy) };
  if (pick(fromYmd) && pick(toYmd)) return { fromY: pick(fromYmd), toY: pick(toYmd) };
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
function xstr(v) {
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  try { return JSON.stringify(v); } catch { return String(v); }
}
function joinEmails(arr) {
  return Array.isArray(arr) ? arr.filter(Boolean).join(';') : '';
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

  const limitTasks = Math.min(800, Math.max(10, Number(body.limitTasks || 400)));
  const includeAudit = body.includeAudit !== false;

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

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Compliance Management';
  wb.created = new Date();

  // ===== Tasks sheet =====
  const wsT = wb.addWorksheet('Tasks');
  wsT.views = [{ state: 'frozen', ySplit: 1 }];

  wsT.addRow([
    'TaskId','Client','Title','Category','Type','Priority','Recurrence',
    'SeriesId','Occurrence',
    'Start (DD-MM-YYYY)','Due (DD-MM-YYYY)',
    'Status','Assignee',
    'StatusNote','DelayReason','DelayNotes',
    'SnoozedUntil (DD-MM-YYYY)',
    'CompletedRequestedAt (IST)','CompletedAt (IST)',

    'SendStartMail',
    'ClientTo','ClientCC','ClientBCC',
    'CcAssigneeOnStart','CcManagerOnStart',
    'ClientStartSubject','ClientStartBody',
    'ClientStartMailSentAt (IST)','ClientStartThreadId','ClientStartMessageId','ClientStartReferences',

    'SendClientCompletionMail',
    'CompletionTo','CompletionCC','CompletionBCC',
    'CcAssigneeOnCompletion','CcManagerOnCompletion',
    'ClientCompletionSubject','ClientCompletionBody',

    'CalendarEventId',
    'AttachmentLinks',
    'CreatedAt (IST)','UpdatedAt (IST)'
  ]);

  wsT.getRow(1).font = { bold: true };
  wsT.columns.forEach(c => { c.width = 22; });
  wsT.getColumn(2).width = 30;
  wsT.getColumn(3).width = 42;
  wsT.getColumn(28).width = 42; // start body
  wsT.getColumn(38).width = 42; // completion body
  wsT.getColumn(40).width = 55; // attachments
  wsT.getColumn(41).width = 22;
  wsT.getColumn(42).width = 22;

  for (const t of tasks) {
    const c = clientsMap.get(t.clientId) || {};
    const occ = t.seriesId ? `${t.occurrenceIndex || ''}/${t.occurrenceTotal || ''}` : '';
    const attachLinks = Array.isArray(t.attachments)
      ? t.attachments.map(a => a.driveWebViewLink || '').filter(Boolean).join(' | ')
      : '';

    wsT.addRow([
      xstr(t.id),
      xstr(c.name || ''),
      xstr(t.title || ''),
      xstr(t.category || ''),
      xstr(t.type || ''),
      xstr(t.priority || 'MEDIUM'),
      xstr(t.recurrence || ''),
      xstr(t.seriesId || ''),
      xstr(occ),
      xstr(ymdToDmy(t.startDateYmd)),
      xstr(ymdToDmy(t.dueDateYmd)),
      xstr(t.status || ''),
      xstr(t.assignedToEmail || ''),
      xstr(t.statusNote || ''),
      xstr(t.delayReason || ''),
      xstr(t.delayNotes || ''),
      xstr(ymdToDmy(t.snoozedUntilYmd)),
      xstr(tsToIstString(t.completedRequestedAt)),
      xstr(tsToIstString(t.completedAt)),

      xstr((t.sendClientStartMail === false) ? 'false' : 'true'),
      xstr(joinEmails(t.clientToEmails)),
      xstr(joinEmails(t.clientCcEmails)),
      xstr(joinEmails(t.clientBccEmails)),
      xstr(t.ccAssigneeOnClientStart === true ? 'true' : 'false'),
      xstr(t.ccManagerOnClientStart === true ? 'true' : 'false'),
      xstr(t.clientStartSubject || ''),
      xstr(t.clientStartBody || ''),
      xstr(tsToIstString(t.clientStartMailSentAt)),
      xstr(t.clientStartGmailThreadId || ''),
      xstr(t.clientStartRfcMessageId || ''),
      xstr(t.clientStartReferences || ''),

      xstr((t.sendClientCompletionMail === false) ? 'false' : 'true'),
      xstr(joinEmails(t.completionToEmails)),
      xstr(joinEmails(t.completionCcEmails)),
      xstr(joinEmails(t.completionBccEmails)),
      xstr(t.ccAssigneeOnCompletion === true ? 'true' : 'false'),
      xstr(t.ccManagerOnCompletion === true ? 'true' : 'false'),
      xstr(t.clientCompletionSubject || ''),
      xstr(t.clientCompletionBody || ''),

      xstr(t.calendarEventId || t.calendarStartEventId || ''),
      xstr(attachLinks),
      xstr(tsToIstString(t.createdAt)),
      xstr(tsToIstString(t.updatedAt)),
    ]);
  }

  // ===== Audit sheet =====
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
    for (const t of tasks) {
      if (auditRows >= maxAuditRows) break;

      const aSnap = await db().collection('auditLogs').where('taskId', '==', t.id).get();
      const logs = aSnap.docs.map(d => d.data());
      logs.sort((x,y) => (x.timestamp?.toMillis?.()||0) - (y.timestamp?.toMillis?.()||0));

      for (const a of logs) {
        if (auditRows >= maxAuditRows) break;
        wsA.addRow([
          a.timestamp?.toDate?.() ? a.timestamp.toDate().toLocaleString('en-IN', { timeZone: IST_TZ }) : '',
          xstr(a.action || ''),
          xstr(a.taskId || ''),
          xstr(a.actorEmail || ''),
          xstr(a.details || {}),
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