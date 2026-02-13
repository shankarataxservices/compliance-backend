// netlify/functions/exports_tasksExportForUpdateXlsx.js
const { withCors, json, db, ymdToDmy } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');
const ExcelJS = require('exceljs');

function xstr(v) {
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  try { return JSON.stringify(v); } catch { return String(v); }
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');

  // Optional filters (to keep file manageable)
  const clientId = body.clientId ? String(body.clientId) : null;
  const status = body.status ? String(body.status) : null;
  const assignedToEmail = body.assignedToEmail ? String(body.assignedToEmail) : null;

  const limitTasks = Math.min(1200, Math.max(10, Number(body.limitTasks || 600)));

  let q = db().collection('tasks');
  if (clientId) q = q.where('clientId', '==', clientId);
  if (status) q = q.where('status', '==', status);
  if (assignedToEmail) q = q.where('assignedToEmail', '==', assignedToEmail);

  const snap = await q.limit(limitTasks).get();
  const tasks = snap.docs.map(d => ({ id: d.id, ...d.data() }))
    .sort((a,b)=>String(a.dueDateYmd||'').localeCompare(String(b.dueDateYmd||'')));

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Compliance Management';
  wb.created = new Date();

  const ws = wb.addWorksheet('Update');
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  const headers = [
    'TaskId',
    'Title',
    'Category',
    'Type',
    'Priority',
    'DueDate (DD-MM-YYYY)',
    'TriggerDays',
    'AssignedToEmail',
    'Status',
    'StatusNote',
    'DelayReason',
    'DelayNotes',
    'SnoozedUntil (DD-MM-YYYY)',
    'SendStartMail (true/false)',
    'ClientTo (emails ; , : separated)',
    'ClientCC (emails ; , : separated)',
    'ClientBCC (emails ; , : separated)',
    'CcAssigneeOnClientStart (true/false)',
    'CcManagerOnClientStart (true/false)',
    'ClientStartSubject',
    'ClientStartBody',
    'SendClientCompletionMail (true/false)',
    'CompletionTo (emails ; , : separated)',
    'CompletionCC (emails ; , : separated)',
    'CompletionBCC (emails ; , : separated)',
    'CcAssigneeOnCompletion (true/false)',
    'CcManagerOnCompletion (true/false)',
    'ClientCompletionSubject',
    'ClientCompletionBody',
  ];

  ws.addRow(headers);
  ws.getRow(1).font = { bold: true };

  ws.columns = [
    { width: 26 }, { width: 36 }, { width: 16 }, { width: 20 }, { width: 12 },
    { width: 20 }, { width: 12 }, { width: 26 }, { width: 18 }, { width: 36 },
    { width: 16 }, { width: 32 }, { width: 22 },
    { width: 18 }, { width: 28 }, { width: 28 }, { width: 28 },
    { width: 22 }, { width: 22 },
    { width: 30 }, { width: 44 },
    { width: 26 }, { width: 28 }, { width: 28 }, { width: 28 },
    { width: 26 }, { width: 26 },
    { width: 30 }, { width: 44 },
  ];

  for (const t of tasks) {
    ws.addRow([
      xstr(t.id),
      xstr(t.title || ''),
      xstr(t.category || ''),
      xstr(t.type || ''),
      xstr(t.priority || 'MEDIUM'),
      xstr(ymdToDmy(t.dueDateYmd)),
      xstr(t.triggerDaysBefore ?? t.triggerDaysBefore ?? t.triggerDaysBefore), // keep simple if field missing
      xstr(t.assignedToEmail || ''),
      xstr(t.status || ''),
      xstr(t.statusNote || ''),
      xstr(t.delayReason || ''),
      xstr(t.delayNotes || ''),
      xstr(ymdToDmy(t.snoozedUntilYmd)),
      xstr((t.sendClientStartMail === false) ? 'false' : 'true'),
      xstr(Array.isArray(t.clientToEmails) ? t.clientToEmails.join(';') : ''),
      xstr(Array.isArray(t.clientCcEmails) ? t.clientCcEmails.join(';') : ''),
      xstr(Array.isArray(t.clientBccEmails) ? t.clientBccEmails.join(';') : ''),
      xstr(t.ccAssigneeOnClientStart === true ? 'true' : 'false'),
      xstr(t.ccManagerOnClientStart === true ? 'true' : 'false'),
      xstr(t.clientStartSubject || ''),
      xstr(t.clientStartBody || ''),
      xstr((t.sendClientCompletionMail === false) ? 'false' : 'true'),
      xstr(Array.isArray(t.completionToEmails) ? t.completionToEmails.join(';') : ''),
      xstr(Array.isArray(t.completionCcEmails) ? t.completionCcEmails.join(';') : ''),
      xstr(Array.isArray(t.completionBccEmails) ? t.completionBccEmails.join(';') : ''),
      xstr(t.ccAssigneeOnCompletion === true ? 'true' : 'false'),
      xstr(t.ccManagerOnCompletion === true ? 'true' : 'false'),
      xstr(t.clientCompletionSubject || ''),
      xstr(t.clientCompletionBody || ''),
    ]);
  }

  const buf = await wb.xlsx.writeBuffer();
  return json(event, 200, {
    ok: true,
    fileName: `tasks_export_for_update.xlsx`,
    count: tasks.length,
    base64: Buffer.from(buf).toString('base64')
  });
});