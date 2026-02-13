// netlify/functions/exports_taskHistoryXlsx.js
const { withCors, json, db, ymdToDmy, IST_TZ } = require('./_common');
const { requireUser } = require('./_auth');
const ExcelJS = require('exceljs');

function tsToIstString(ts) {
  if (!ts || !ts.toDate) return '';
  return ts.toDate().toLocaleString('en-IN', { timeZone: IST_TZ });
}
function xstr(v) {
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  try { return JSON.stringify(v); } catch { return String(v); }
}
function roleOf(user) {
  let r = String(user?.role || 'ASSOCIATE').toUpperCase().trim();
  if (r === 'WORKER') r = 'ASSOCIATE';
  return r || 'ASSOCIATE';
}
function isPrivileged(role) {
  return role === 'PARTNER' || role === 'MANAGER';
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const role = roleOf(user);
  const privileged = isPrivileged(role);

  const body = JSON.parse(event.body || '{}');
  const taskId = String(body.taskId || '').trim();
  if (!taskId) return json(event, 400, { ok:false, error:'taskId required' });

  const tSnap = await db().collection('tasks').doc(taskId).get();
  if (!tSnap.exists) return json(event, 404, { ok:false, error:'Task not found' });
  const task = tSnap.data();

  if (!privileged && task.assignedToUid !== user.uid) {
    return json(event, 403, { ok:false, error:'Not allowed' });
  }

  const clientSnap = task.clientId ? await db().collection('clients').doc(task.clientId).get() : null;
  const client = clientSnap && clientSnap.exists ? clientSnap.data() : {};

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Compliance Management';
  wb.created = new Date();

  // Sheet: Task
  const wsT = wb.addWorksheet('Task');
  wsT.addRow(['Field','Value']);
  wsT.getRow(1).font = { bold: true };
  const addKV = (k, v) => wsT.addRow([k, xstr(v)]);

  addKV('TaskId', taskId);
  addKV('Client', client.name || '');
  addKV('Title', task.title || '');
  addKV('Category', task.category || '');
  addKV('Type', task.type || '');
  addKV('Priority', task.priority || '');
  addKV('Recurrence', task.recurrence || '');
  addKV('SeriesId', task.seriesId || '');
  addKV('Occurrence', task.seriesId ? `${task.occurrenceIndex || ''}/${task.occurrenceTotal || ''}` : '');
  addKV('Start (DD-MM-YYYY)', ymdToDmy(task.startDateYmd));
  addKV('Due (DD-MM-YYYY)', ymdToDmy(task.dueDateYmd));
  addKV('Status', task.status || '');
  addKV('Assignee', task.assignedToEmail || '');
  addKV('StatusNote', task.statusNote || '');
  addKV('DelayReason', task.delayReason || '');
  addKV('DelayNotes', task.delayNotes || '');
  addKV('SnoozedUntil', ymdToDmy(task.snoozedUntilYmd));
  addKV('StartMailSentAt (IST)', tsToIstString(task.clientStartMailSentAt));
  addKV('StartThreadId', task.clientStartGmailThreadId || '');
  addKV('SendCompletionMail', (task.sendClientCompletionMail === false) ? 'false' : 'true');
  addKV('CalendarEventId', task.calendarEventId || task.calendarStartEventId || '');
  addKV('CreatedAt (IST)', tsToIstString(task.createdAt));
  addKV('UpdatedAt (IST)', tsToIstString(task.updatedAt));

  wsT.getColumn(1).width = 28;
  wsT.getColumn(2).width = 80;

  // Sheet: AuditLogs
  const wsA = wb.addWorksheet('AuditLogs');
  wsA.views = [{ state: 'frozen', ySplit: 1 }];
  wsA.addRow(['Time (IST)','Action','ActorEmail','Details']);
  wsA.getRow(1).font = { bold: true };
  wsA.getColumn(1).width = 22;
  wsA.getColumn(2).width = 22;
  wsA.getColumn(3).width = 26;
  wsA.getColumn(4).width = 90;

  const aSnap = await db().collection('auditLogs').where('taskId', '==', taskId).get();
  const logs = aSnap.docs.map(d => d.data());
  logs.sort((x,y)=> (x.timestamp?.toMillis?.()||0) - (y.timestamp?.toMillis?.()||0));
  for (const a of logs) {
    wsA.addRow([
      a.timestamp?.toDate?.() ? a.timestamp.toDate().toLocaleString('en-IN', { timeZone: IST_TZ }) : '',
      xstr(a.action || ''),
      xstr(a.actorEmail || ''),
      xstr(a.details || {})
    ]);
  }

  // Sheet: Comments
  const wsC = wb.addWorksheet('Comments');
  wsC.views = [{ state: 'frozen', ySplit: 1 }];
  wsC.addRow(['Time (IST)','Author','Text']);
  wsC.getRow(1).font = { bold: true };
  wsC.getColumn(1).width = 22;
  wsC.getColumn(2).width = 26;
  wsC.getColumn(3).width = 90;

  const cSnap = await db().collection('tasks').doc(taskId).collection('comments')
    .orderBy('createdAt', 'asc')
    .limit(500)
    .get();
  cSnap.docs.forEach(d => {
    const c = d.data();
    wsC.addRow([
      c.createdAt?.toDate?.() ? c.createdAt.toDate().toLocaleString('en-IN', { timeZone: IST_TZ }) : '',
      (c.authorName || c.authorEmail || ''),
      xstr(c.text || '')
    ]);
  });

  // Sheet: Attachments
  const wsF = wb.addWorksheet('Attachments');
  wsF.views = [{ state: 'frozen', ySplit: 1 }];
  wsF.addRow(['Type','FileName','Link','UploadedBy','UploadedAt (IST)']);
  wsF.getRow(1).font = { bold: true };
  wsF.getColumn(1).width = 14;
  wsF.getColumn(2).width = 32;
  wsF.getColumn(3).width = 70;
  wsF.getColumn(4).width = 20;
  wsF.getColumn(5).width = 22;

  const att = Array.isArray(task.attachments) ? task.attachments : [];
  att.forEach(a => {
    wsF.addRow([
      xstr(a.type || ''),
      xstr(a.fileName || ''),
      xstr(a.driveWebViewLink || ''),
      xstr(a.uploadedByUid || ''),
      a.uploadedAt?.toDate?.() ? a.uploadedAt.toDate().toLocaleString('en-IN', { timeZone: IST_TZ }) : ''
    ]);
  });

  const buf = await wb.xlsx.writeBuffer();
  const safeTitle = String(task.title || 'task').replace(/[^\w\- ]+/g, '').slice(0, 40).trim() || 'task';

  return json(event, 200, {
    ok: true,
    fileName: `task_history_${safeTitle}_${taskId}.xlsx`,
    base64: Buffer.from(buf).toString('base64')
  });
});