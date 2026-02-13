// netlify/functions/exports_clientHistoryXlsx.js
const { withCors, json, db, ymdToDmy, IST_TZ } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');
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
  const clientId = String(body.clientId || '').trim();
  const fromYmd = String(body.fromYmd || '').trim();
  const toYmd = String(body.toYmd || '').trim();

  if (!clientId || !fromYmd || !toYmd) {
    return json(event, 400, { ok:false, error:'clientId,fromYmd,toYmd required' });
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fromYmd) || !/^\d{4}-\d{2}-\d{2}$/.test(toYmd)) {
    return json(event, 400, { ok:false, error:'fromYmd/toYmd must be YYYY-MM-DD' });
  }

  const clientSnap = await db().collection('clients').doc(clientId).get();
  if (!clientSnap.exists) return json(event, 404, { ok:false, error:'Client not found' });

  const client = clientSnap.data() || {};

  const tasksSnap = await db().collection('tasks')
    .where('clientId', '==', clientId)
    .where('dueDateYmd', '>=', fromYmd)
    .where('dueDateYmd', '<=', toYmd)
    .limit(1500)
    .get();

  const tasks = tasksSnap.docs.map(d => ({ id: d.id, ...d.data() }))
    .sort((a, b) => String(a.dueDateYmd || '').localeCompare(String(b.dueDateYmd || '')));

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Compliance Management';
  wb.created = new Date();

  const ws = wb.addWorksheet('ClientHistory');
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  ws.addRow([
    'TaskId',
    'Client',
    'Title',
    'Category',
    'Type',
    'Priority',
    'Recurrence',
    'SeriesId',
    'Occurrence',

    'Start (DD-MM-YYYY)',
    'Due (DD-MM-YYYY)',
    'TriggerDays',

    'Status',
    'Assignee',
    'StatusNote',
    'DelayReason',
    'DelayNotes',
    'SnoozedUntil (DD-MM-YYYY)',

    'CompletedRequestedAt (IST)',
    'CompletedAt (IST)',

    'SendStartMail',
    'ClientTo',
    'ClientCC',
    'ClientBCC',
    'CcAssigneeOnStart',
    'CcManagerOnStart',
    'ClientStartSubject',
    'ClientStartBody',
    'ClientStartMailSentAt (IST)',
    'ClientStartThreadId',
    'ClientStartMessageId',
    'ClientStartReferences',

    'SendClientCompletionMail',
    'CompletionTo',
    'CompletionCC',
    'CompletionBCC',
    'CcAssigneeOnCompletion',
    'CcManagerOnCompletion',
    'ClientCompletionSubject',
    'ClientCompletionBody',

    'CalendarEventId',
    'AttachmentLinks',

    'CreatedAt (IST)',
    'UpdatedAt (IST)'
  ]);

  ws.getRow(1).font = { bold: true };
  ws.columns.forEach(c => (c.width = 22));
  ws.getColumn(2).width = 30;   // Client
  ws.getColumn(3).width = 42;   // Title
  ws.getColumn(28).width = 44;  // Start body
  ws.getColumn(39).width = 44;  // Completion body
  ws.getColumn(41).width = 55;  // Attachment links

  for (const t of tasks) {
    const occ = t.seriesId ? `${t.occurrenceIndex || ''}/${t.occurrenceTotal || ''}` : '';
    const attachLinks = Array.isArray(t.attachments)
      ? t.attachments.map(a => a.driveWebViewLink || '').filter(Boolean).join(' | ')
      : '';

    ws.addRow([
      xstr(t.id),
      xstr(client.name || ''),
      xstr(t.title || ''),
      xstr(t.category || ''),
      xstr(t.type || ''),
      xstr(t.priority || 'MEDIUM'),
      xstr(t.recurrence || ''),
      xstr(t.seriesId || ''),
      xstr(occ),

      xstr(ymdToDmy(t.startDateYmd)),
      xstr(ymdToDmy(t.dueDateYmd)),
      xstr(t.triggerDaysBefore ?? ''),

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
      xstr(tsToIstString(t.updatedAt))
    ]);
  }

  const buf = await wb.xlsx.writeBuffer();

  const safeClient = String(client.name || 'client')
    .replace(/[^\w\- ]+/g, '')
    .slice(0, 50)
    .trim() || 'client';

  return json(event, 200, {
    ok: true,
    fileName: `${safeClient}_history_${ymdToDmy(fromYmd)}_to_${ymdToDmy(toYmd)}.xlsx`,
    base64: Buffer.from(buf).toString('base64')
  });
});