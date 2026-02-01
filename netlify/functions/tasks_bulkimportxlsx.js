const Busboy = require('busboy');
const ExcelJS = require('exceljs');

const {
  withCors, json, db, admin,
  calendar,
  dmyToYmd, ymdIST, dateFromYmdIST,
  addInterval, addDays,
  getCalendarWindow, calTimeRange,
  auditLog, asEmailList
} = require('./_common');

const { requireUser, requirePartner } = require('./_auth');

async function findOrCreateClientByName(clientName) {
  const snap = await db().collection('clients').where('name', '==', clientName).limit(1).get();
  if (!snap.empty) return snap.docs[0].id;

  const ref = db().collection('clients').doc();
  await ref.set({
    name: clientName,
    pan: '', gstin: '', cin: '',
    assessmentYear: '', engagementType: '',
    primaryEmail: '',
    ccEmails: [], bccEmails: [],
    driveFolderId: null,
    createdAt: admin.firestore.FieldValue.serverTimestamp()
  });
  return ref.id;
}

async function findUserUidByEmail(email) {
  if (!email) return null;
  const snap = await db().collection('users').where('email', '==', email).limit(1).get();
  if (snap.empty) return null;
  return snap.docs[0].id;
}

function normalizeRecurrence(x) {
  const r = String(x || 'AD_HOC').toUpperCase().trim();
  const allowed = ['AD_HOC','DAILY','WEEKLY','BIWEEKLY','MONTHLY','BIMONTHLY','QUARTERLY','HALF_YEARLY','YEARLY'];
  return allowed.includes(r) ? r : 'AD_HOC';
}

function normalizeCategory(x) {
  const raw = String(x || 'OTHER').trim();
  const u = raw.toUpperCase().replace(/\s+/g, '_');
  if (u === 'ITR' || u === 'INCOME_TAX' || u === 'INCOME-TAX' || u === 'INCOME' || u === 'INCOME_TAX_RETURN') return 'INCOME_TAX';
  if (u === 'INCOME_TAX' || u === 'INCOME_TAX') return 'INCOME_TAX';
  if (u === 'INCOME_TAX' || u === 'INCOME_TAX') return 'INCOME_TAX';
  if (u === 'INCOME_TAX' || u === 'INCOME_TAX') return 'INCOME_TAX';
  if (u === 'INCOME_TAX') return 'INCOME_TAX';

  // allow friendly labels from template
  if (String(x).trim().toLowerCase() === 'income tax') return 'INCOME_TAX';

  if (u === 'GST') return 'GST';
  if (u === 'TDS') return 'TDS';
  if (u === 'ROC') return 'ROC';
  if (u === 'ACCOUNTING') return 'ACCOUNTING';
  if (u === 'AUDIT') return 'AUDIT';
  return 'OTHER';
}

function truthy(x) {
  const s = String(x ?? '').trim().toLowerCase();
  if (!s) return false;
  return ['1','true','yes','y'].includes(s);
}

async function createStartCalendarEvent({ title, clientId, startDateYmd, dueDateYmd, window }) {
  const cal = calendar();
  const range = calTimeRange(startDateYmd, window.startHH, window.endHH, window.timeZone);

  const res = await cal.events.insert({
    calendarId: 'primary',
    sendUpdates: 'none',
    requestBody: {
      summary: `START: ${title}`,
      description:
        `ClientId: ${clientId}\n` +
        `Start: ${startDateYmd}\n` +
        `Due: ${dueDateYmd}\n`,
      ...range
    }
  });

  return { calendarEventId: res.data.id, calendarHtmlLink: res.data.htmlLink || null };
}

function getCellString(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (v == null) return '';
  if (typeof v === 'string') return v.trim();
  if (typeof v === 'number') return String(v);
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'object' && v.text) return String(v.text).trim();
  return String(v).trim();
}

function parseDueDateDmyFromExcel(cell) {
  if (!cell || cell.value == null) return null;
  const v = cell.value;

  // If Excel stored as Date:
  if (v instanceof Date) {
    // convert Date -> DD-MM-YYYY in IST-ish (date portion)
    const yyyy = v.getFullYear();
    const mm = String(v.getMonth() + 1).padStart(2,'0');
    const dd = String(v.getDate()).padStart(2,'0');
    return `${dd}-${mm}-${yyyy}`;
  }

  // If string already DD-MM-YYYY
  const s = getCellString(cell);
  if (!s) return null;
  return s;
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const busboy = Busboy({ headers: event.headers });
  let fileBuffer = Buffer.alloc(0);
  let fileName = 'import.xlsx';

  busboy.on('file', (name, file, info) => {
    fileName = info.filename || 'import.xlsx';
    file.on('data', (d) => { fileBuffer = Buffer.concat([fileBuffer, d]); });
  });

  const done = new Promise((resolve, reject) => {
    busboy.on('finish', resolve);
    busboy.on('error', reject);
  });

  const bodyBuf = event.isBase64Encoded ? Buffer.from(event.body, 'base64') : Buffer.from(event.body || '');
  busboy.end(bodyBuf);
  await done;

  if (!fileBuffer.length) return json(event, 400, { ok:false, error:'XLSX file missing in form-data' });

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(fileBuffer);
  const ws = wb.getWorksheet('Import') || wb.worksheets[0];
  if (!ws) return json(event, 400, { ok:false, error:'No worksheet found' });

  // Read header row (row 1)
  const headerRow = ws.getRow(1);
  const headers = {};
  headerRow.eachCell((cell, colNumber) => {
    const key = getCellString(cell);
    if (key) headers[key] = colNumber;
  });

  function colOf(names) {
    for (const n of names) {
      if (headers[n]) return headers[n];
    }
    return null;
  }

  const cTitle = colOf(['Title']);
  const cClient = colOf(['Client']);
  const cClientEmail = colOf(['ClientEmail']);
  const cDue = colOf(['DueDate (DD-MM-YYYY)','DueDate','Due']);
  const cCategory = colOf(['Category']);
  const cType = colOf(['Type (you can type custom)','Type']);
  const cRec = colOf(['Recurrence']);
  const cGen = colOf(['GenerateCount']);
  const cTrig = colOf(['TriggerDays']);
  const cAssigned = colOf(['AssignedToEmail']);

  const cTo = colOf(['ClientTo (emails ; , : separated)','ClientTo']);
  const cCc = colOf(['ClientCC (emails ; , : separated)','ClientCC']);
  const cBcc = colOf(['ClientBCC (emails ; , : separated)','ClientBCC']);

  const cStartSub = colOf(['ClientStartSubject']);
  const cStartBody = colOf(['ClientStartBody']);

  const cSendComp = colOf(['SendClientCompletionMail (true/false)','SendClientCompletionMail']);
  const cCompSub = colOf(['ClientCompletionSubject']);
  const cCompBody = colOf(['ClientCompletionBody']);

  const cCcAssignee = colOf(['CcAssigneeOnClientStart (true/false)','CcAssigneeOnClientStart']);
  const cCcManager = colOf(['CcManagerOnClientStart (true/false)','CcManagerOnClientStart']);

  if (!cTitle || !cClient || !cDue) {
    return json(event, 400, { ok:false, error:'Missing required columns. Need Title, Client, DueDate (DD-MM-YYYY)' });
  }

  const window = await getCalendarWindow();

  let created = 0;
  const startRow = 2;
  const lastRow = ws.lastRow ? ws.lastRow.number : 1;

  for (let r = startRow; r <= lastRow; r++) {
    const row = ws.getRow(r);
    const title = getCellString(row.getCell(cTitle));
    const clientName = getCellString(row.getCell(cClient));
    const dueDmy = parseDueDateDmyFromExcel(row.getCell(cDue));

    if (!title || !clientName || !dueDmy) continue;

    const dueBaseYmd = dmyToYmd(dueDmy);
    const dueBase = dateFromYmdIST(dueBaseYmd);

    const recurrence = normalizeRecurrence(getCellString(row.getCell(cRec)) || 'AD_HOC');
    const generateCount = Math.max(1, parseInt(getCellString(row.getCell(cGen)) || '1', 10));
    const triggerDaysBefore = Math.max(0, parseInt(getCellString(row.getCell(cTrig)) || '15', 10));

    const category = normalizeCategory(getCellString(row.getCell(cCategory)) || 'OTHER');
    const type = getCellString(row.getCell(cType)) || 'FILING'; // free text

    const assignedEmail = (getCellString(row.getCell(cAssigned)) || '').trim() || null;
    const clientEmail = (getCellString(row.getCell(cClientEmail)) || '').trim() || null;

    const clientToEmails = asEmailList(getCellString(row.getCell(cTo)) || clientEmail || null);
    const clientCcEmails = asEmailList(getCellString(row.getCell(cCc)) || null);
    const clientBccEmails = asEmailList(getCellString(row.getCell(cBcc)) || null);

    const clientStartSubject = getCellString(row.getCell(cStartSub)) || '';
    const clientStartBody = getCellString(row.getCell(cStartBody)) || '';

    const sendClientCompletionMail = (cSendComp ? truthy(getCellString(row.getCell(cSendComp)) || 'true') : true);
    const clientCompletionSubject = getCellString(row.getCell(cCompSub)) || '';
    const clientCompletionBody = getCellString(row.getCell(cCompBody)) || '';

    const ccAssigneeOnClientStart = cCcAssignee ? truthy(getCellString(row.getCell(cCcAssignee)) || 'false') : false;
    const ccManagerOnClientStart = cCcManager ? truthy(getCellString(row.getCell(cCcManager)) || 'false') : false;

    const clientId = await findOrCreateClientByName(clientName);

    // If template has client email, set as primary if empty
    if (clientEmail) {
      const cRef = db().collection('clients').doc(clientId);
      const cSnap = await cRef.get();
      if (cSnap.exists && !cSnap.data().primaryEmail) {
        await cRef.update({ primaryEmail: clientEmail });
      }
    }

    const assignedToUid = (await findUserUidByEmail(assignedEmail)) || user.uid;

    const isSeries = recurrence !== 'AD_HOC' && generateCount > 1;
    const seriesId = isSeries ? db().collection('taskSeries').doc().id : null;

    for (let i = 0; i < generateCount; i++) {
      const dueDate = addInterval(dueBase, recurrence, i);
      const dueDateYmd = ymdIST(dueDate);

      const startDate = addDays(dateFromYmdIST(dueDateYmd), -triggerDaysBefore);
      const startDateYmd = ymdIST(startDate);

      const ev = await createStartCalendarEvent({
        title, clientId, startDateYmd, dueDateYmd, window
      });

      const tRef = db().collection('tasks').doc();
      await tRef.set({
        clientId,
        title,
        category,
        type,
        recurrence,

        seriesId,
        occurrenceIndex: i + 1,
        occurrenceTotal: generateCount,

        dueDate: admin.firestore.Timestamp.fromDate(dateFromYmdIST(dueDateYmd)),
        dueDateYmd,

        triggerDaysBefore,
        startDate: admin.firestore.Timestamp.fromDate(dateFromYmdIST(startDateYmd)),
        startDateYmd,

        assignedToUid,
        assignedToEmail: assignedEmail || user.email,

        status: 'PENDING',
        statusNote: '',
        delayReason: null,
        delayNotes: '',

        // Single calendar event
        calendarEventId: ev.calendarEventId,
        calendarHtmlLink: ev.calendarHtmlLink || null,
        calendarStartEventId: ev.calendarEventId,
        calendarDueEventId: null,

        // Per-task recipient controls
        clientToEmails,
        clientCcEmails,
        clientBccEmails,
        ccAssigneeOnClientStart,
        ccManagerOnClientStart,

        // Templates
        clientStartSubject,
        clientStartBody,

        clientStartMailSent: false,
        clientStartMailSentAt: null,
        clientStartGmailThreadId: null,
        clientStartGmailId: null,
        clientStartRfcMessageId: null,

        sendClientCompletionMail,
        clientCompletionSubject,
        clientCompletionBody,

        createdByUid: user.uid,
        createdAt: admin.firestore.FieldValue.serverTimestamp(),
        updatedAt: admin.firestore.FieldValue.serverTimestamp(),

        completedRequestedAt: null,
        completedAt: null,

        attachments: []
      });

      await auditLog({
        taskId: tRef.id,
        action: 'TASK_CREATED',
        actorUid: user.uid,
        actorEmail: user.email,
        details: { source:'XLSX', seriesId, occurrenceIndex: i+1, startDateYmd, dueDateYmd }
      });

      created++;
    }
  }

  return json(event, 200, { ok:true, created });
});