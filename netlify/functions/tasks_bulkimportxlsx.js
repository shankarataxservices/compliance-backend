// netlify/functions/tasks_bulkimportxlsx.js
const Busboy = require('busboy');
const ExcelJS = require('exceljs');
const crypto = require('crypto');

const {
  withCors, json, db, admin,
  dmyToYmd, ymdIST, dateFromYmdIST,
  addInterval, addDays,
  getCalendarWindow,
  auditLog, asEmailList,
  createStartCalendarEvent, trySendStartMailImmediately
} = require('./_common');
const { requireUser } = require('./_auth');

// ===== Worker(import) password verify (PBKDF2) =====
function pbkdf2Hash(password, saltBuf, iter) {
  const keyLen = 32;
  const digest = 'sha256';
  return crypto.pbkdf2Sync(password, saltBuf, iter, keyLen, digest);
}
async function verifyWorkerImportPassword(plainPassword) {
  const sSnap = await db().collection('settings').doc('security').get();
  if (!sSnap.exists) return false;
  const s = sSnap.data() || {};
  if (s.workerImportEnabled !== true) return false;
  if (!s.workerImportHash || !s.workerImportSalt || !s.workerImportIter) return false;

  const salt = Buffer.from(String(s.workerImportSalt), 'base64');
  const iter = Number(s.workerImportIter || 0);
  const expected = Buffer.from(String(s.workerImportHash), 'base64');

  if (!salt.length || !iter || !expected.length) return false;

  const got = pbkdf2Hash(String(plainPassword || ''), salt, iter);
  if (got.length !== expected.length) return false;
  return crypto.timingSafeEqual(got, expected);
}

// ===== Helpers =====
async function findOrCreateClientByName(clientName) {
  const snap = await db().collection('clients').where('name', '==', clientName).limit(1).get();
  if (!snap.empty) return { id: snap.docs[0].id, data: snap.docs[0].data() };

  const ref = db().collection('clients').doc();
  const newData = {
    name: clientName,
    pan: '', gstin: '', cin: '',
    assessmentYear: '', engagementType: '',
    primaryEmail: '',
    ccEmails: [], bccEmails: [],
    driveFolderId: null,
    createdAt: admin.firestore.FieldValue.serverTimestamp()
  };
  await ref.set(newData);
  return { id: ref.id, data: newData };
}

async function findUserUidByEmail(email) {
  if (!email) return null;
  const e = String(email).trim().toLowerCase();
  let snap = await db().collection('users').where('emailLower', '==', e).limit(1).get();
  if (snap.empty) snap = await db().collection('users').where('email', '==', email).limit(1).get();
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
  if (String(x).trim().toLowerCase() === 'income tax') return 'INCOME_TAX';
  if (u === 'GST') return 'GST';
  if (u === 'TDS') return 'TDS';
  if (u === 'ROC') return 'ROC';
  if (u === 'ACCOUNTING') return 'ACCOUNTING';
  if (u === 'AUDIT') return 'AUDIT';
  return 'OTHER';
}
function normalizePriority(x) {
  const v = String(x || 'MEDIUM').trim().toUpperCase();
  if (v === 'HIGH' || v === 'LOW') return v;
  return 'MEDIUM';
}
function truthy(x) {
  const s = String(x ?? '').trim().toLowerCase();
  if (!s) return false;
  return ['1','true','yes','y'].includes(s);
}
function getCellString(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (v == null) return '';
  if (typeof v === 'string') return v.trim();
  if (typeof v === 'number') return String(v);
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'object') {
    if (v.text) return String(v.text).trim();
    if (v.richText && Array.isArray(v.richText)) return v.richText.map(r => r.text || '').join('').trim();
  }
  return String(v).trim();
}
function parseDueDateDmyFromExcel(cell) {
  if (!cell || cell.value == null) return null;
  const v = cell.value;
  if (v instanceof Date) {
    const yyyy = v.getFullYear();
    const mm = String(v.getMonth() + 1).padStart(2,'0');
    const dd = String(v.getDate()).padStart(2,'0');
    return `${dd}-${mm}-${yyyy}`;
  }
  const s = getCellString(cell);
  return s || null;
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  // Parse multipart
  const busboy = Busboy({ headers: event.headers });
  let fileBuffer = Buffer.alloc(0);
  let importPassword = '';

  busboy.on('field', (name, val) => {
    if (name === 'importPassword') importPassword = String(val || '');
  });
  busboy.on('file', (name, file) => {
    file.on('data', (d) => { fileBuffer = Buffer.concat([fileBuffer, d]); });
  });

  const done = new Promise((resolve, reject) => {
    busboy.on('finish', resolve);
    busboy.on('error', reject);
  });
  const bodyBuf = event.isBase64Encoded ? Buffer.from(event.body, 'base64') : Buffer.from(event.body || '');
  busboy.end(bodyBuf);
  await done;

  if (!fileBuffer.length) return json(event, 400, { ok:false, error:'XLSX file missing' });

  // Authorization:
  // - PARTNER/MANAGER can import
  // - ASSOCIATE can import only if importPassword matches settings/security
  const role = String(user.role || 'ASSOCIATE').toUpperCase().trim();
  const privileged = (role === 'PARTNER' || role === 'MANAGER');

  if (!privileged) {
    const okPwd = await verifyWorkerImportPassword(importPassword);
    if (!okPwd) {
      return json(event, 403, { ok:false, error:'Associate import password required or incorrect' });
    }
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(fileBuffer);
  const ws = wb.getWorksheet('Import') || wb.worksheets[0];
  if (!ws) return json(event, 400, { ok:false, error:'No worksheet found' });

  const headerRow = ws.getRow(1);
  const headers = {};
  headerRow.eachCell((cell, colNumber) => {
    const key = getCellString(cell);
    if (key) headers[key] = colNumber;
  });
  function colOf(names) {
    for (const n of names) { if (headers[n]) return headers[n]; }
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
  const cPriority = colOf(['Priority']);

  const cTo = colOf(['ClientTo (emails ; , : separated)','ClientTo']);
  const cCc = colOf(['ClientCC (emails ; , : separated)','ClientCC']);
  const cBcc = colOf(['ClientBCC (emails ; , : separated)','ClientBCC']);
  const cStartSub = colOf(['ClientStartSubject']);
  const cStartBody = colOf(['ClientStartBody']);
  const cSendStart = colOf(['SendStartMail (true/false)','SendStartMail','SendClientStartMail (true/false)','SendClientStartMail']);
  const cCcAssignee = colOf(['CcAssigneeOnClientStart (true/false)','CcAssigneeOnClientStart']);
  const cCcManager = colOf(['CcManagerOnClientStart (true/false)','CcManagerOnClientStart']);

  const cSendComp = colOf(['SendClientCompletionMail (true/false)','SendClientCompletionMail']);
  const cCompSub = colOf(['ClientCompletionSubject']);
  const cCompBody = colOf(['ClientCompletionBody']);
  
  // NEW: Calendar description for Google Calendar event
  const cCalDesc = colOf(['GoogleCalendarDescription','CalendarDescription']);
  
  // NEW: completion overrides
  const cCompTo = colOf(['CompletionTo (emails ; , : separated)','CompletionTo']);
  const cCompCc = colOf(['CompletionCC (emails ; , : separated)','CompletionCC']);
  const cCompBcc = colOf(['CompletionBCC (emails ; , : separated)','CompletionBCC']);
  const cCcAssigneeComp = colOf(['CcAssigneeOnCompletion (true/false)','CcAssigneeOnCompletion']);
  const cCcManagerComp = colOf(['CcManagerOnCompletion (true/false)','CcManagerOnCompletion']);

  if (!cTitle || !cClient || !cDue) {
    return json(event, 400, { ok:false, error:'Missing required columns (Title, Client, DueDate)' });
  }

  const window = await getCalendarWindow();
  const todayYmd = ymdIST(new Date());

  let created = 0;
  const seriesIdsCreated = new Set();

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
    const type = getCellString(row.getCell(cType)) || 'FILING';
    const priority = cPriority ? normalizePriority(getCellString(row.getCell(cPriority))) : 'MEDIUM';

    const assignedEmail = (getCellString(row.getCell(cAssigned)) || '').trim() || null;
    const clientEmail = (getCellString(row.getCell(cClientEmail)) || '').trim() || null;

    const clientToEmails = asEmailList(getCellString(row.getCell(cTo)) || clientEmail || null);
    const clientCcEmails = asEmailList(getCellString(row.getCell(cCc)) || null);
    const clientBccEmails = asEmailList(getCellString(row.getCell(cBcc)) || null);

    const clientStartSubject = getCellString(row.getCell(cStartSub)) || '';
    const clientStartBody = getCellString(row.getCell(cStartBody)) || '';

    const sendClientStartMail = (cSendStart ? truthy(getCellString(row.getCell(cSendStart)) || 'true') : true);
    const ccAssigneeOnClientStart = cCcAssignee ? truthy(getCellString(row.getCell(cCcAssignee)) || 'false') : false;
    const ccManagerOnClientStart = cCcManager ? truthy(getCellString(row.getCell(cCcManager)) || 'false') : false;

    const sendClientCompletionMail = (cSendComp ? truthy(getCellString(row.getCell(cSendComp)) || 'true') : true);
    const clientCompletionSubject = getCellString(row.getCell(cCompSub)) || '';
    const clientCompletionBody = getCellString(row.getCell(cCompBody)) || '';
    
    // NEW: calendar description from XLSX
    const calendarDescription = cCalDesc ? (getCellString(row.getCell(cCalDesc)) || '') : '';
    
    const completionToEmails = asEmailList(getCellString(row.getCell(cCompTo)) || null);
    const completionCcEmails = asEmailList(getCellString(row.getCell(cCompCc)) || null);
    const completionBccEmails = asEmailList(getCellString(row.getCell(cCompBcc)) || null);
    const ccAssigneeOnCompletion = cCcAssigneeComp ? truthy(getCellString(row.getCell(cCcAssigneeComp)) || 'false') : false;
    const ccManagerOnCompletion = cCcManagerComp ? truthy(getCellString(row.getCell(cCcManagerComp)) || 'false') : false;

    // Client get/create
    const { id: clientId, data: clientData0 } = await findOrCreateClientByName(clientName);
    const clientData = { ...(clientData0 || {}) };

    if (clientEmail && !clientData.primaryEmail) {
      await db().collection('clients').doc(clientId).update({ primaryEmail: clientEmail });
      clientData.primaryEmail = clientEmail;
    }

    // Assignee rules
    let assignedToUid = null;
    let assignedToEmailFinal = assignedEmail || user.email;

    if (!privileged) {
      // Associate import: force to self
      assignedToUid = user.uid;
      assignedToEmailFinal = user.email;
    } else {
      assignedToUid = (await findUserUidByEmail(assignedEmail)) || user.uid;
      if (!assignedToUid) assignedToUid = user.uid;
      if (!assignedEmail) assignedToEmailFinal = user.email;
    }

    const isSeries = recurrence !== 'AD_HOC' && generateCount > 1;
    const seriesId = isSeries ? db().collection('taskSeries').doc().id : null;
    if (seriesId) seriesIdsCreated.add(seriesId);

    for (let i = 0; i < generateCount; i++) {
      const dueDate = addInterval(dueBase, recurrence, i);
      const dueDateYmd = ymdIST(dueDate);
      const startDate = addDays(dateFromYmdIST(dueDateYmd), -triggerDaysBefore);
      const startDateYmd = ymdIST(startDate);

      const ev = await createStartCalendarEvent({
  title,
  clientId,
  clientName: clientData.name || clientName,
  startDateYmd,
  dueDateYmd,
  window,
  calendarDescription
});

      // immediate mail if start is today
      const taskMailObj = {
        title,
        startDateYmd,
        dueDateYmd,
        sendClientStartMail,
        clientStartSubject,
        clientStartBody,
        clientToEmails, clientCcEmails, clientBccEmails,
        ccAssigneeOnClientStart,
        ccManagerOnClientStart,
        assignedToEmail: assignedToEmailFinal,
        assignedToUid
      };

      let mailResult = null;
      if (startDateYmd === todayYmd) {
        mailResult = await trySendStartMailImmediately({
          task: taskMailObj,
          client: clientData,
          window
        });
      }

      const tRef = db().collection('tasks').doc();
      await tRef.set({
        clientId,
        clientNameSnapshot: clientData.name || clientName, // NEW
        calendarDescription: String(calendarDescription || '').trim(), // NEW
        title,
        category,
        type,
        priority,        recurrence,
        seriesId,
        occurrenceIndex: i + 1,
        occurrenceTotal: generateCount,

        dueDate: admin.firestore.Timestamp.fromDate(dateFromYmdIST(dueDateYmd)),
        dueDateYmd,
        triggerDaysBefore,

        startDate: admin.firestore.Timestamp.fromDate(dateFromYmdIST(startDateYmd)),
        startDateYmd,

        assignedToUid,
        assignedToEmail: assignedToEmailFinal,

        status: 'PENDING',
        statusNote: '',
        delayReason: null,
        delayNotes: '',
        snoozedUntilYmd: null,

        calendarEventId: ev.calendarEventId,
        calendarHtmlLink: ev.calendarHtmlLink || null,
        calendarStartEventId: ev.calendarEventId,
        calendarDueEventId: null,

        // Start mail fields
        sendClientStartMail,
        clientToEmails, clientCcEmails, clientBccEmails,
        ccAssigneeOnClientStart, ccManagerOnClientStart,
        clientStartSubject,
        clientStartBody,
        clientStartMailSent: mailResult ? true : false,
        clientStartMailSentAt: mailResult?.clientStartMailSentAt || null,
        clientStartGmailThreadId: mailResult?.clientStartGmailThreadId || null,
        clientStartGmailId: mailResult?.clientStartGmailId || null,
        clientStartRfcMessageId: mailResult?.clientStartRfcMessageId || null,
        clientStartReferences: mailResult?.clientStartReferences || null,

        // Completion mail fields + overrides
        sendClientCompletionMail,
        clientCompletionSubject,
        clientCompletionBody,
        completionToEmails,
        completionCcEmails,
        completionBccEmails,
        ccAssigneeOnCompletion,
        ccManagerOnCompletion,

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
        details: {
          source: 'XLSX',
          seriesId,
          occurrenceIndex: i + 1,
          startDateYmd,
          sentMailNow: !!mailResult
        }
      });

      created++;
    }
  }

  return json(event, 200, {
    ok:true,
    created,
    seriesCreated: seriesIdsCreated.size
  });
});