// netlify/functions/tasks_bulkimportcsv.js
const {
  withCors, json, db, admin,
  calendar, ymdIST, dateFromYmdIST, addDays, addInterval,
  getCalendarWindow, calTimeRange,
  auditLog, asEmailList
} = require('./_common');
const { requireUser } = require('./_auth');

function roleOf(user) {
  let r = String(user?.role || 'ASSOCIATE').toUpperCase().trim();
  if (r === 'WORKER') r = 'ASSOCIATE'; // compat
  return r || 'ASSOCIATE';
}
function isPrivileged(role) {
  return role === 'PARTNER' || role === 'MANAGER';
}

async function findOrCreateClientByIdOrName({ clientId, clientName, clientEmail }) {
  if (clientId) {
    const cRef = db().collection('clients').doc(clientId);
    const cSnap = await cRef.get();
    if (!cSnap.exists) throw new Error('Client not found: ' + clientId);
    if (clientEmail && !cSnap.data().primaryEmail) await cRef.update({ primaryEmail: clientEmail });
    return clientId;
  }
  if (!clientName) throw new Error('clientId or clientName required');

  const snap = await db().collection('clients').where('name', '==', clientName).limit(1).get();
  if (!snap.empty) {
    const id = snap.docs[0].id;
    const cRef = db().collection('clients').doc(id);
    const cSnap = await cRef.get();
    if (clientEmail && cSnap.exists && !cSnap.data().primaryEmail) await cRef.update({ primaryEmail: clientEmail });
    return id;
  }

  const ref = db().collection('clients').doc();
  await ref.set({
    name: clientName,
    primaryEmail: clientEmail || '',
    ccEmails: [],
    bccEmails: [],
    createdAt: admin.firestore.FieldValue.serverTimestamp()
  });
  return ref.id;
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

// Use shared createStartCalendarEvent from _common.js (supports clientName + calendarDescription)

/**
 * This endpoint is a JSON bulk importer (legacy name).
 * It is kept for backward compatibility and now supports the new mail flags/fields too.
 *
 * Request body (example):
 * {
 *   clientId?, clientName?, clientEmail?,
 *   title, dueDateYmd,
 *   category, type, priority,
 *   recurrence, generateCount, triggerDaysBefore,
 *   assignedToEmail,
 *   sendClientStartMail, clientToEmails, clientCcEmails, clientBccEmails,
 *   ccAssigneeOnClientStart, ccManagerOnClientStart,
 *   clientStartSubject, clientStartBody,
 *   sendClientCompletionMail, clientCompletionSubject, clientCompletionBody,
 *   completionToEmails, completionCcEmails, completionBccEmails,
 *   ccAssigneeOnCompletion, ccManagerOnCompletion
 * }
 */
exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const role = roleOf(user);
  const privileged = isPrivileged(role);

  // For safety: only Partner/Manager should use this legacy JSON import
  if (!privileged) {
    return json(event, 403, { ok:false, error:'Partner/Manager only' });
  }

  const body = JSON.parse(event.body || '{}');

  const clientId = await findOrCreateClientByIdOrName({
    clientId: body.clientId || null,
    clientName: body.clientName || null,
    clientEmail: body.clientEmail || null,
  });

  const title = body.title || 'Untitled';

  const dueDateYmd = String(body.dueDateYmd || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dueDateYmd)) {
    return json(event, 400, { ok:false, error:'dueDateYmd required (YYYY-MM-DD)' });
  }

  const recurrence = normalizeRecurrence(body.recurrence || 'AD_HOC');
  const generateCount = Math.max(1, parseInt(body.generateCount || '1', 10));
  const triggerDaysBefore = Math.max(0, parseInt(body.triggerDaysBefore ?? 15, 10));

  const assignedToEmail = (body.assignedToEmail || user.email || '').trim();
  const assignedToUid = (await findUserUidByEmail(assignedToEmail)) || user.uid;

  const category = normalizeCategory(body.category || 'OTHER');
  const type = String(body.type || 'FILING').trim();
  const priority = normalizePriority(body.priority || 'MEDIUM');

  // Mail flags + templates + recipients
  const sendClientStartMail = (body.sendClientStartMail !== false);
  const clientToEmails = asEmailList(body.clientToEmails || body.clientTo || body.clientEmail || null);
  const clientCcEmails = asEmailList(body.clientCcEmails || body.clientCc || null);
  const clientBccEmails = asEmailList(body.clientBccEmails || body.clientBcc || null);
  const ccAssigneeOnClientStart = body.ccAssigneeOnClientStart === true;
  const ccManagerOnClientStart = body.ccManagerOnClientStart === true;
  const clientStartSubject = String(body.clientStartSubject || '').trim();
  const clientStartBody = String(body.clientStartBody || '').trim();

  const sendClientCompletionMail = body.sendClientCompletionMail !== false;
  const clientCompletionSubject = String(body.clientCompletionSubject || '').trim();
  const clientCompletionBody = String(body.clientCompletionBody || '').trim();

  const completionToEmails = asEmailList(body.completionToEmails || body.completionTo || null);
  const completionCcEmails = asEmailList(body.completionCcEmails || body.completionCc || null);
  const completionBccEmails = asEmailList(body.completionBccEmails || body.completionBcc || null);
  const ccAssigneeOnCompletion = body.ccAssigneeOnCompletion === true;
  const ccManagerOnCompletion = body.ccManagerOnCompletion === true;

  const isSeries = recurrence !== 'AD_HOC' && generateCount > 1;
  const seriesId = isSeries ? db().collection('taskSeries').doc().id : null;

  const window = await getCalendarWindow();
  let created = 0;

  for (let i = 0; i < generateCount; i++) {
    const dueDate = addInterval(dateFromYmdIST(dueDateYmd), recurrence, i);
    const dueYmd = ymdIST(dueDate);
    const startDate = addDays(dateFromYmdIST(dueYmd), -triggerDaysBefore);
    const startYmd = ymdIST(startDate);

    const cSnap = await db().collection('clients').doc(clientId).get();
    const cName = cSnap.exists ? String((cSnap.data() || {}).name || '').trim() : '';
    const calendarDescription = String(body.calendarDescription || body.googleCalendarDescription || '').trim();

    const ev = await createStartCalendarEvent({
     title,
     clientId,
     clientName: cName,
     startDateYmd: startYmd,
     dueDateYmd: dueYmd,
     window,
     calendarDescription
    });

    const tRef = db().collection('tasks').doc();
    await tRef.set({
      clientId,
      title,
      category,
      type,
      priority,

      recurrence,
      seriesId,
      occurrenceIndex: i + 1,
      occurrenceTotal: generateCount,

      dueDate: admin.firestore.Timestamp.fromDate(dateFromYmdIST(dueYmd)),
      dueDateYmd: dueYmd,
      triggerDaysBefore,

      startDate: admin.firestore.Timestamp.fromDate(dateFromYmdIST(startYmd)),
      startDateYmd: startYmd,

      assignedToUid,
      assignedToEmail,

      status: 'PENDING',
      statusNote: '',
      delayReason: null,
      delayNotes: '',
      snoozedUntilYmd: null,

      calendarEventId: ev.calendarEventId,
      calendarHtmlLink: ev.calendarHtmlLink || null,
      calendarStartEventId: ev.calendarEventId,
      calendarDueEventId: null,

      // Start mail
      sendClientStartMail,
      clientToEmails,
      clientCcEmails,
      clientBccEmails,
      ccAssigneeOnClientStart,
      ccManagerOnClientStart,
      clientStartSubject,
      clientStartBody,
      clientStartMailSent: false,
      clientStartMailSentAt: null,
      clientStartGmailThreadId: null,
      clientStartGmailId: null,
      clientStartRfcMessageId: null,
      clientStartReferences: null,

      // Completion mail + overrides
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
      details: { source:'JSON_IMPORT', seriesId, occurrenceIndex: i+1, startDateYmd: startYmd, dueDateYmd: dueYmd }
    });

    created++;
  }

  return json(event, 200, { ok:true, created, seriesId });
});