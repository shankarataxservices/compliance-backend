// netlify/functions/tasks_updatetask.js
const {
  withCors, json, db, admin,
  calendar, ymdIST, addDays, dateFromYmdIST,
  getCalendarWindow, calTimeRange,
  auditLog
} = require('./_common');
const { requireUser } = require('./_auth');

async function findUserUidByEmail(email) {
  if (!email) return null;
  const e = String(email).trim().toLowerCase();
  let snap = await db().collection('users').where('emailLower', '==', e).limit(1).get();
  if (snap.empty) snap = await db().collection('users').where('email', '==', email).limit(1).get();
  if (snap.empty) return null;
  return snap.docs[0].id;
}

function normalizeCategory(x) {
  const raw = String(x || 'OTHER').trim();
  const u = raw.toUpperCase().replace(/\s+/g, '_');
  if (u === 'ITR' || u === 'INCOME_TAX' || raw.toLowerCase() === 'income tax') return 'INCOME_TAX';
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
function asEmailListLoose(x) {
  if (Array.isArray(x)) return x.map(s => String(s).trim()).filter(Boolean);
  if (typeof x === 'string') return x.split(/[;,:]/).map(s => s.trim()).filter(Boolean);
  return [];
}

async function patchEvent({ eventId, whenYmd, summary, description, colorId=null, window }) {
  if (!eventId) return;
  const cal = calendar();
  const range = calTimeRange(whenYmd, window.startHH, window.endHH, window.timeZone);
  await cal.events.patch({
    calendarId: 'primary',
    eventId,
    sendUpdates: 'none',
    requestBody: {
      summary,
      description,
      ...range,
      ...(colorId ? { colorId } : {})
    }
  });
}
function completedPrefix(status) {
  return status === 'COMPLETED' ? '[COMPLETED] ' : '';
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const role = String(user.role || 'ASSOCIATE').toUpperCase().trim();
  const canEdit = (role === 'PARTNER' || role === 'MANAGER');
  if (!canEdit) return json(event, 403, { ok:false, error:'Partner/Manager only' });

  const body = JSON.parse(event.body || '{}');
  const taskId = String(body.taskId || '').trim();
  if (!taskId) return json(event, 400, { ok:false, error:'taskId required' });

  const baseRef = db().collection('tasks').doc(taskId);
  const baseSnap = await baseRef.get();
  if (!baseSnap.exists) return json(event, 404, { ok:false, error:'Task not found' });
  const base = baseSnap.data();

  const applyToSeries = !!body.applyToSeries && !!base.seriesId;

  // Editable fields
  const newTitle = String(body.title ?? base.title ?? '').trim() || base.title || 'Untitled';
  const newCategory = normalizeCategory(body.category ?? base.category ?? 'OTHER');
  const newType = String(body.type ?? base.type ?? 'FILING').trim();
  const newPriority = normalizePriority(body.priority ?? base.priority ?? 'MEDIUM');

  const newTrigger = body.triggerDaysBefore !== undefined
    ? Math.max(0, parseInt(body.triggerDaysBefore, 10))
    : (base.triggerDaysBefore ?? 15);

  const newAssignedEmail = String(body.assignedToEmail ?? base.assignedToEmail ?? '').trim();
  const newAssignedUid = newAssignedEmail
    ? ((await findUserUidByEmail(newAssignedEmail)) || base.assignedToUid)
    : base.assignedToUid;

  // Snooze
  const snoozedUntilYmd = (body.snoozedUntilYmd === null || body.snoozedUntilYmd === undefined)
    ? (base.snoozedUntilYmd || null)
    : String(body.snoozedUntilYmd || '').trim() || null;
  if (snoozedUntilYmd && !/^\d{4}-\d{2}-\d{2}$/.test(snoozedUntilYmd)) {
    return json(event, 400, { ok:false, error:'snoozedUntilYmd must be YYYY-MM-DD or null' });
  }

  // Due date change allowed ONLY for single occurrence
  const newDueDateYmd = (!applyToSeries && body.dueDateYmd)
    ? String(body.dueDateYmd).trim()
    : base.dueDateYmd;

  // ===== Start mail controls + recipients =====
  const sendClientStartMail =
    (body.sendClientStartMail === undefined)
      ? (base.sendClientStartMail !== false)
      : (body.sendClientStartMail !== false);

  const clientToEmails = asEmailListLoose(body.clientToEmails ?? base.clientToEmails ?? []);
  const clientCcEmails = asEmailListLoose(body.clientCcEmails ?? base.clientCcEmails ?? []);
  const clientBccEmails = asEmailListLoose(body.clientBccEmails ?? base.clientBccEmails ?? []);

  const ccAssigneeOnClientStart = body.ccAssigneeOnClientStart === true ? true :
    (body.ccAssigneeOnClientStart === false ? false : !!base.ccAssigneeOnClientStart);

  const ccManagerOnClientStart = body.ccManagerOnClientStart === true ? true :
    (body.ccManagerOnClientStart === false ? false : !!base.ccManagerOnClientStart);

  const newClientStartSubject = String(body.clientStartSubject ?? base.clientStartSubject ?? '');
  const newClientStartBody = String(body.clientStartBody ?? base.clientStartBody ?? '');

  // ===== Completion mail controls + overrides =====
  const sendClientCompletionMail =
    (body.sendClientCompletionMail === undefined)
      ? (base.sendClientCompletionMail !== false)
      : (body.sendClientCompletionMail !== false);

  const newClientCompletionSubject = String(body.clientCompletionSubject ?? base.clientCompletionSubject ?? '');
  const newClientCompletionBody = String(body.clientCompletionBody ?? base.clientCompletionBody ?? '');

  const completionToEmails = asEmailListLoose(body.completionToEmails ?? base.completionToEmails ?? []);
  const completionCcEmails = asEmailListLoose(body.completionCcEmails ?? base.completionCcEmails ?? []);
  const completionBccEmails = asEmailListLoose(body.completionBccEmails ?? base.completionBccEmails ?? []);

  const ccAssigneeOnCompletion = body.ccAssigneeOnCompletion === true ? true :
    (body.ccAssigneeOnCompletion === false ? false : !!base.ccAssigneeOnCompletion);

  const ccManagerOnCompletion = body.ccManagerOnCompletion === true ? true :
    (body.ccManagerOnCompletion === false ? false : !!base.ccManagerOnCompletion);

  // Targets
  let targets = [];
  if (applyToSeries) {
    const snap = await db().collection('tasks').where('seriesId', '==', base.seriesId).get();
    targets = snap.docs.map(d => ({ id: d.id, ref: d.ref, data: d.data() }));
  } else {
    targets = [{ id: taskId, ref: baseRef, data: base }];
  }

  const window = await getCalendarWindow();
  let updatedCount = 0;

  for (const t of targets) {
    const old = t.data;
    const dueYmd = (t.id === taskId) ? newDueDateYmd : old.dueDateYmd;

    if (!/^\d{4}-\d{2}-\d{2}$/.test(dueYmd)) {
      return json(event, 400, { ok:false, error:`Invalid dueDateYmd for task ${t.id}` });
    }

    const startYmd = ymdIST(addDays(dateFromYmdIST(dueYmd), -newTrigger));

    const updateDoc = {
      title: newTitle,
      category: newCategory,
      type: newType,
      priority: newPriority,

      triggerDaysBefore: newTrigger,

      assignedToEmail: newAssignedEmail || old.assignedToEmail || '',
      assignedToUid: newAssignedUid || old.assignedToUid || null,

      snoozedUntilYmd: snoozedUntilYmd || null,

      startDateYmd: startYmd,
      startDate: admin.firestore.Timestamp.fromDate(dateFromYmdIST(startYmd)),

      // Start mail
      sendClientStartMail,
      clientToEmails,
      clientCcEmails,
      clientBccEmails,
      ccAssigneeOnClientStart,
      ccManagerOnClientStart,
      clientStartSubject: newClientStartSubject,
      clientStartBody: newClientStartBody,

      // Completion mail + overrides
      sendClientCompletionMail,
      clientCompletionSubject: newClientCompletionSubject,
      clientCompletionBody: newClientCompletionBody,

      completionToEmails,
      completionCcEmails,
      completionBccEmails,
      ccAssigneeOnCompletion,
      ccManagerOnCompletion,

      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    };

    if (t.id === taskId) {
      updateDoc.dueDateYmd = dueYmd;
      updateDoc.dueDate = admin.firestore.Timestamp.fromDate(dateFromYmdIST(dueYmd));
    }

    await t.ref.update(updateDoc);

    // Patch calendar event
    const prefix = completedPrefix(old.status);
    const cn = String(old.clientNameSnapshot || '').trim() || String(old.clientId || '').trim();
const extra = String(old.calendarDescription || '').trim();
const descBase =
  `Client: ${cn}\n` +
  `Start: ${startYmd}\n` +
  `Due: ${dueYmd}\n`;
const desc = extra ? `${descBase}\n${extra}` : descBase;

    const eventId = old.calendarEventId || old.calendarStartEventId || null;
    await patchEvent({
      eventId,
      whenYmd: startYmd,
      summary: `${prefix}START: ${newTitle}`,
      description: desc,
      colorId: (old.status === 'COMPLETED') ? '2' : null,
      window
    });

    // Back-compat: patch old due event if exists
    if (old.calendarDueEventId) {
      await patchEvent({
        eventId: old.calendarDueEventId,
        whenYmd: dueYmd,
        summary: `${prefix}DUE: ${newTitle}`,
        description: desc,
        colorId: (old.status === 'COMPLETED') ? '2' : null,
        window
      });
    }

    await auditLog({
      taskId: t.id,
      action: 'TASK_EDITED',
      actorUid: user.uid,
      actorEmail: user.email,
      details: { applyToSeries, baseTaskId: taskId }
    });

    updatedCount++;
  }

  return json(event, 200, {
    ok:true,
    updatedCount,
    applyToSeries,
    seriesId: base.seriesId || null
  });
});