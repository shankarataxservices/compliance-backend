// netlify/functions/series_rebuild.js
const {
  withCors, json, db, admin,
  calendar, ymdIST, addDays, addInterval, dateFromYmdIST,
  getCalendarWindow, calTimeRange,
  auditLog
} = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

async function findUserUidByEmail(email) {
  if (!email) return null;
  const e = String(email).trim().toLowerCase();

  let snap = await db().collection('users').where('emailLower', '==', e).limit(1).get();
  if (snap.empty) snap = await db().collection('users').where('email', '==', email).limit(1).get();
  if (snap.empty) return null;

  return snap.docs[0].id;
}

// Use shared createStartCalendarEvent from _common.js (supports clientName + calendarDescription)

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');
  const { seriesId, addCount } = body;

  if (!seriesId) return json(event, 400, { ok:false, error:'seriesId required' });

  const n = Math.max(1, parseInt(addCount || '1', 10));

  // Load existing tasks in series
  const snap = await db().collection('tasks').where('seriesId', '==', seriesId).get();
  if (snap.empty) return json(event, 404, { ok:false, error:'No tasks found for seriesId' });

  const tasks = snap.docs.map(d => ({ id: d.id, ref: d.ref, data: d.data() }));
  const first = tasks[0].data;

// Load client name once (for Calendar description)
const cSnap = first.clientId ? await db().collection('clients').doc(first.clientId).get() : null;
const clientName = (cSnap && cSnap.exists) ? String((cSnap.data() || {}).name || '').trim() : '';
const calendarDescription = String(first.calendarDescription || '').trim();

  // Base due date = occurrenceIndex==1 dueDateYmd if exists; else min dueDateYmd
  let baseDueYmd = null;
  const occ1 = tasks.find(x => x.data.occurrenceIndex === 1);
  if (occ1) baseDueYmd = occ1.data.dueDateYmd;
  if (!baseDueYmd) baseDueYmd = tasks.map(x => x.data.dueDateYmd).sort()[0];

  const baseDueDate = dateFromYmdIST(baseDueYmd);

  // Determine existing max idx
  let maxIdx = 0;
  const idxSet = new Set();
  for (const t of tasks) {
    const idx = Number(t.data.occurrenceIndex || 0);
    if (idx > maxIdx) maxIdx = idx;
    if (idx) idxSet.add(idx);
  }

  const recurrence = first.recurrence || 'MONTHLY';
  const triggerDaysBefore = Number(first.triggerDaysBefore || 15);

  const assignedToEmail = body.assignedToEmail || first.assignedToEmail || user.email;
  const assignedToUid = (await findUserUidByEmail(assignedToEmail)) || first.assignedToUid || user.uid;

  const window = await getCalendarWindow();

  let created = 0;
  const startIdx = maxIdx + 1;
  const endIdx = maxIdx + n;

  for (let idx = startIdx; idx <= endIdx; idx++) {
    if (idxSet.has(idx)) continue;

    const dueDate = addInterval(baseDueDate, recurrence, idx - 1);
    const dueDateYmd = ymdIST(dueDate);
    const startDateYmd = ymdIST(addDays(dateFromYmdIST(dueDateYmd), -triggerDaysBefore));

    const ev = await createStartCalendarEvent({
     title: first.title,
     clientId: first.clientId,
     clientName: clientName || first.clientNameSnapshot || '',
     startDateYmd,
     dueDateYmd,
     window,
     calendarDescription
    });

    const tRef = db().collection('tasks').doc();
    await tRef.set({
     ...first,
     clientNameSnapshot: (clientName || first.clientNameSnapshot || ''),
     calendarDescription: String(calendarDescription || first.calendarDescription || '').trim(),
     occurrenceIndex: idx,
     occurrenceTotal: (first.occurrenceTotal || maxIdx) + n,

      dueDate: admin.firestore.Timestamp.fromDate(dateFromYmdIST(dueDateYmd)),
      dueDateYmd,

      startDate: admin.firestore.Timestamp.fromDate(dateFromYmdIST(startDateYmd)),
      startDateYmd,

      assignedToEmail,
      assignedToUid,

      calendarEventId: ev.calendarEventId,
      calendarHtmlLink: ev.calendarHtmlLink || null,
      calendarStartEventId: ev.calendarEventId,
      calendarDueEventId: null,

      clientStartMailSent: false,
      clientStartMailSentAt: null,
      clientStartGmailThreadId: null,
      clientStartGmailId: null,
      clientStartRfcMessageId: null,
      clientStartReferences: null,

      createdAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    await auditLog({
      taskId: tRef.id,
      action: 'SERIES_REBUILD_CREATED',
      actorUid: user.uid,
      actorEmail: user.email,
      details: { seriesId, occurrenceIndex: idx, startDateYmd, dueDateYmd }
    });

    created++;
  }

  // Update occurrenceTotal on all existing tasks to match new endIdx
  const newTotal = endIdx;
  const batch = db().batch();
  tasks.forEach(t => batch.update(t.ref, { occurrenceTotal: newTotal }));
  await batch.commit();

  return json(event, 200, { ok:true, created, newTotal });
});