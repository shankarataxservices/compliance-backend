const { withCors, json, db, admin, calendar, ymd, addInterval, addDays, auditLog } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');
const { parse } = require('csv-parse/sync');

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
  const allowed = ['AD_HOC','WEEKLY','MONTHLY','QUARTERLY','HALF_YEARLY','YEARLY'];
  return allowed.includes(r) ? r : 'AD_HOC';
}

async function createCalendarEventForTask(taskDoc) {
  const cal = calendar();
  const endDateYmd = ymd(addDays(new Date(taskDoc.dueDateYmd), 1));

  const res = await cal.events.insert({
    calendarId: 'primary',
    sendUpdates: 'none', // IMPORTANT: no guest invites/emails
    requestBody: {
      summary: `${taskDoc.title} (Due ${taskDoc.dueDateYmd})`,
      description:
        `ClientId: ${taskDoc.clientId}\n` +
        `Status: ${taskDoc.status}\n` +
        `Start: ${taskDoc.startDateYmd}\n` +
        `Due: ${taskDoc.dueDateYmd}`,
      start: { date: taskDoc.startDateYmd },
      end: { date: endDateYmd }
    }
  });

  return res.data.id;
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');
  const csvText = body.csvText;
  if (!csvText) return json(event, 400, { ok:false, error:'csvText required' });

  const records = parse(csvText, { columns: true, skip_empty_lines: true, trim: true });

  let created = 0;

  for (const r of records) {
    const title = r.Title;
    const clientName = r.Client;
    const dueDateBase = new Date(r.DueDate); // YYYY-MM-DD
    const recurrence = normalizeRecurrence(r.Recurrence);
    const generateCount = parseInt(r.GenerateCount || '1', 10);
    const triggerDaysBefore = parseInt(r.TriggerDays || '15', 10);
    const type = (r.Type || 'FILING').toUpperCase();
    const category = (r.Category || 'OTHER').toUpperCase();

    const assignedEmail = (r.AssignedToEmail || '').trim() || null;
    const clientEmail = (r.ClientEmail || '').trim() || null;

    const clientStartSubject = (r.ClientStartSubject || '').trim();
    const clientStartBody = (r.ClientStartBody || '').trim();

    if (!title || !clientName || !r.DueDate) continue;

    const clientId = await findOrCreateClientByName(clientName);

    // If CSV has client email, update client.primaryEmail if empty
    if (clientEmail) {
      const cRef = db().collection('clients').doc(clientId);
      const cSnap = await cRef.get();
      if (cSnap.exists && !cSnap.data().primaryEmail) {
        await cRef.update({ primaryEmail: clientEmail });
      }
    }

    const assignedToUid = (await findUserUidByEmail(assignedEmail)) || user.uid;

    for (let i = 0; i < generateCount; i++) {
      const dueDate = addInterval(dueDateBase, recurrence, i);
      const dueDateYmd = ymd(dueDate);
      const startDate = addDays(dueDate, -triggerDaysBefore);
      const startDateYmd = ymd(startDate);

      const tRef = db().collection('tasks').doc();

      const taskDoc = {
        clientId,
        title,
        category,
        type,
        recurrence,

        dueDate: admin.firestore.Timestamp.fromDate(dueDate),
        dueDateYmd,

        triggerDaysBefore,

        startDate: admin.firestore.Timestamp.fromDate(startDate),
        startDateYmd,

        assignedToUid,
        assignedToEmail: assignedEmail || user.email,

        status: 'PENDING',
        statusNote: '',
        delayReason: null,
        delayNotes: '',

        calendarEventId: null,

        // Client start mail (send on start date)
        clientStartSubject,
        clientStartBody,
        clientStartMailSent: false,
        clientStartMailSentAt: null,

        createdByUid: user.uid,
        createdAt: admin.firestore.FieldValue.serverTimestamp(),
        updatedAt: admin.firestore.FieldValue.serverTimestamp(),

        completedRequestedAt: null,
        completedAt: null,

        attachments: []
      };

      const calendarEventId = await createCalendarEventForTask(taskDoc);
      taskDoc.calendarEventId = calendarEventId;

      await tRef.set(taskDoc);

      await auditLog({
        taskId: tRef.id,
        action: 'TASK_CREATED',
        actorUid: user.uid,
        actorEmail: user.email,
        details: { source: 'CSV', recurrence, dueDateYmd, startDateYmd }
      });

      created++;
    }
  }

  return json(event, 200, { ok:true, created });
});