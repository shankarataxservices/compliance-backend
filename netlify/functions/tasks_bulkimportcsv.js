const { json, db, admin, calendar, ymd, addInterval, addDays, auditLog } = require('./_common');
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

async function createCalendarEventForTask(taskDoc) {
  const cal = calendar();
  // All-day event on due date (keeps calendar clean)
  const startDate = taskDoc.dueDateYmd;
  const endDate = ymd(addDays(new Date(taskDoc.dueDateYmd), 1));

  const res = await cal.events.insert({
    calendarId: 'primary',
    requestBody: {
      summary: `${taskDoc.title}`,
      description: `ClientId: ${taskDoc.clientId}\nStatus: ${taskDoc.status}\nDue: ${taskDoc.dueDateYmd}`,
      start: { date: startDate },
      end: { date: endDate }
    }
  });

  return res.data.id;
}

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: require('./_common').cors() };
  if (event.httpMethod !== 'POST') return json(405, { error: 'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');
  const csvText = body.csvText;
  if (!csvText) return json(400, { error: 'csvText required' });

  // CSV columns (recommended):
  // Title,Client,DueDate,Recurrence,GenerateCount,TriggerDays,Type,Category,AssignedToEmail,ClientEmail
  const records = parse(csvText, { columns: true, skip_empty_lines: true, trim: true });

  let created = 0;

  for (const r of records) {
    const title = r.Title;
    const clientName = r.Client;
    const dueDateBase = new Date(r.DueDate); // YYYY-MM-DD
    const recurrence = (r.Recurrence || 'AD_HOC').toUpperCase();
    const generateCount = parseInt(r.GenerateCount || '1', 10);
    const triggerDaysBefore = parseInt(r.TriggerDays || '15', 10);
    const type = (r.Type || 'FILING').toUpperCase();
    const category = (r.Category || 'OTHER').toUpperCase();

    const assignedEmail = r.AssignedToEmail || null;
    const clientEmail = r.ClientEmail || null;

    if (!title || !clientName || !r.DueDate) continue;

    const clientId = await findOrCreateClientByName(clientName);

    // If CSV included client email, update client primaryEmail if empty
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
        status: 'PENDING',
        delayReason: null,
        delayNotes: '',
        calendarEventId: null,
        createdByUid: user.uid,
        createdAt: admin.firestore.FieldValue.serverTimestamp(),
        updatedAt: admin.firestore.FieldValue.serverTimestamp(),
        completedRequestedAt: null,
        completedAt: null,
        escalation: { startSent: false, d7: false, d3: false, d0: false, overdue: false },
        attachments: []
      };

      // Create calendar event now
      const calendarEventId = await createCalendarEventForTask(taskDoc);
      taskDoc.calendarEventId = calendarEventId;

      await tRef.set(taskDoc);

      await auditLog({
        taskId: tRef.id,
        action: 'TASK_CREATED',
        actorUid: user.uid,
        actorEmail: user.email,
        details: { source: 'CSV', recurrence, dueDateYmd }
      });

      created++;
    }
  }

  return json(200, { ok: true, created });
};
