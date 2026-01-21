const { withCors, json, db, admin, calendar, ymd, addDays, auditLog } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

async function findOrCreateClientByIdOrName({ clientId, clientName, clientEmail }) {
  if (clientId) {
    const cRef = db().collection('clients').doc(clientId);
    const cSnap = await cRef.get();
    if (!cSnap.exists) throw new Error('Client not found: ' + clientId);

    if (clientEmail && !cSnap.data().primaryEmail) {
      await cRef.update({ primaryEmail: clientEmail });
    }
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
  const snap = await db().collection('users').where('email', '==', email).limit(1).get();
  if (snap.empty) return null;
  return snap.docs[0].id;
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');

  const clientId = await findOrCreateClientByIdOrName({
    clientId: body.clientId || null,
    clientName: body.clientName || null,
    clientEmail: body.clientEmail || null,
  });

  const dueDateYmd = body.dueDateYmd;
  if (!dueDateYmd) return json(event, 400, { ok:false, error:'dueDateYmd required' });

  const triggerDaysBefore = Number(body.triggerDaysBefore ?? 15);
  const startDateYmd = ymd(addDays(new Date(dueDateYmd), -triggerDaysBefore));
  const endDateYmd = ymd(addDays(new Date(dueDateYmd), 1));

  const assignedToEmail = body.assignedToEmail || user.email;
  const assignedToUid = (await findUserUidByEmail(assignedToEmail)) || user.uid;

  const taskDoc = {
    clientId,
    title: body.title || 'Untitled',
    category: (body.category || 'OTHER').toUpperCase(),
    type: (body.type || 'FILING').toUpperCase(),
    recurrence: 'AD_HOC',

    dueDateYmd,
    startDateYmd,
    triggerDaysBefore,

    status: 'PENDING',
    statusNote: '',
    delayReason: null,
    delayNotes: '',

    assignedToUid,
    assignedToEmail,

    // Client start mail
    clientStartSubject: body.clientStartSubject || '',
    clientStartBody: body.clientStartBody || '',
    clientStartMailSent: false,
    clientStartMailSentAt: null,

    calendarEventId: null,

    createdByUid: user.uid,
    createdAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),

    completedRequestedAt: null,
    completedAt: null,

    attachments: []
  };

  // Calendar event spanning start -> due+1
  const created = await calendar().events.insert({
    calendarId: 'primary',
    sendUpdates: 'none',
    requestBody: {
      summary: `${taskDoc.title} (Due ${dueDateYmd})`,
      start: { date: startDateYmd },
      end: { date: endDateYmd },
      description: `ClientId: ${clientId}\nStatus: ${taskDoc.status}\nStart: ${startDateYmd}\nDue: ${dueDateYmd}`
    }
  });

  taskDoc.calendarEventId = created.data.id;

  const tRef = db().collection('tasks').doc();
  await tRef.set(taskDoc);

  await auditLog({ taskId: tRef.id, action:'TASK_CREATED', actorUid:user.uid, actorEmail:user.email, details:{ source:'UI_CREATE_ONE' } });

  return json(event, 200, { ok:true, taskId: tRef.id });
});