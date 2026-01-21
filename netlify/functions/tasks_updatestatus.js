const { withCors, json, db, admin, calendar, sendEmail, auditLog } = require('./_common');
const { requireUser } = require('./_auth');

function renderTemplate(str, vars) {
  return String(str || '')
    .replaceAll('{{clientName}}', vars.clientName || '')
    .replaceAll('{{taskTitle}}', vars.taskTitle || '')
    .replaceAll('{{startDate}}', vars.startDate || '')
    .replaceAll('{{dueDate}}', vars.dueDate || '');
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const body = JSON.parse(event.body || '{}');
  const { taskId, newStatus, statusNote, delayReason, delayNotes } = body;
  if (!taskId || !newStatus) return json(event, 400, { ok:false, error:'taskId,newStatus required' });

  const tRef = db().collection('tasks').doc(taskId);
  const tSnap = await tRef.get();
  if (!tSnap.exists) return json(event, 404, { ok:false, error:'Task not found' });

  const task = tSnap.data();
  const isPartner = user.role === 'PARTNER';
  const isAssignee = task.assignedToUid === user.uid;
  if (!isPartner && !isAssignee) return json(event, 403, { ok:false, error:'Not allowed' });

  if (!isPartner && newStatus === 'COMPLETED') {
    return json(event, 403, { ok:false, error:'Only partner can mark COMPLETED' });
  }

  const updates = {
    status: newStatus,
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
  };

  if (typeof statusNote === 'string') updates.statusNote = statusNote;
  if (delayReason) updates.delayReason = delayReason;
  if (typeof delayNotes === 'string') updates.delayNotes = delayNotes;

  if (newStatus === 'APPROVAL_PENDING') updates.completedRequestedAt = admin.firestore.FieldValue.serverTimestamp();
  if (newStatus === 'COMPLETED') updates.completedAt = admin.firestore.FieldValue.serverTimestamp();

  await tRef.update(updates);

  await auditLog({
    taskId,
    action: 'STATUS_CHANGE',
    actorUid: user.uid,
    actorEmail: user.email,
    details: { from: task.status, to: newStatus, statusNote: statusNote || '' }
  });

  // Only on COMPLETED: email client
  if (newStatus === 'COMPLETED') {
    // Patch calendar (no guest mails)
    if (task.calendarEventId) {
      try {
        await calendar().events.patch({
          calendarId: 'primary',
          eventId: task.calendarEventId,
          sendUpdates: 'none',
          requestBody: { summary: `[COMPLETED] ${task.title} (Due ${task.dueDateYmd})`, colorId: '2' }
        });
      } catch (e) {}
    }

    const cSnap = await db().collection('clients').doc(task.clientId).get();
    const client = cSnap.exists ? cSnap.data() : {};
    const to = client.primaryEmail ? [client.primaryEmail] : [];
    const cc = client.ccEmails || [];
    const bcc = client.bccEmails || [];

    if (to.length) {
      const html = `
        <p>Dear ${client.name || 'Client'},</p>
        <p>We have completed: <b>${task.title}</b></p>
        <p>Due date: <b>${task.dueDateYmd}</b></p>
        <p>Status note: ${updates.statusNote || task.statusNote || ''}</p>
        <p>Regards,<br>${process.env.MAIL_SIGNATURE || 'Compliance Team'}</p>
      `;

      await sendEmail({
        to, cc, bcc,
        subject: `Completed: ${task.title} (${client.name || ''})`,
        html
      });

      await auditLog({ taskId, action:'EMAIL_SENT', actorUid:null, actorEmail:null, details:{ type:'CLIENT_COMPLETION' } });
    }
  }

  return json(event, 200, { ok:true });
});