const { withCors, json, db, admin, calendar, sendEmail, auditLog } = require('./_common');
const { requireUser } = require('./_auth');

exports.handler = withCors(async (event) => {
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

  // On completion: email client (only)
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

    // IMPORTANT: no daily mails to clients. Only this completion mail and D-15 reminder.
    if (to.length) {
      await sendEmail({
        to,
        cc: [], bcc: [],
        subject: `Completed: ${task.title} (${client.name || ''})`,
        html: `
          <p>Compliance completed.</p>
          <p><b>${task.title}</b></p>
          <p>Client: ${client.name || '-'}</p>
          <p>Due: ${task.dueDateYmd}</p>
          <p>Status Note: ${updates.statusNote || task.statusNote || ''}</p>
        `
      });
      await auditLog({ taskId, action:'EMAIL_SENT', actorUid:null, actorEmail:null, details:{ type:'CLIENT_COMPLETION' } });
    }
  }

  return json(event, 200, { ok:true });
});