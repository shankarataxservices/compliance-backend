const { json, db, admin, calendar, sendEmail, auditLog, drive } = require('./_common');
const { requireUser } = require('./_auth');

async function downloadDriveFileBase64(fileId) {
  const d = drive();
  const res = await d.files.get({ fileId, alt: 'media' }, { responseType: 'arraybuffer' });
  return Buffer.from(res.data).toString('base64');
}

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: require('./_common').cors() };
  if (event.httpMethod !== 'POST') return json(405, { error: 'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const body = JSON.parse(event.body || '{}');
  const { taskId, newStatus, delayReason, delayNotes } = body;
  if (!taskId || !newStatus) return json(400, { error: 'taskId,newStatus required' });

  const tRef = db().collection('tasks').doc(taskId);
  const tSnap = await tRef.get();
  if (!tSnap.exists) return json(404, { error: 'Task not found' });

  const task = tSnap.data();

  // Authorization
  const isPartner = user.role === 'PARTNER';
  const isAssignee = task.assignedToUid === user.uid;

  if (!isPartner && !isAssignee) return json(403, { error: 'Not allowed' });

  // Workflow rules
  if (!isPartner) {
    // Worker cannot mark COMPLETED
    if (newStatus === 'COMPLETED') return json(403, { error: 'Only partner can complete' });
  }

  const updates = {
    status: newStatus,
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  };

  if (delayReason) updates.delayReason = delayReason;
  if (delayNotes !== undefined) updates.delayNotes = delayNotes;

  if (newStatus === 'APPROVAL_PENDING') {
    updates.completedRequestedAt = admin.firestore.FieldValue.serverTimestamp();
  }

  if (newStatus === 'COMPLETED') {
    updates.completedAt = admin.firestore.FieldValue.serverTimestamp();
  }

  await tRef.update(updates);

  await auditLog({
    taskId,
    action: 'STATUS_CHANGE',
    actorUid: user.uid,
    actorEmail: user.email,
    details: { from: task.status, to: newStatus, delayReason: delayReason || null }
  });

  // If partner completed: patch calendar + email
  if (newStatus === 'COMPLETED') {
    if (task.calendarEventId) {
      try {
        await calendar().events.patch({
          calendarId: 'primary',
          eventId: task.calendarEventId,
          requestBody: { summary: `[COMPLETED] ${task.title}`, colorId: '2' }
        });
      } catch (e) {
        // still continue email
      }
    }

    const clientSnap = await db().collection('clients').doc(task.clientId).get();
    const client = clientSnap.exists ? clientSnap.data() : {};

    const to = client.primaryEmail ? [client.primaryEmail] : [];
    const cc = client.ccEmails || [];
    const bcc = client.bccEmails || [];

    // Attachments (download from Drive, attach if reasonably small)
    const attachments = [];
    for (const a of (task.attachments || [])) {
      try {
        const contentBase64 = await downloadDriveFileBase64(a.driveFileId);
        attachments.push({
          filename: a.fileName,
          mimeType: a.mimeType || 'application/pdf',
          contentBase64
        });
      } catch (e) {
        // fallback: ignore attachment if download fails
      }
    }

    const subject = `Completed: ${task.title} (${client.name || ''})`;
    const html = `
      <p>Task <b>${task.title}</b> completed.</p>
      <p>Client: ${client.name || '-'}</p>
      <p>Due: ${task.dueDateYmd}</p>
      <p>Status: COMPLETED</p>
    `;

    if (to.length) {
      await sendEmail({ to, cc, bcc, subject, html, attachments });
      await auditLog({ taskId, action: 'EMAIL_SENT', actorUid: null, actorEmail: null, details: { type: 'COMPLETION' } });
    }
  }

  return json(200, { ok: true });
};
