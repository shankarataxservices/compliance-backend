const {
  withCors, json, db, admin,
  calendar, getCalendarWindow, calTimeRange,
  sendEmailReply, sendEmail,
  auditLog, renderTemplate, ymdToDmy, uniqEmails
} = require('./_common');
const { requireUser } = require('./_auth');

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

function mergeRecipients({ client, task }) {
  // To: task overrides, else client.primaryEmail
  const to = (task.clientToEmails && task.clientToEmails.length)
    ? task.clientToEmails
    : (client.primaryEmail ? [client.primaryEmail] : []);

  // CC/BCC: client defaults + task overrides
  const cc = [
    ...(client.ccEmails || []),
    ...(task.clientCcEmails || []),
  ];
  const bcc = [
    ...(client.bccEmails || []),
    ...(task.clientBccEmails || []),
  ];

  return {
    to: uniqEmails(to),
    cc: uniqEmails(cc),
    bcc: uniqEmails(bcc),
  };
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

  // ===== On COMPLETED =====
  if (newStatus === 'COMPLETED') {
    const window = await getCalendarWindow();

    const desc =
      `ClientId: ${task.clientId}\n` +
      `Start: ${task.startDateYmd}\n` +
      `Due: ${task.dueDateYmd}\n` +
      `TaskId: ${taskId}`;

    // Patch calendar event (single event). Also patch old fields if task was created earlier.
    const eventId = task.calendarEventId || task.calendarStartEventId || null;
    try {
      if (eventId) {
        await patchEvent({
          eventId,
          whenYmd: task.startDateYmd,
          summary: `[COMPLETED] START: ${task.title}`,
          description: desc,
          colorId: '2',
          window
        });
      }
      // Backward compatibility: if old due event exists, patch it too.
      if (task.calendarDueEventId) {
        await patchEvent({
          eventId: task.calendarDueEventId,
          whenYmd: task.dueDateYmd,
          summary: `[COMPLETED] DUE: ${task.title}`,
          description: desc,
          colorId: '2',
          window
        });
      }
    } catch (e) {
      console.warn('Calendar patch failed (ignored):', e.message || e);
    }

    // Respect per-task flag
    if (task.sendClientCompletionMail === false) {
      return json(event, 200, { ok:true, note:'Completed. Client completion mail disabled for this task.' });
    }

    // Load client
    const cSnap = await db().collection('clients').doc(task.clientId).get();
    const client = cSnap.exists ? cSnap.data() : {};

    const { to, cc, bcc } = mergeRecipients({ client, task });
    if (!to.length) {
      return json(event, 200, { ok:true, note:'Completed. No client email found to send completion mail.' });
    }

    const completedAtText = `${ymdToDmy(task.dueDateYmd)} (completed at: ${new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' })})`;

    const vars = {
      clientName: client.name || '',
      taskTitle: task.title || '',
      startDate: ymdToDmy(task.startDateYmd),
      dueDate: ymdToDmy(task.dueDateYmd),
      completedAt: new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
      statusNote: (typeof updates.statusNote === 'string' ? updates.statusNote : (task.statusNote || '')),
    };

    const subject = renderTemplate(
      task.clientCompletionSubject || `Completed: {{taskTitle}}`,
      vars
    );

    const html = renderTemplate(
      task.clientCompletionBody || (
        `Dear {{clientName}},\n\n` +
        `We have completed: {{taskTitle}}\n` +
        `Due date: {{dueDate}}\n` +
        `Completed at: {{completedAt}}\n` +
        `Status note: {{statusNote}}\n\n` +
        `Regards,\n${process.env.MAIL_SIGNATURE || 'Compliance Team'}`
      ),
      vars
    );

    // Reply in same thread if we have it
    const threadId = task.clientStartGmailThreadId || null;
    const inReplyTo = task.clientStartRfcMessageId || null;

    if (threadId) {
      await sendEmailReply({
        threadId,
        inReplyTo,
        references: inReplyTo,
        to, cc, bcc,
        subject,
        html
      });
    } else {
      // fallback: send as new mail
      await sendEmail({ to, cc, bcc, subject, html });
    }

    await auditLog({
      taskId,
      action: 'EMAIL_SENT',
      actorUid: null,
      actorEmail: null,
      details: { type:'CLIENT_COMPLETION', repliedToStartThread: !!threadId }
    });
  }

  return json(event, 200, { ok:true });
});