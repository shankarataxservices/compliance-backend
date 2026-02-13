// netlify/functions/tasks_updatestatus.js
const {
  withCors, json, db, admin,
  calendar, getCalendarWindow, calTimeRange,
  sendEmailReply, sendEmail,
  auditLog, renderTemplate, ymdToDmy, uniqEmails,
  resolveCompletionRecipients,
  resolveStartRecipients
} = require('./_common');
const { requireUser } = require('./_auth');

async function patchEvent({ eventId, whenYmd, summary, description, colorId = null, window }) {
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

function roleOf(user) {
  const r = String(user?.role || 'ASSOCIATE').toUpperCase().trim();
  // Back-compat already handled, but keep safe
  return (r === 'WORKER') ? 'ASSOCIATE' : (r || 'ASSOCIATE');
}
function isPrivileged(role) {
  return role === 'PARTNER' || role === 'MANAGER';
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const role = roleOf(user);
  const body = JSON.parse(event.body || {});
  const { taskId, newStatus, statusNote, delayReason, delayNotes } = body;

  if (!taskId || !newStatus) return json(event, 400, { ok:false, error:'taskId,newStatus required' });

  const tRef = db().collection('tasks').doc(taskId);
  const tSnap = await tRef.get();
  if (!tSnap.exists) return json(event, 404, { ok:false, error:'Task not found' });
  const task = tSnap.data();

  const privileged = isPrivileged(role);
  const isAssignee = task.assignedToUid === user.uid;

  if (!privileged && !isAssignee) return json(event, 403, { ok:false, error:'Not allowed' });

  // Completion permission:
  // Associates are allowed to mark COMPLETED as long as they are the assignee (already enforced above).
  const updates = {
    status: newStatus,
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
  };
  if (typeof statusNote === 'string') updates.statusNote = statusNote;
  if (delayReason !== undefined) updates.delayReason = delayReason || null;
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
  if (String(newStatus).toUpperCase() === 'COMPLETED') {
    const window = await getCalendarWindow();
    const cn = String(task.clientNameSnapshot || '').trim() || String(task.clientId || '').trim();
    const extra = String(task.calendarDescription || '').trim();
    const descBase =
     `Client: ${cn}\n` +
     `Start: ${task.startDateYmd}\n` +
     `Due: ${task.dueDateYmd}\n`;
    const desc = extra ? `${descBase}\n${extra}` : descBase;

    // Patch calendar event(s)
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

    // Respect per-task completion flag
    if (task.sendClientCompletionMail === false) {
      return json(event, 200, { ok:true, note:'Completed. Client completion email disabled for this task.' });
    }

    // Load client
    const cSnap = await db().collection('clients').doc(task.clientId).get();
    const client = cSnap.exists ? cSnap.data() : {};

    // Recipients: reply-all behaviour with overrides + internal trail
    const rec = await resolveCompletionRecipients({ client, task });
    if (!rec.to.length) {
      return json(event, 200, { ok:true, note:'Completed. No email recipient found for completion email.' });
    }

    const completedAtStr = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
    const vars = {
      clientName: client.name || '',
      taskTitle: task.title || '',
      startDate: ymdToDmy(task.startDateYmd),
      dueDate: ymdToDmy(task.dueDateYmd),
      completedAt: completedAtStr,
      statusNote: (typeof updates.statusNote === 'string') ? updates.statusNote : (task.statusNote || ''),
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

    // Thread reply: reply in same thread if we have it
    const threadId = task.clientStartGmailThreadId || null;
    const inReplyTo = task.clientStartRfcMessageId || null;
    const references = task.clientStartReferences || inReplyTo || null;

    if (threadId) {
      await sendEmailReply({
        threadId,
        inReplyTo,
        references,
        to: rec.to,
        cc: rec.cc,
        bcc: rec.bcc,
        subject,
        html
      });
    } else {
      // fallback: send as new mail
      await sendEmail({
        to: rec.to,
        cc: rec.cc,
        bcc: rec.bcc,
        subject,
        html
      });
    }

    await auditLog({
      taskId,
      action: 'EMAIL_SENT',
      actorUid: null,
      actorEmail: null,
      details: {
        type:'CLIENT_COMPLETION',
        repliedToStartThread: !!threadId,
        to: rec.to,
        cc: rec.cc,
        bcc: rec.bcc
      }
    });
  }

  return json(event, 200, { ok:true });
});