const { json, db, sendEmail, auditLog, ymd, addDays } = require('./_common');
const { requireCron } = require('./_auth');

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: require('./_common').cors() };
  if (event.httpMethod !== 'POST') return json(405, { error: 'POST only' });

  const cron = requireCron(event);
  if (cron.error) return cron.error;

  const today = new Date();
  const todayYmd = ymd(today);

  const tasksSnap = await db().collection('tasks')
    .where('status', 'in', ['PENDING', 'IN_PROGRESS', 'CLIENT_PENDING', 'APPROVAL_PENDING'])
    .get();

  for (const doc of tasksSnap.docs) {
    const t = doc.data();

    const clientSnap = await db().collection('clients').doc(t.clientId).get();
    const client = clientSnap.exists ? clientSnap.data() : {};
    const to = client.primaryEmail ? [client.primaryEmail] : [];
    const cc = client.ccEmails || [];
    const bcc = client.bccEmails || [];

    if (!to.length) continue;

    // Start notification (startDate)
    if (t.startDateYmd === todayYmd && !t.escalation?.startSent) {
      await sendEmail({
        to, cc, bcc,
        subject: `Task started: ${t.title} (${client.name || ''})`,
        html: `<p>Task started: <b>${t.title}</b></p><p>Due: ${t.dueDateYmd}</p>`
      });
      await doc.ref.update({ 'escalation.startSent': true });
      await auditLog({ taskId: doc.id, action: 'EMAIL_SENT', actorUid: null, actorEmail: null, details: { type: 'START' } });
    }

    // Escalations by due date
    const d7 = ymd(addDays(new Date(t.dueDateYmd), -7));
    const d3 = ymd(addDays(new Date(t.dueDateYmd), -3));

    if (todayYmd === d7 && !t.escalation?.d7) {
      await sendEmail({ to, cc, bcc, subject: `Escalation: 7 days left - ${t.title}`, html: `<p>Due in 7 days: <b>${t.title}</b> (${t.dueDateYmd})</p>` });
      await doc.ref.update({ 'escalation.d7': true });
    }
    if (todayYmd === d3 && !t.escalation?.d3) {
      await sendEmail({ to, cc, bcc, subject: `Escalation: 3 days left - ${t.title}`, html: `<p>Due in 3 days: <b>${t.title}</b> (${t.dueDateYmd})</p>` });
      await doc.ref.update({ 'escalation.d3': true });
    }
    if (todayYmd === t.dueDateYmd && !t.escalation?.d0) {
      await sendEmail({ to, cc, bcc, subject: `Due today: ${t.title}`, html: `<p>Due today: <b>${t.title}</b></p>` });
      await doc.ref.update({ 'escalation.d0': true });
    }
    if (todayYmd > t.dueDateYmd && !t.escalation?.overdue) {
      await sendEmail({ to, cc, bcc, subject: `Overdue: ${t.title}`, html: `<p>Overdue: <b>${t.title}</b> (was due ${t.dueDateYmd})</p>` });
      await doc.ref.update({ 'escalation.overdue': true });
    }
  }

  return json(200, { ok: true, date: todayYmd });
};
