// netlify/functions/jobs_client0815.js
const {
  withCors, json, db, admin,
  ymdIST, ymdToDmy,
  getCalendarWindow,
  auditLog, renderTemplate,
  sendEmail,
  resolveStartRecipients,
  buildGoogleCalendarTemplateUrl
} = require('./_common');
const { requireCron } = require('./_auth');

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const cron = requireCron(event);
  if (cron.error) return cron.error;

  const body = event.body ? JSON.parse(event.body) : {};
  const force = !!body.force;

  const todayYmd = ymdIST(new Date());
  const settingsRef = db().collection('settings').doc('notifications');
  const settingsSnap = await settingsRef.get();
  const settings = settingsSnap.exists ? settingsSnap.data() : {};
  const last = settings.lastClient0815RunYmd || null;

  if (!force && last === todayYmd) {
    return json(event, 200, { ok:true, skipped:true, reason:'Already ran today', todayYmd });
  }

  const activeStatuses = ['PENDING','IN_PROGRESS','CLIENT_PENDING','APPROVAL_PENDING'];

  const snap = await db().collection('tasks')
    .where('startDateYmd', '==', todayYmd)
    .where('clientStartMailSent', '==', false)
    .get();

  const window = await getCalendarWindow();

  let sentCount = 0;
  let skippedNoTemplate = 0;
  let skippedNoEmail = 0;
  let skippedNoTaskClient = 0;

  for (const doc of snap.docs) {
    const t = doc.data();
    if (!activeStatuses.includes(t.status)) continue;

    // Must have template
    if (!t.clientStartSubject && !t.clientStartBody) {
      skippedNoTemplate++;
      continue;
    }

    // Load client
    if (!t.clientId) {
      skippedNoTaskClient++;
      continue;
    }
    const cSnap = await db().collection('clients').doc(t.clientId).get();
    const client = cSnap.exists ? cSnap.data() : {};

    // Resolve recipients with sendClientStartMail + internal trail + To promotion
    const recipients = await resolveStartRecipients({ client, task: t });
    if (!recipients.to.length) {
      skippedNoEmail++;
      continue;
    }

    const addToCalendarUrl = buildGoogleCalendarTemplateUrl({
      title: `START: ${t.title || 'Task'}`,
      startYmd: t.startDateYmd,
      startHH: window.startHH,
      endHH: window.endHH,
      timeZone: window.timeZone,
      details:
        `Client: ${client.name || ''}\n` +
        `Task: ${t.title || ''}\n` +
        `Start: ${ymdToDmy(t.startDateYmd)}\n` +
        `Due: ${ymdToDmy(t.dueDateYmd)}\n`
    });

    const vars = {
      clientName: client.name || '',
      taskTitle: t.title || '',
      startDate: ymdToDmy(t.startDateYmd),
      dueDate: ymdToDmy(t.dueDateYmd),
      addToCalendarUrl
    };

    const subject = renderTemplate(
      t.clientStartSubject || `We started {{taskTitle}}`,
      vars
    );

    const baseBody = renderTemplate(
      t.clientStartBody || `Dear {{clientName}},\n\nWe started work on {{taskTitle}}.\nDue: {{dueDate}}\n\nRegards,\nCompliance Team`,
      vars
    );

    const appended =
      `${baseBody}\n\n---\n` +
      `Add to your Google Calendar:\n` +
      `${addToCalendarUrl}`;

    const mailRes = await sendEmail({
      to: recipients.to,
      cc: recipients.cc,
      bcc: recipients.bcc,
      subject,
      html: appended
    });

    await doc.ref.update({
      clientStartMailSent: true,
      clientStartMailSentAt: admin.firestore.FieldValue.serverTimestamp(),
      clientStartGmailThreadId: mailRes?.threadId || null,
      clientStartGmailId: mailRes?.gmailId || null,
      clientStartRfcMessageId: mailRes?.rfcMessageId || null,
      clientStartReferences: mailRes?.references || null,
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    await auditLog({
      taskId: doc.id,
      action: 'EMAIL_SENT',
      actorUid: null,
      actorEmail: null,
      details: {
        type: 'CLIENT_START',
        to: recipients.to,
        cc: recipients.cc,
        bcc: recipients.bcc,
        sendClientStartMail: (t.sendClientStartMail !== false),
        addToCalendarUrl
      }
    });

    sentCount++;
  }

  await settingsRef.set({ lastClient0815RunYmd: todayYmd }, { merge: true });

  return json(event, 200, {
    ok: true,
    todayYmd,
    sentCount,
    skippedNoTemplate,
    skippedNoEmail,
    skippedNoTaskClient
  });
});