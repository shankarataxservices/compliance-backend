const {
  withCors, json, db, admin,
  ymdIST, ymdToDmy,
  getCalendarWindow,
  auditLog, renderTemplate,
  sendEmail,
  uniqEmails
} = require('./_common');
const { requireCron } = require('./_auth');

async function getManagerEmailForAssignee(assignedToUid) {
  if (!assignedToUid) return null;
  try {
    const snap = await db().collection('users').doc(assignedToUid).get();
    if (!snap.exists) return null;
    const u = snap.data() || {};
    return u.managerEmail || null;
  } catch {
    return null;
  }
}

function buildGoogleCalendarTemplateUrl({ title, startYmd, startHH, endHH, timeZone, details }) {
  const ymdToCompact = (s) => String(s).replaceAll('-', '');
  const hh2 = (h) => String(h).padStart(2, '0');

  // single day timed event
  const start = `${ymdToCompact(startYmd)}T${hh2(startHH)}0000`;
  const end = `${ymdToCompact(startYmd)}T${hh2(endHH)}0000`;

  const params = new URLSearchParams({
    action: 'TEMPLATE',
    text: title || 'Compliance Task',
    dates: `${start}/${end}`,
    ctz: timeZone || 'Asia/Kolkata',
    details: details || ''
  });

  return `https://calendar.google.com/calendar/render?${params.toString()}`;
}

function mergeRecipients({ client, task, managerEmail }) {
  // To: task overrides, else client.primaryEmail
  const to = (task.clientToEmails && task.clientToEmails.length)
    ? task.clientToEmails
    : (client.primaryEmail ? [client.primaryEmail] : []);

  // Base CC/BCC
  const cc = [
    ...(client.ccEmails || []),
    ...(task.clientCcEmails || []),
  ];
  const bcc = [
    ...(client.bccEmails || []),
    ...(task.clientBccEmails || []),
  ];

  // Optional CC controls
  if (task.ccAssigneeOnClientStart === true && task.assignedToEmail) cc.push(task.assignedToEmail);
  if (task.ccManagerOnClientStart === true && managerEmail) cc.push(managerEmail);

  return {
    to: uniqEmails(to),
    cc: uniqEmails(cc),
    bcc: uniqEmails(bcc),
  };
}

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

  for (const doc of snap.docs) {
    const t = doc.data();
    if (!activeStatuses.includes(t.status)) continue;

    // Must have template
    if (!t.clientStartSubject && !t.clientStartBody) {
      skippedNoTemplate++;
      continue;
    }

    const cSnap = await db().collection('clients').doc(t.clientId).get();
    const client = cSnap.exists ? cSnap.data() : {};

    // auto manager mapping from assignee user doc
    const managerEmail = await getManagerEmailForAssignee(t.assignedToUid);

    const recipients = mergeRecipients({ client, task: t, managerEmail });

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

    // Add calendar link even if template doesn't include it
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
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    await auditLog({
      taskId: doc.id,
      action: 'EMAIL_SENT',
      actorUid: null,
      actorEmail: null,
      details: {
        type:'CLIENT_START',
        to: recipients.to,
        cc: recipients.cc,
        bcc: recipients.bcc,
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
    skippedNoEmail
  });
});