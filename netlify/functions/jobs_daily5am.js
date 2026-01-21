const { withCors, json, db, ymd, addDays, sendEmail, auditLog, admin } = require('./_common');
const { requireCron } = require('./_auth');

function uniq(arr) { return [...new Set((arr||[]).map(x=>String(x).trim()).filter(Boolean))]; }

function renderTemplate(str, vars) {
  return String(str || '')
    .replaceAll('{{clientName}}', vars.clientName || '')
    .replaceAll('{{taskTitle}}', vars.taskTitle || '')
    .replaceAll('{{startDate}}', vars.startDate || '')
    .replaceAll('{{dueDate}}', vars.dueDate || '');
}

function buildDigestHtml(tasks) {
  const lines = tasks
    .sort((a,b)=>String(a.dueDateYmd||'').localeCompare(String(b.dueDateYmd||'')))
    .map(t => `<li><b>${t.title}</b> — Client: ${t.clientId} — Start ${t.startDateYmd} — Due ${t.dueDateYmd} — <b>${t.status}</b></li>`)
    .join('');
  return `<p>Tasks requiring action:</p><ul>${lines || '<li>None</li>'}</ul>`;
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const cron = requireCron(event);
  if (cron.error) return cron.error;

  const body = event.body ? JSON.parse(event.body) : {};
  const force = !!body.force;

  const todayYmd = ymd(new Date());

  // settings are only for INTERNAL daily digest
  const settingsRef = db().collection('settings').doc('notifications');
  const settingsSnap = await settingsRef.get();
  const settings = settingsSnap.exists ? settingsSnap.data() : {
    dailyInternalEmails: [],
    dailyWindowDays: 30,
    sendDailyToAssignees: true,
    lastDailyRunYmd: null
  };

  if (!force && settings.lastDailyRunYmd === todayYmd) {
    return json(event, 200, { ok:true, skipped:true, reason:'Already ran today' });
  }

  const activeStatuses = ['PENDING','IN_PROGRESS','CLIENT_PENDING','APPROVAL_PENDING'];

  // ========= PART 1: CLIENT START MAILS (send only on start date, once) =========
  // Query minimal fields to reduce index requirements:
  const startSnap = await db().collection('tasks')
    .where('startDateYmd', '==', todayYmd)
    .where('clientStartMailSent', '==', false)
    .get();

  let clientMailsSent = 0;

  for (const doc of startSnap.docs) {
    const t = doc.data();
    if (!activeStatuses.includes(t.status)) continue;

    // If no template in CSV, don't mail.
    if (!t.clientStartSubject && !t.clientStartBody) continue;

    const cSnap = await db().collection('clients').doc(t.clientId).get();
    const client = cSnap.exists ? cSnap.data() : {};
    if (!client.primaryEmail) continue;

    const vars = {
      clientName: client.name || '',
      taskTitle: t.title || '',
      startDate: t.startDateYmd || '',
      dueDate: t.dueDateYmd || ''
    };

    const subject = renderTemplate(t.clientStartSubject || `We have started work on {{taskTitle}}`, vars);
    const html = renderTemplate(
      t.clientStartBody || `Dear {{clientName}},<br><br>We have started working on <b>{{taskTitle}}</b>.<br>Due date: <b>{{dueDate}}</b>.<br><br>Regards,<br>Compliance Team`,
      vars
    );

    await sendEmail({
      to: [client.primaryEmail],
      cc: client.ccEmails || [],
      bcc: client.bccEmails || [],
      subject,
      html
    });

    await doc.ref.update({
      clientStartMailSent: true,
      clientStartMailSentAt: admin.firestore.FieldValue.serverTimestamp()
    });

    await auditLog({ taskId: doc.id, action:'EMAIL_SENT', actorUid:null, actorEmail:null, details:{ type:'CLIENT_START' } });
    clientMailsSent++;
  }

  // ========= PART 2: INTERNAL DAILY DIGEST (optional) =========
  const windowDays = Number(settings.dailyWindowDays || 30);
  const endYmd = ymd(addDays(new Date(), windowDays));

  const dueSoonSnap = await db().collection('tasks')
    .where('status', 'in', activeStatuses)
    .where('dueDateYmd', '>=', todayYmd)
    .where('dueDateYmd', '<=', endYmd)
    .get();

  const overdueSnap = await db().collection('tasks')
    .where('status', 'in', activeStatuses)
    .where('dueDateYmd', '<', todayYmd)
    .get();

  const digestTasks = [
    ...dueSoonSnap.docs.map(d => ({ id:d.id, ...d.data() })),
    ...overdueSnap.docs.map(d => ({ id:d.id, ...d.data() })),
  ];

  // Group by assignee
  const byAssignee = new Map();
  for (const t of digestTasks) {
    const key = t.assignedToEmail || '';
    if (!key) continue;
    if (!byAssignee.has(key)) byAssignee.set(key, []);
    byAssignee.get(key).push(t);
  }

  const internalExtra = uniq(settings.dailyInternalEmails || []);

  if (settings.sendDailyToAssignees !== false) {
    for (const [email, list] of byAssignee.entries()) {
      await sendEmail({
        to: [email],
        subject: `Daily Task Digest (${todayYmd})`,
        html: buildDigestHtml(list)
      });
    }
  }

  if (internalExtra.length) {
    await sendEmail({
      to: internalExtra,
      subject: `Firm Daily Digest (${todayYmd})`,
      html: buildDigestHtml(digestTasks)
    });
  }

  await settingsRef.set({ lastDailyRunYmd: todayYmd }, { merge: true });

  return json(event, 200, {
    ok:true,
    todayYmd,
    clientMailsSent,
    digestCount: digestTasks.length
  });
});