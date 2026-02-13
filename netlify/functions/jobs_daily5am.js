// netlify/functions/jobs_daily5am.js
const {
  withCors, json, db,
  ymdIST, ymdToDmy,
  addDays, dateFromYmdIST,
  sendEmail
} = require('./_common');
const { requireCron } = require('./_auth');

function uniq(arr) {
  return [...new Set((arr||[]).map(x=>String(x).trim()).filter(Boolean))];
}

function groupTasksForDigest(tasks, todayYmd) {
  const sections = {
    overdue: [],
    dueToday: [],
    dueIn3: [],
    dueIn7: [],
    dueIn15: [],
    dueIn30: [],
    approvalPending: [],
  };

  const toMidnight = (ymd) => new Date(`${ymd}T00:00:00+05:30`).getTime(); // IST anchored
  const t0 = toMidnight(todayYmd);

  for (const t of tasks) {
    if (!t.dueDateYmd) continue;

    if (t.status === 'APPROVAL_PENDING') sections.approvalPending.push(t);

    const d0 = toMidnight(t.dueDateYmd);
    const diffDays = Math.floor((d0 - t0) / (24 * 3600 * 1000));

    if (t.status === 'COMPLETED') continue;

    if (diffDays < 0) sections.overdue.push({ ...t, diffDays });
    else if (diffDays === 0) sections.dueToday.push({ ...t, diffDays });
    else if (diffDays <= 3) sections.dueIn3.push({ ...t, diffDays });
    else if (diffDays <= 7) sections.dueIn7.push({ ...t, diffDays });
    else if (diffDays <= 15) sections.dueIn15.push({ ...t, diffDays });
    else if (diffDays <= 30) sections.dueIn30.push({ ...t, diffDays });
  }

  const sortByDue = (a,b)=>String(a.dueDateYmd||'').localeCompare(String(b.dueDateYmd||''));
  Object.keys(sections).forEach(k => sections[k].sort(sortByDue));
  return sections;
}

async function loadClientsMap(clientIds) {
  const ids = uniq(clientIds);
  const map = new Map();
  if (!ids.length) return map;
  const refs = ids.map(id => db().collection('clients').doc(id));
  const snaps = await db().getAll(...refs);
  snaps.forEach(s => { if (s.exists) map.set(s.id, s.data()); });
  return map;
}

function escapeHtml(s) {
  return String(s || '')
    .replaceAll('&','&amp;')
    .replaceAll('<','&lt;')
    .replaceAll('>','&gt;')
    .replaceAll('"','&quot;')
    .replaceAll("'",'&#039;');
}

function renderSection(title, items, clientsMap) {
  const li = (t) => {
    const c = clientsMap.get(t.clientId) || {};
    const clientName = c.name || '';
    const due = ymdToDmy(t.dueDateYmd);
    const start = ymdToDmy(t.startDateYmd);
    const assignee = t.assignedToEmail || '';
    const status = t.status || '';
    const note = (t.statusNote || '').trim();

    // IMPORTANT CHANGE: do NOT show TaskId in digest email
    return `<li>
      <b>${escapeHtml(t.title || '')}</b>
      <div style="color:#555;font-size:12px;margin-top:2px">
        Client: ${escapeHtml(clientName)} |
        Start: ${escapeHtml(start)} |
        Due: <b>${escapeHtml(due)}</b> |
        Status: <b>${escapeHtml(status)}</b> |
        Assignee: ${escapeHtml(assignee)}
      </div>
      ${note ? `<div style="color:#666;font-size:12px;margin-top:2px">Note: ${escapeHtml(note)}</div>` : ``}
    </li>`;
  };

  if (!items.length) return '';
  return `
    <h3 style="margin:14px 0 6px">${escapeHtml(title)} (${items.length})</h3>
    <ul style="margin:0 0 10px 18px;padding:0">${items.map(li).join('')}</ul>
  `;
}

function isDigestEmpty(sections) {
  return Object.values(sections).every(arr => !arr || arr.length === 0);
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
  const settings = settingsSnap.exists ? settingsSnap.data() : {
    dailyInternalEmails: [],
    dailyWindowDays: 30,
    sendDailyToAssignees: true,
    lastDailyRunYmd: null
  };

  if (!force && settings.lastDailyRunYmd === todayYmd) {
    return json(event, 200, { ok:true, skipped:true, reason:'Already ran today', todayYmd });
  }

  const activeStatuses = ['PENDING','IN_PROGRESS','CLIENT_PENDING','APPROVAL_PENDING'];
  const windowDays = Number(settings.dailyWindowDays || 30);
  const endYmd = ymdIST(addDays(dateFromYmdIST(todayYmd), windowDays));

  const dueSoonSnap = await db().collection('tasks')
    .where('status', 'in', activeStatuses)
    .where('dueDateYmd', '>=', todayYmd)
    .where('dueDateYmd', '<=', endYmd)
    .get();

  const overdueSnap = await db().collection('tasks')
    .where('status', 'in', activeStatuses)
    .where('dueDateYmd', '<', todayYmd)
    .get();

  const tasks = [
    ...dueSoonSnap.docs.map(d => ({ id:d.id, ...d.data() })),
    ...overdueSnap.docs.map(d => ({ id:d.id, ...d.data() })),
  ];

  const clientsMap = await loadClientsMap(tasks.map(t => t.clientId));
  const sectionsFirm = groupTasksForDigest(tasks, todayYmd);

  const byAssignee = new Map();
  for (const t of tasks) {
    const email = String(t.assignedToEmail || '').trim();
    if (!email) continue;
    if (!byAssignee.has(email)) byAssignee.set(email, []);
    byAssignee.get(email).push(t);
  }

  const internalExtra = uniq(settings.dailyInternalEmails || []);
  const subject = `Daily Digest (${ymdToDmy(todayYmd)})`;

  const makeHtml = (list) => {
    const s = groupTasksForDigest(list, todayYmd);

    // IMPORTANT CHANGE: if empty -> "No Tasks To Display"
    if (isDigestEmpty(s)) {
      return `
        <div style="font-family:Arial,sans-serif;line-height:1.35">
          <h2 style="margin:0 0 6px">Daily Digest — ${escapeHtml(ymdToDmy(todayYmd))}</h2>
          <div style="margin-top:10px;color:#555;font-size:13px;font-weight:700">
            No Tasks To Display
          </div>
        </div>
      `;
    }

    return `
      <div style="font-family:Arial,sans-serif;line-height:1.35">
        <h2 style="margin:0 0 6px">Firm Task Digest — ${escapeHtml(ymdToDmy(todayYmd))}</h2>
        <div style="color:#555;font-size:12px;margin-bottom:10px">
          Window: next ${windowDays} days + overdue. Statuses: ${activeStatuses.join(', ')}
        </div>
        ${renderSection('Overdue', s.overdue, clientsMap)}
        ${renderSection('Due Today', s.dueToday, clientsMap)}
        ${renderSection('Due in 1–3 days', s.dueIn3, clientsMap)}
        ${renderSection('Due in 4–7 days', s.dueIn7, clientsMap)}
        ${renderSection('Due in 8–15 days', s.dueIn15, clientsMap)}
        ${renderSection('Due in 16–30 days', s.dueIn30, clientsMap)}
        ${renderSection('Waiting for approval', s.approvalPending, clientsMap)}
        <hr style="border:none;border-top:1px solid #ddd;margin:14px 0">
        <div style="color:#777;font-size:12px">
          Tip: Use the web app to filter by client, status and due date.
        </div>
      </div>
    `;
  };

  let sentToAssignees = 0;
  if (settings.sendDailyToAssignees !== false) {
    for (const [email, list] of byAssignee.entries()) {
      await sendEmail({ to: [email], subject, html: makeHtml(list) });
      sentToAssignees++;
    }
  }

  let sentToInternal = 0;
  if (internalExtra.length) {
    await sendEmail({
      to: internalExtra,
      subject: `Firm ${subject}`,
      html: makeHtml(tasks)
    });
    sentToInternal = internalExtra.length;
  }

  await settingsRef.set({ lastDailyRunYmd: todayYmd }, { merge: true });

  return json(event, 200, {
    ok:true,
    todayYmd,
    tasksCount: tasks.length,
    sentToAssignees,
    sentToInternal
  });
});