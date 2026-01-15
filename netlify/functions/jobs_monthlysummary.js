const { json, db, sendEmail } = require('./_common');
const { requireCron } = require('./_auth');

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: require('./_common').cors() };
  if (event.httpMethod !== 'POST') return json(405, { error: 'POST only' });

  const cron = requireCron(event);
  if (cron.error) return cron.error;

  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth(); // current
  const lastMonthStart = new Date(year, month - 1, 1);
  const lastMonthEnd = new Date(year, month, 0);

  const startYmd = `${lastMonthStart.getFullYear()}-${String(lastMonthStart.getMonth()+1).padStart(2,'0')}-01`;
  const endYmd = `${lastMonthEnd.getFullYear()}-${String(lastMonthEnd.getMonth()+1).padStart(2,'0')}-${String(lastMonthEnd.getDate()).padStart(2,'0')}`;

  const clientsSnap = await db().collection('clients').get();

  for (const cDoc of clientsSnap.docs) {
    const client = cDoc.data();
    if (!client.primaryEmail) continue;

    const tasksSnap = await db().collection('tasks')
      .where('clientId', '==', cDoc.id)
      .where('dueDateYmd', '>=', startYmd)
      .where('dueDateYmd', '<=', endYmd)
      .get();

    const completed = [];
    const pending = [];

    for (const tDoc of tasksSnap.docs) {
      const t = tDoc.data();
      if (t.status === 'COMPLETED') completed.push(t);
      else pending.push(t);
    }

    const html = `
      <p>Monthly Compliance Summary for <b>${client.name}</b></p>
      <p>Period: ${startYmd} to ${endYmd}</p>
      <h3>Completed</h3>
      <ul>${completed.map(x => `<li>${x.title} (Due ${x.dueDateYmd})</li>`).join('') || '<li>None</li>'}</ul>
      <h3>Pending / In progress</h3>
      <ul>${pending.map(x => `<li>${x.title} (Due ${x.dueDateYmd}) - ${x.status}</li>`).join('') || '<li>None</li>'}</ul>
    `;

    await sendEmail({
      to: [client.primaryEmail],
      cc: client.ccEmails || [],
      bcc: client.bccEmails || [],
      subject: `Monthly Compliance Summary - ${client.name}`,
      html
    });
  }

  return json(200, { ok: true });
};
