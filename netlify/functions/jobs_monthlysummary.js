const { withCors, json, db, sendEmail } = require('./_common');
const { requireCron } = require('./_auth');

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const cron = requireCron(event);
  if (cron.error) return cron.error;

  // Keeping this as-is, but it will run ONLY if you schedule it.
  // If you don't want monthly mails, simply don't create a cron workflow for this.
  return json(event, 200, { ok:true, note:'Monthly summary not enabled in this setup (no-op).' });
});