const { withCors, json, db } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const snap = await db().collection('settings').doc('calendar').get();
  const data = snap.exists ? snap.data() : null;

  return json(event, 200, {
    ok:true,
    data: data || { startHH: 10, endHH: 12, timeZone: 'Asia/Kolkata' }
  });
});