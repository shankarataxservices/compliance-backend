const { withCors, json, db } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = event.body ? JSON.parse(event.body) : {};
  const onlyActive = body.onlyActive === true;

  let q = db().collection('users');
  if (onlyActive) q = q.where('active', '==', true);

  // orderBy on email is fine (single-field index)
  const snap = await q.orderBy('email', 'asc').limit(500).get();

  const users = snap.docs.map(d => ({ uid: d.id, ...d.data() }));
  return json(event, 200, { ok:true, users });
});