// netlify/functions/users_list.js
const { withCors, json, db } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

function normalizeRole(role) {
  const r = String(role || '').toUpperCase().trim();
  if (r === 'WORKER') return 'ASSOCIATE';
  return r || 'ASSOCIATE';
}

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

  const snap = await q.orderBy('email', 'asc').limit(500).get();

  const users = snap.docs.map(d => {
    const u = d.data() || {};
    return {
      uid: d.id,
      ...u,
      role: normalizeRole(u.role)
    };
  });

  return json(event, 200, { ok:true, users });
});