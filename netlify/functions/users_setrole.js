const { withCors, json, db, admin } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

const ROLES = new Set(['PARTNER','MANAGER','WORKER']);

async function findUserUidByEmail(email) {
  if (!email) return null;
  const e = String(email).trim().toLowerCase();
  const snap = await db().collection('users').where('emailLower', '==', e).limit(1).get();
  if (snap.empty) return null;
  return snap.docs[0].id;
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');
  const uid = body.uid || null;
  const email = body.email || null;
  const role = String(body.role || '').toUpperCase().trim();
  const active = body.active !== false;

  if (!uid && !email) return json(event, 400, { ok:false, error:'uid or email required' });
  if (!ROLES.has(role)) return json(event, 400, { ok:false, error:'role must be PARTNER / MANAGER / WORKER' });

  const targetUid = uid || await findUserUidByEmail(email);
  if (!targetUid) return json(event, 404, { ok:false, error:'User not found' });

  await db().collection('users').doc(targetUid).set({
    role,
    active,
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedBy: user.email
  }, { merge: true });

  return json(event, 200, { ok:true, uid: targetUid, role, active });
});