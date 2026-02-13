// netlify/functions/users_setrole.js
const { withCors, json, db, admin } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

const ROLES = new Set(['PARTNER','MANAGER','ASSOCIATE','WORKER']); // keep WORKER for compat/migration

async function findUserUidByEmail(email) {
  if (!email) return null;
  const e = String(email).trim().toLowerCase();
  const snap = await db().collection('users').where('emailLower', '==', e).limit(1).get();
  if (snap.empty) return null;
  return snap.docs[0].id;
}

function normalizeRole(role) {
  const r = String(role || '').toUpperCase().trim();
  if (r === 'WORKER') return 'ASSOCIATE'; // enforce rename but accept input
  return r;
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

  const roleIn = body.role || '';
  const role = normalizeRole(roleIn);
  const active = body.active !== false;

  if (!uid && !email) return json(event, 400, { ok:false, error:'uid or email required' });
  if (!ROLES.has(String(roleIn || '').toUpperCase().trim()) && String(roleIn || '').toUpperCase().trim() !== 'WORKER') {
    return json(event, 400, { ok:false, error:'role must be PARTNER / MANAGER / ASSOCIATE' });
  }
  if (!['PARTNER','MANAGER','ASSOCIATE'].includes(role)) {
    return json(event, 400, { ok:false, error:'role must be PARTNER / MANAGER / ASSOCIATE' });
  }

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