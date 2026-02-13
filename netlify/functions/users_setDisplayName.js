// netlify/functions/users_setdisplayname.js
const { withCors, json, db, admin } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

function normName(s) {
  return String(s || '').trim().replace(/\s+/g, ' ');
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');
  const uid = String(body.uid || '').trim();
  const displayName = normName(body.displayName);

  if (!uid) return json(event, 400, { ok:false, error:'uid required' });
  if (!displayName) return json(event, 400, { ok:false, error:'displayName required' });
  if (displayName.length > 60) return json(event, 400, { ok:false, error:'displayName too long (max 60)' });

  await db().collection('users').doc(uid).set({
    displayName,
    displayNameLower: displayName.toLowerCase(),
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedBy: user.email
  }, { merge: true });

  return json(event, 200, { ok:true, uid, displayName });
});