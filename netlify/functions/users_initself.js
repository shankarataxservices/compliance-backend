// netlify/functions/users_initself.js
const { withCors, json, db, admin } = require('./_common');
const { requireUser } = require('./_auth');

function normName(s) {
  return String(s || '').trim().replace(/\s+/g, ' ').slice(0, 60);
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const body = event.body ? JSON.parse(event.body) : {};
  const displayName = normName(body.displayName || (user.email ? user.email.split('@')[0] : ''));

  const ref = db().collection('users').doc(user.uid);
  const snap = await ref.get();
  const now = admin.firestore.FieldValue.serverTimestamp();

  const base = {
    email: user.email || '',
    emailLower: (user.email || '').toLowerCase(),
    displayName: displayName || '',
    displayNameLower: (displayName || '').toLowerCase(),
    role: snap.exists ? (snap.data()?.role || 'ASSOCIATE') : 'ASSOCIATE',
    active: snap.exists ? (snap.data()?.active !== false) : true,
    managerEmail: snap.exists ? (snap.data()?.managerEmail || null) : null,
    managerUid: snap.exists ? (snap.data()?.managerUid || null) : null,
    updatedAt: now,
    updatedBy: user.email || null,
  };

  if (!snap.exists) base.createdAt = now;

  await ref.set(base, { merge: true });
  return json(event, 200, { ok:true, uid: user.uid });
});