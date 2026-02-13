// netlify/functions/series_reassign.js
const { withCors, json, db, admin, auditLog } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

async function findUserUidByEmail(email) {
  if (!email) return null;
  const e = String(email).trim().toLowerCase();

  let snap = await db().collection('users').where('emailLower', '==', e).limit(1).get();
  if (snap.empty) snap = await db().collection('users').where('email', '==', email).limit(1).get();
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
  const { seriesId, assignedToEmail } = body;

  if (!seriesId || !assignedToEmail) {
    return json(event, 400, { ok:false, error:'seriesId and assignedToEmail required' });
  }

  const uid = await findUserUidByEmail(assignedToEmail);
  if (!uid) {
    return json(event, 400, { ok:false, error:'assignedToEmail not found in users collection' });
  }

  const snap = await db().collection('tasks').where('seriesId', '==', seriesId).get();
  if (snap.empty) return json(event, 404, { ok:false, error:'No tasks found for seriesId' });

  const batch = db().batch();
  snap.docs.forEach(d => batch.update(d.ref, {
    assignedToEmail,
    assignedToUid: uid,
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }));
  await batch.commit();

  await auditLog({
    taskId: null,
    action: 'SERIES_REASSIGN',
    actorUid: user.uid,
    actorEmail: user.email,
    details: { seriesId, assignedToEmail }
  });

  return json(event, 200, { ok:true, updatedCount: snap.size });
});