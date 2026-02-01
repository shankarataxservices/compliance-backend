const { withCors, json, db, admin } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

async function findUserByEmailLower(email) {
  const e = String(email || '').trim().toLowerCase();
  if (!e) return null;
  const snap = await db().collection('users').where('emailLower', '==', e).limit(1).get();
  if (snap.empty) return null;
  return { uid: snap.docs[0].id, ...snap.docs[0].data() };
}

async function findUserByUid(uid) {
  const s = await db().collection('users').doc(uid).get();
  if (!s.exists) return null;
  return { uid: s.id, ...s.data() };
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');
  const targetUid = body.uid || null;
  const targetEmail = body.email || null;

  const managerEmail = (body.managerEmail || '').trim(); // can be empty to clear

  if (!targetUid && !targetEmail) return json(event, 400, { ok:false, error:'uid or email required' });

  const target = targetUid ? await findUserByUid(targetUid) : await findUserByEmailLower(targetEmail);
  if (!target) return json(event, 404, { ok:false, error:'Target user not found' });

  let managerUid = null;
  let managerRole = null;

  if (managerEmail) {
    const mgr = await findUserByEmailLower(managerEmail);
    if (!mgr) return json(event, 400, { ok:false, error:'managerEmail not found in users' });

    managerUid = mgr.uid;
    managerRole = mgr.role || 'WORKER';

    // Allow manager to be PARTNER or MANAGER (recommended)
    if (!['PARTNER','MANAGER'].includes(String(managerRole).toUpperCase())) {
      return json(event, 400, { ok:false, error:'managerEmail must belong to a PARTNER or MANAGER user' });
    }
  }

  await db().collection('users').doc(target.uid).set({
    managerEmail: managerEmail || null,
    managerUid: managerUid || null,
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedBy: user.email
  }, { merge: true });

  return json(event, 200, { ok:true, uid: target.uid, managerEmail: managerEmail || null, managerUid: managerUid || null });
});