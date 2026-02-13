// netlify/functions/users_setManager.js
const { withCors, json, db, admin } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

/**
 * Allows saving ANY managerEmail (even if manager user not created yet)
 * - If managerEmail exists in users (emailLower match), also store managerUid (only if role is PARTNER/MANAGER)
 * - If managerEmail is empty => clear managerEmail + managerUid
 */

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

  const body = JSON.parse(event.body || '{}');
  const targetUid = body.uid || null;
  const targetEmail = body.email || null;
  const managerEmailRaw = String(body.managerEmail || '').trim();

  if (!targetUid && !targetEmail) {
    return json(event, 400, { ok:false, error:'uid or email required' });
  }

  const target = targetUid ? await findUserByUid(targetUid) : await findUserByEmailLower(targetEmail);
  if (!target) return json(event, 404, { ok:false, error:'Target user not found' });

  // Clear manager mapping
  if (!managerEmailRaw) {
    await db().collection('users').doc(target.uid).set({
      managerEmail: null,
      managerUid: null,
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedBy: user.email
    }, { merge: true });

    return json(event, 200, { ok:true, uid: target.uid, managerEmail: null, managerUid: null });
  }

  const managerEmail = managerEmailRaw;
  const looksLikeEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(managerEmail);
  if (!looksLikeEmail) {
    return json(event, 400, { ok:false, error:'managerEmail is not a valid email address' });
  }

  // Best-effort: resolve managerUid if manager is already in users
  let managerUid = null;
  let managerRole = null;

  const mgr = await findUserByEmailLower(managerEmail);
  if (mgr) {
    managerUid = mgr.uid;
    managerRole = normalizeRole(mgr.role);
    // Only allow PARTNER/MANAGER to be linked as managerUid
    if (!['PARTNER','MANAGER'].includes(managerRole)) {
      managerUid = null;
    }
  }

  await db().collection('users').doc(target.uid).set({
    managerEmail,
    managerUid: managerUid || null,
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedBy: user.email
  }, { merge: true });

  return json(event, 200, {
    ok:true,
    uid: target.uid,
    managerEmail,
    managerUid: managerUid || null,
    note: mgr ? 'manager resolved' : 'manager email saved (user not found yet)'
  });
});