// netlify/functions/roles_migrate_worker_to_associate.js
const { withCors, json, db, admin } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

/**
 * One-time migration: users.role WORKER -> ASSOCIATE
 * Partner-only.
 *
 * Body:
 * { dryRun?: boolean, limit?: number }
 */
exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;

  const { user } = authRes;
  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');
  const dryRun = body.dryRun !== false; // default true
  const limit = Math.min(2000, Math.max(1, Number(body.limit || 800)));

  // Find users with role == WORKER
  const snap = await db().collection('users').where('role', '==', 'WORKER').limit(limit).get();
  if (snap.empty) {
    return json(event, 200, { ok:true, dryRun, updated: 0, note: 'No WORKER roles found' });
  }

  const users = snap.docs.map(d => ({ uid: d.id, ...d.data() }));
  if (dryRun) {
    return json(event, 200, {
      ok: true,
      dryRun: true,
      found: users.length,
      sample: users.slice(0, 20).map(u => ({ uid: u.uid, email: u.email }))
    });
  }

  // Batch updates (max 500 per batch)
  let updated = 0;
  let batch = db().batch();
  let ops = 0;

  for (const u of users) {
    const ref = db().collection('users').doc(u.uid);
    batch.set(ref, {
      role: 'ASSOCIATE',
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedBy: user.email
    }, { merge: true });
    ops++;
    updated++;

    if (ops >= 450) {
      await batch.commit();
      batch = db().batch();
      ops = 0;
    }
  }
  if (ops) await batch.commit();

  return json(event, 200, { ok:true, dryRun: false, updated });
});