const { withCors, json, db, admin } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

function clampHH(x, def) {
  const n = Number(x);
  if (!Number.isFinite(n)) return def;
  return Math.min(23, Math.max(0, Math.floor(n)));
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');

  const startHH = clampHH(body.startHH, 10);
  const endHH = clampHH(body.endHH, 12);
  if (endHH <= startHH) {
    return json(event, 400, { ok:false, error:'endHH must be > startHH' });
  }

  const timeZone = String(body.timeZone || 'Asia/Kolkata').trim() || 'Asia/Kolkata';

  await db().collection('settings').doc('calendar').set({
    startHH,
    endHH,
    timeZone,
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedBy: user.email
  }, { merge: true });

  return json(event, 200, { ok:true });
});