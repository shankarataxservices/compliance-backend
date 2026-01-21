const { withCors, json, db, admin } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

function asEmailList(x) {
  if (Array.isArray(x)) return x.map(s => String(s).trim()).filter(Boolean);
  if (typeof x === 'string') return x.split(/[;,]/).map(s => s.trim()).filter(Boolean);
  return [];
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');

  const doc = {
    dailyInternalEmails: asEmailList(body.dailyInternalEmails),
    dailyWindowDays: Number(body.dailyWindowDays || 30),
    sendDailyToAssignees: body.sendDailyToAssignees !== false,
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedBy: user.email
  };

  await db().collection('settings').doc('notifications').set(doc, { merge: true });
  return json(event, 200, { ok:true });
});