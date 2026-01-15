const { json, db, admin, auditLog } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: require('./_common').cors() };
  if (event.httpMethod !== 'POST') return json(405, { error: 'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');

  const ref = db().collection('clients').doc();
  await ref.set({
    name: body.name || '',
    pan: body.pan || '',
    gstin: body.gstin || '',
    cin: body.cin || '',
    assessmentYear: body.assessmentYear || '',
    engagementType: body.engagementType || '',
    primaryEmail: body.primaryEmail || '',
    ccEmails: body.ccEmails || [],
    bccEmails: body.bccEmails || [],
    driveFolderId: null,
    createdAt: admin.firestore.FieldValue.serverTimestamp()
  });

  await auditLog({ action: 'CLIENT_CREATED', actorUid: user.uid, actorEmail: user.email, details: { clientId: ref.id } });
  return json(200, { ok: true, clientId: ref.id });
};
