const { withCors, json, db, admin, auditLog } = require('./_common');
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

  const ref = db().collection('clients').doc();
  await ref.set({
    name: body.name || '',
    pan: body.pan || '',
    gstin: body.gstin || '',
    cin: body.cin || '',
    assessmentYear: body.assessmentYear || body.ay || '',
    engagementType: body.engagementType || body.eng || '',
    primaryEmail: body.primaryEmail || body.email || '',
    ccEmails: asEmailList(body.ccEmails || body.cc),
    bccEmails: asEmailList(body.bccEmails || body.bcc),
    driveFolderId: null,
    createdAt: admin.firestore.FieldValue.serverTimestamp()
  });

  await auditLog({ action:'CLIENT_CREATED', actorUid:user.uid, actorEmail:user.email, details:{ clientId: ref.id } });
  return json(event, 200, { ok:true, clientId: ref.id });
});