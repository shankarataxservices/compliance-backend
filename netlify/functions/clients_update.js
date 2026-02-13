// netlify/functions/clients_update.js
const { withCors, json, db, admin, auditLog } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

function asEmailList(x) {
  if (Array.isArray(x)) return x.map(s => String(s).trim()).filter(Boolean);
  if (typeof x === 'string') return x.split(/[;,:]/).map(s => s.trim()).filter(Boolean);
  return [];
}

function cleanStr(x) {
  return String(x ?? '').trim();
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');
  const clientId = cleanStr(body.clientId);
  if (!clientId) return json(event, 400, { ok:false, error:'clientId required' });

  const ref = db().collection('clients').doc(clientId);
  const snap = await ref.get();
  if (!snap.exists) return json(event, 404, { ok:false, error:'Client not found' });

  const patch = {
    name: cleanStr(body.name),
    pan: cleanStr(body.pan),
    gstin: cleanStr(body.gstin),
    cin: cleanStr(body.cin),
    assessmentYear: cleanStr(body.assessmentYear || body.ay),
    engagementType: cleanStr(body.engagementType || body.eng),
    primaryEmail: cleanStr(body.primaryEmail || body.email),
    ccEmails: asEmailList(body.ccEmails || body.cc),
    bccEmails: asEmailList(body.bccEmails || body.bcc),
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedBy: user.email
  };

  await ref.set(patch, { merge: true });

  await auditLog({
    taskId: null,
    action: 'CLIENT_UPDATED',
    actorUid: user.uid,
    actorEmail: user.email,
    details: { clientId }
  });

  return json(event, 200, { ok:true, clientId });
});