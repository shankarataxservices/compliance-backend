// netlify/functions/settings_workerImportPassword_set.js
const crypto = require('crypto');
const { withCors, json, db, admin } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');

/**
 * We store a PBKDF2 hash so we never store plaintext password.
 * Doc: settings/security
 * Fields:
 *  - workerImportEnabled: boolean
 *  - workerImportHash: string (base64)
 *  - workerImportSalt: string (base64)
 *  - workerImportIter: number
 *  - updatedAt, updatedBy
 */

function pbkdf2Hash(password, saltBuf, iter = 120000) {
  const keyLen = 32;
  const digest = 'sha256';
  const derived = crypto.pbkdf2Sync(password, saltBuf, iter, keyLen, digest);
  return derived;
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');
  const password = String(body.password || '');

  const ref = db().collection('settings').doc('security');

  // Remove/disable password
  if (!password) {
    await ref.set({
      workerImportEnabled: false,
      workerImportHash: null,
      workerImportSalt: null,
      workerImportIter: null,
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedBy: user.email
    }, { merge: true });

    return json(event, 200, { ok:true, enabled:false });
  }

  if (password.length < 4) {
    return json(event, 400, { ok:false, error:'Password too short (min 4 characters)' });
  }

  const salt = crypto.randomBytes(16);
  const iter = 120000;
  const hash = pbkdf2Hash(password, salt, iter);

  await ref.set({
    workerImportEnabled: true,
    workerImportHash: hash.toString('base64'),
    workerImportSalt: salt.toString('base64'),
    workerImportIter: iter,
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedBy: user.email
  }, { merge: true });

  return json(event, 200, { ok:true, enabled:true });
});