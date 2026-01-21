const { auth, db, json } = require('./_common');

async function requireUser(event) {
  const h = event.headers.authorization || event.headers.Authorization || '';
  const token = h.startsWith('Bearer ') ? h.slice(7) : null;
  if (!token) return { error: json(event, 401, { ok:false, error:'Missing Bearer token' }) };

  const decoded = await auth().verifyIdToken(token);
  const uSnap = await db().collection('users').doc(decoded.uid).get();
  const u = uSnap.exists ? uSnap.data() : null;

  return { user: { uid: decoded.uid, email: decoded.email, role: u?.role || 'WORKER' } };
}

function requirePartner(event, user) {
  if (user.role !== 'PARTNER') return { error: json(event, 403, { ok:false, error:'Partner only' }) };
  return { ok:true };
}

function requireCron(event) {
  const s = event.headers['x-cron-secret'] || event.headers['X-Cron-Secret'];
  if (!s || s !== process.env.CRON_SECRET) return { error: json(event, 401, { ok:false, error:'Bad cron secret' }) };
  return { ok:true };
}

module.exports = { requireUser, requirePartner, requireCron };