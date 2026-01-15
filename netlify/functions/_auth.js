const { auth, db, json } = require('./_common');

async function requireUser(event) {
  const h = event.headers.authorization || event.headers.Authorization || '';
  const token = h.startsWith('Bearer ') ? h.slice(7) : null;
  if (!token) return { error: json(401, { error: 'Missing Bearer token' }) };

  const decoded = await auth().verifyIdToken(token);
  const uSnap = await db().collection('users').doc(decoded.uid).get();
  const u = uSnap.exists ? uSnap.data() : null;

  return { user: { uid: decoded.uid, email: decoded.email, role: u?.role || 'WORKER' } };
}

function requirePartner(user) {
  if (user.role !== 'PARTNER') return { error: json(403, { error: 'Partner only' }) };
  return { ok: true };
}

function requireCron(event) {
  const s = event.headers['x-cron-secret'] || event.headers['X-Cron-Secret'];
  if (!s || s !== process.env.CRON_SECRET) return { error: json(401, { error: 'Bad cron secret' }) };
  return { ok: true };
}

module.exports = { requireUser, requirePartner, requireCron };
