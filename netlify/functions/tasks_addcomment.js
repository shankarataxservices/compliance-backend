// netlify/functions/tasks_addcomment.js
const { withCors, json, db, admin, auditLog } = require('./_common');
const { requireUser } = require('./_auth');

/**
 * Adds a comment to tasks/{taskId}/comments and creates notification docs for @mentions.
 *
 * Request body:
 * { taskId: string, text: string }
 *
 * Mention formats supported:
 * - @email@example.com
 * - @Full Name  (matches users.displayName case-insensitive; spaces allowed until punctuation/end)
 *
 * Notifications collection:
 * notifications/{autoId} with:
 * - toUid
 * - type: 'MENTION'
 * - createdAt
 * - readAt: null
 * - payload: { taskId, taskTitle, byEmail, byName, snippet }
 *
 * Permissions:
 * - PARTNER/MANAGER can comment on any task
 * - ASSOCIATE can comment only on tasks assigned to them
 */

function roleOf(user) {
  let r = String(user?.role || 'ASSOCIATE').toUpperCase().trim() || 'ASSOCIATE';
  if (r === 'WORKER') r = 'ASSOCIATE'; // backward compatibility
  return r;
}
function isPrivileged(role) {
  return role === 'PARTNER' || role === 'MANAGER';
}

function cleanText(x) {
  return String(x || '').replace(/\r\n/g, '\n').trim();
}

function extractMentions(text) {
  const t = String(text || '');

  // 1) emails like @a@b.com
  const emailMentions = [];
  const emailRe = /@([^\s@]+@[^\s@]+\.[^\s@]+)/g;
  let m;
  while ((m = emailRe.exec(t)) !== null) {
    emailMentions.push(m[1]);
  }

  // 2) name mentions like @John Doe (stop at newline or ".,;:!?)(")
  const nameMentions = [];
  const nameRe = /@([A-Za-z][A-Za-z0-9 _-]{1,40})(?=$|[\n\r\t.,;:!?()[\]{}])/g;
  while ((m = nameRe.exec(t)) !== null) {
    const name = String(m[1] || '').trim();
    if (name.includes('@')) continue;
    nameMentions.push(name);
  }

  return {
    emails: [...new Set(emailMentions.map(e => e.trim()).filter(Boolean))],
    names: [...new Set(nameMentions.map(n => n.trim()).filter(Boolean))]
  };
}

async function findUsersByEmails(emails) {
  const out = [];
  const uniq = [...new Set((emails || []).map(e => String(e).trim().toLowerCase()).filter(Boolean))];
  if (!uniq.length) return out;

  // Firestore doesn't support "in" with more than 10 values reliably; chunk
  const chunks = [];
  for (let i = 0; i < uniq.length; i += 10) chunks.push(uniq.slice(i, i + 10));

  for (const c of chunks) {
    const snap = await db().collection('users').where('emailLower', 'in', c).get();
    snap.docs.forEach(d => out.push({ uid: d.id, ...d.data() }));
  }
  return out;
}

async function findUsersByDisplayNames(names) {
  const uniq = [...new Set((names || []).map(n => String(n).trim().toLowerCase()).filter(Boolean))];
  if (!uniq.length) return [];

  // Fast path using displayNameLower
  let out = [];
  try {
    const chunks = [];
    for (let i = 0; i < uniq.length; i += 10) chunks.push(uniq.slice(i, i + 10));
    for (const c of chunks) {
      const snap = await db().collection('users').where('displayNameLower', 'in', c).get();
      snap.docs.forEach(d => out.push({ uid: d.id, ...d.data() }));
    }
    if (out.length) return out;
  } catch {
    // ignore and fallback scan
  }

  // Fallback: scan up to 500 users and match in memory
  const snap = await db().collection('users').orderBy('email', 'asc').limit(500).get();
  const all = snap.docs.map(d => ({ uid: d.id, ...d.data() }));
  out = all.filter(u => {
    const dn = String(u.displayName || '').trim().toLowerCase();
    return dn && uniq.includes(dn);
  });
  return out;
}

async function createNotifications({ toUids, payload }) {
  const uniq = [...new Set((toUids || []).map(x => String(x || '').trim()).filter(Boolean))];
  if (!uniq.length) return 0;

  const ref = db().collection('notifications');
  let created = 0;

  const batches = [];
  for (let i = 0; i < uniq.length; i += 400) batches.push(uniq.slice(i, i + 400));

  for (const group of batches) {
    const b = db().batch();
    for (const uid of group) {
      const docRef = ref.doc();
      b.set(docRef, {
        toUid: uid,
        type: 'MENTION',
        createdAt: admin.firestore.FieldValue.serverTimestamp(),
        readAt: null,
        payload: payload || {}
      });
      created++;
    }
    await b.commit();
  }

  return created;
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok: false, error: 'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const role = roleOf(user);

  const body = JSON.parse(event.body || '{}');
  const taskId = String(body.taskId || '').trim();
  const text = cleanText(body.text);

  if (!taskId) return json(event, 400, { ok: false, error: 'taskId required' });
  if (!text) return json(event, 400, { ok: false, error: 'text required' });
  if (text.length > 5000) return json(event, 400, { ok: false, error: 'text too long (max 5000)' });

  const tRef = db().collection('tasks').doc(taskId);
  const tSnap = await tRef.get();
  if (!tSnap.exists) return json(event, 404, { ok: false, error: 'Task not found' });
  const task = tSnap.data();

  if (!isPrivileged(role)) {
    const isAssignee = task.assignedToUid === user.uid;
    if (!isAssignee) return json(event, 403, { ok: false, error: 'Not allowed' });
  }

  // Load author display name if available
  let authorName = '';
  try {
    const uSnap = await db().collection('users').doc(user.uid).get();
    if (uSnap.exists) {
      const u = uSnap.data() || {};
      authorName = u.displayName || '';
    }
  } catch {}

  // Create comment doc
  const cRef = tRef.collection('comments').doc();
  await cRef.set({
    text,
    authorUid: user.uid,
    authorEmail: user.email || '',
    authorName: authorName || '',
    createdAt: admin.firestore.FieldValue.serverTimestamp()
  });

  await auditLog({
    taskId,
    action: 'COMMENT_ADDED',
    actorUid: user.uid,
    actorEmail: user.email,
    details: { length: text.length }
  });

  // Mentions -> notifications
  const { emails, names } = extractMentions(text);
  const usersByEmail = await findUsersByEmails(emails);
  const usersByName = await findUsersByDisplayNames(names);

  const mentionedUids = [...new Set([
    ...usersByEmail.map(u => u.uid),
    ...usersByName.map(u => u.uid),
  ])].filter(uid => uid && uid !== user.uid);

  const snippet = text.length > 180 ? text.slice(0, 177) + '...' : text;

  const notifCount = await createNotifications({
    toUids: mentionedUids,
    payload: {
      taskId,
      taskTitle: task.title || '',
      byEmail: user.email || '',
      byName: authorName || '',
      snippet
    }
  });

  return json(event, 200, {
    ok: true,
    commentId: cRef.id,
    mentions: { emails, names, notified: mentionedUids.length },
    notificationsCreated: notifCount
  });
});