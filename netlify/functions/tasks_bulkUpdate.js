// netlify/functions/tasks_bulkUpdate.js
const { withCors, json, db, admin, auditLog, calendar } = require('./_common');
const { requireUser } = require('./_auth');

const ACTIVE_STATUSES = new Set(['PENDING','IN_PROGRESS','CLIENT_PENDING','APPROVAL_PENDING','COMPLETED']);

function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

function mustArrayIds(x) {
  if (!Array.isArray(x)) return [];
  return x.map(v => String(v || '').trim()).filter(Boolean);
}

function normEmail(x) {
  return String(x || '').trim();
}

function roleOf(user) {
  let r = String(user?.role || 'ASSOCIATE').toUpperCase().trim();
  if (r === 'WORKER') r = 'ASSOCIATE'; // compat
  return r || 'ASSOCIATE';
}

function isPrivileged(role) {
  return role === 'PARTNER' || role === 'MANAGER';
}

async function findUserUidByEmail(email) {
  const e = String(email || '').trim().toLowerCase();
  if (!e) return null;
  let snap = await db().collection('users').where('emailLower', '==', e).limit(1).get();
  if (snap.empty) snap = await db().collection('users').where('email', '==', email).limit(1).get();
  if (snap.empty) return null;
  return snap.docs[0].id;
}

async function safeDeleteCalendarEvent(eventId) {
  if (!eventId) return { ok:false, skipped:true };
  try {
    await calendar().events.delete({
      calendarId: 'primary',
      eventId,
      sendUpdates: 'none'
    });
    return { ok:true };
  } catch (e) {
    return { ok:false, error: e?.message || String(e) };
  }
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const body = JSON.parse(event.body || '{}');
  const op = String(body.op || '').toUpperCase().trim();
  const taskIds = mustArrayIds(body.taskIds);

  if (!op) return json(event, 400, { ok:false, error:'op required' });
  if (!taskIds.length) return json(event, 400, { ok:false, error:'taskIds required' });
  if (taskIds.length > 2000) return json(event, 400, { ok:false, error:'Too many tasks (max 2000)' });

  const role = roleOf(user);
  const privileged = isPrivileged(role);

  // Load tasks
  const refs = taskIds.map(id => db().collection('tasks').doc(id));
  const snaps = await db().getAll(...refs);
  const tasks = snaps.filter(s => s.exists).map(s => ({ id: s.id, ref: s.ref, data: s.data() }));

  if (!tasks.length) return json(event, 404, { ok:false, error:'No tasks found' });

  // Associate: only their tasks
  if (!privileged) {
    const notMine = tasks.filter(t => t.data.assignedToUid !== user.uid);
    if (notMine.length) return json(event, 403, { ok:false, error:'Some selected tasks are not assigned to you' });
  }

  let updatePatch = null;
  let updatedCount = 0;
  let deletedCount = 0;

  if (op === 'STATUS') {
    const newStatus = String(body.newStatus || '').toUpperCase().trim();
    if (!ACTIVE_STATUSES.has(newStatus)) return json(event, 400, { ok:false, error:'Invalid newStatus' });

    // Associates can set COMPLETED for tasks assigned to them (ownership enforced earlier)

    updatePatch = () => {
      const patch = {
        status: newStatus,
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
      };
      if (newStatus === 'APPROVAL_PENDING') patch.completedRequestedAt = admin.firestore.FieldValue.serverTimestamp();
      if (newStatus === 'COMPLETED') patch.completedAt = admin.firestore.FieldValue.serverTimestamp();
      return patch;
    };

  } else if (op === 'REASSIGN') {
    const assignedToEmail = normEmail(body.assignedToEmail);
    if (!assignedToEmail) return json(event, 400, { ok:false, error:'assignedToEmail required' });

    const assignedToUid = await findUserUidByEmail(assignedToEmail);
    if (!assignedToUid) return json(event, 400, { ok:false, error:'assignedToEmail not found in users' });

    // Associates cannot reassign
    if (!privileged) return json(event, 403, { ok:false, error:'Only Partner/Manager can reassign' });

    updatePatch = () => ({
      assignedToEmail,
      assignedToUid,
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

  } else if (op === 'SNOOZE') {
    const snoozedUntilYmd = String(body.snoozedUntilYmd || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(snoozedUntilYmd)) {
      return json(event, 400, { ok:false, error:'snoozedUntilYmd must be YYYY-MM-DD' });
    }

    updatePatch = () => ({
      snoozedUntilYmd,
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

  } else if (op === 'DELETE') {
    // handled below

  } else {
    return json(event, 400, { ok:false, error:'Invalid op' });
  }

  // Apply in batches
  const batches = chunk(tasks, 350);

  for (const group of batches) {
    if (op === 'DELETE') {
      for (const t of group) {
        const singleId = t.data.calendarEventId || t.data.calendarStartEventId || null;
        const dueId = t.data.calendarDueEventId || null;

        await safeDeleteCalendarEvent(singleId);
        await safeDeleteCalendarEvent(dueId);

        await t.ref.delete();
        deletedCount++;

        await auditLog({
          taskId: t.id,
          action: 'TASK_DELETED_BULK',
          actorUid: user.uid,
          actorEmail: user.email,
          details: { op:'DELETE' }
        });

        await new Promise(res => setTimeout(res, 40));
      }
      continue;
    }

    const batch = db().batch();
    for (const t of group) batch.update(t.ref, updatePatch(t));
    await batch.commit();
    updatedCount += group.length;

    for (const t of group) {
      await auditLog({
        taskId: t.id,
        action:
          op === 'STATUS' ? 'BULK_STATUS_CHANGE' :
          op === 'REASSIGN' ? 'BULK_REASSIGN' :
          op === 'SNOOZE' ? 'BULK_SNOOZE' :
          'BULK_UPDATE',
        actorUid: user.uid,
        actorEmail: user.email,
        details: {
          op,
          ...(op === 'STATUS' ? { to: String(body.newStatus||'') } : {}),
          ...(op === 'REASSIGN' ? { assignedToEmail: String(body.assignedToEmail||'') } : {}),
          ...(op === 'SNOOZE' ? { snoozedUntilYmd: String(body.snoozedUntilYmd||'') } : {})
        }
      });
    }
  }

  return json(event, 200, {
    ok: true,
    op,
    requested: taskIds.length,
    found: tasks.length,
    updatedCount,
    deletedCount
  });
});