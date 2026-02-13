// netlify/functions/tasks_delete.js
const { withCors, json, db, calendar, auditLog } = require('./_common');
const { requireUser } = require('./_auth');

async function deleteEvent(eventId) {
  if (!eventId) return;
  try {
    await calendar().events.delete({
      calendarId: 'primary',
      eventId,
      sendUpdates: 'none'
    });
  } catch {
    // ignore if already deleted / not found
  }
}

function roleOf(user) {
  let r = String(user?.role || 'ASSOCIATE').toUpperCase().trim();
  if (r === 'WORKER') r = 'ASSOCIATE'; // compat
  return r || 'ASSOCIATE';
}
function isPrivileged(role) {
  return role === 'PARTNER' || role === 'MANAGER';
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const role = roleOf(user);

  const body = JSON.parse(event.body || '{}');
  const taskId = String(body.taskId || '').trim();
  const applyToSeries = !!body.applyToSeries;

  if (!taskId) return json(event, 400, { ok:false, error:'taskId required' });

  const baseRef = db().collection('tasks').doc(taskId);
  const baseSnap = await baseRef.get();
  if (!baseSnap.exists) return json(event, 404, { ok:false, error:'Task not found' });
  const base = baseSnap.data();

  // Permission:
  // - PARTNER/MANAGER: can delete any task, series allowed
  // - ASSOCIATE: can delete only tasks assigned to them; series delete NOT allowed
  if (!isPrivileged(role)) {
    const isAssignee = base.assignedToUid === user.uid;
    if (!isAssignee) return json(event, 403, { ok:false, error:'Not allowed' });
    if (applyToSeries) return json(event, 403, { ok:false, error:'Associates cannot delete entire series' });
  }

  let targets = [];
  if (applyToSeries && base.seriesId) {
    const snap = await db().collection('tasks').where('seriesId', '==', base.seriesId).get();
    targets = snap.docs.map(d => ({ id: d.id, ref: d.ref, data: d.data() }));
  } else {
    targets = [{ id: taskId, ref: baseRef, data: base }];
  }

  // For associates, enforce again for every target
  if (!isPrivileged(role)) {
    const bad = targets.find(t => t.data.assignedToUid !== user.uid);
    if (bad) return json(event, 403, { ok:false, error:'Not allowed on some tasks' });
  }

  for (const t of targets) {
    const singleId = t.data.calendarEventId || t.data.calendarStartEventId || null;
    await deleteEvent(singleId);
    await deleteEvent(t.data.calendarDueEventId || null);

    await t.ref.delete();

    await auditLog({
      taskId: t.id,
      action: 'TASK_DELETED',
      actorUid: user.uid,
      actorEmail: user.email,
      details: { applyToSeries: !!applyToSeries, seriesId: base.seriesId || null, role }
    });
  }

  if (applyToSeries && base.seriesId) {
    try { await db().collection('taskSeries').doc(base.seriesId).delete(); } catch {}
  }

  return json(event, 200, { ok:true, deletedCount: targets.length });
});