// netlify/functions/tasks_bulkupdate_from_xlsx.js
const Busboy = require('busboy');
const ExcelJS = require('exceljs');

const {
  withCors, json, db, admin,
  dmyToYmd, ymdIST, addDays, dateFromYmdIST,
  getCalendarWindow, calTimeRange,
  auditLog, calendar
} = require('./_common');
const { requireUser } = require('./_auth');

function roleOf(user) {
  let r = String(user?.role || 'ASSOCIATE').toUpperCase().trim();
  if (r === 'WORKER') r = 'ASSOCIATE';
  return r || 'ASSOCIATE';
}
function isPrivileged(role) {
  return role === 'PARTNER' || role === 'MANAGER';
}

function asEmailListLoose(x) {
  if (Array.isArray(x)) return x.map(s => String(s).trim()).filter(Boolean);
  if (typeof x === 'string') return x.split(/[;,:]/).map(s => s.trim()).filter(Boolean);
  return [];
}

function getCellString(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (v == null) return '';
  if (typeof v === 'string') return v.trim();
  if (typeof v === 'number') return String(v);
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'object') {
    if (v.text) return String(v.text).trim();
    if (v.richText && Array.isArray(v.richText)) return v.richText.map(r => r.text || '').join('').trim();
  }
  return String(v).trim();
}
function truthy(x) {
  const s = String(x ?? '').trim().toLowerCase();
  if (!s) return false;
  return ['1','true','yes','y'].includes(s);
}
function parseDmyCell(cell) {
  if (!cell || cell.value == null) return '';
  const v = cell.value;
  if (v instanceof Date) {
    const yyyy = v.getFullYear();
    const mm = String(v.getMonth() + 1).padStart(2,'0');
    const dd = String(v.getDate()).padStart(2,'0');
    return `${dd}-${mm}-${yyyy}`;
  }
  return getCellString(cell);
}

async function patchEvent({ eventId, whenYmd, summary, description, colorId=null, window }) {
  if (!eventId) return;
  const cal = calendar();
  const range = calTimeRange(whenYmd, window.startHH, window.endHH, window.timeZone);
  await cal.events.patch({
    calendarId: 'primary',
    eventId,
    sendUpdates: 'none',
    requestBody: {
      summary,
      description,
      ...range,
      ...(colorId ? { colorId } : {})
    }
  });
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const role = roleOf(user);
  const privileged = isPrivileged(role);

  // Parse multipart
  const busboy = Busboy({ headers: event.headers });
  let fileBuffer = Buffer.alloc(0);

  busboy.on('file', (name, file) => {
    file.on('data', (d) => { fileBuffer = Buffer.concat([fileBuffer, d]); });
  });

  const done = new Promise((resolve, reject) => {
    busboy.on('finish', resolve);
    busboy.on('error', reject);
  });

  const bodyBuf = event.isBase64Encoded ? Buffer.from(event.body, 'base64') : Buffer.from(event.body || '');
  busboy.end(bodyBuf);
  await done;

  if (!fileBuffer.length) return json(event, 400, { ok:false, error:'XLSX file missing' });

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(fileBuffer);
  const ws = wb.getWorksheet('Update') || wb.worksheets[0];
  if (!ws) return json(event, 400, { ok:false, error:'No worksheet found' });

  // Header map
  const headerRow = ws.getRow(1);
  const headers = {};
  headerRow.eachCell((cell, colNumber) => {
    const key = getCellString(cell);
    if (key) headers[key] = colNumber;
  });
  const colOf = (names) => {
    for (const n of names) { if (headers[n]) return headers[n]; }
    return null;
  };

  const cTaskId = colOf(['TaskId']);
  if (!cTaskId) return json(event, 400, { ok:false, error:'Missing required column: TaskId' });

  // Optional columns
  const cTitle = colOf(['Title']);
  const cCategory = colOf(['Category']);
  const cType = colOf(['Type']);
  const cPriority = colOf(['Priority']);
  const cDue = colOf(['DueDate (DD-MM-YYYY)', 'DueDate']);
  const cTrigger = colOf(['TriggerDays']);
  const cAssigned = colOf(['AssignedToEmail']);
  const cStatus = colOf(['Status']);
  const cStatusNote = colOf(['StatusNote']);
  const cDelayReason = colOf(['DelayReason']);
  const cDelayNotes = colOf(['DelayNotes']);
  const cSnooze = colOf(['SnoozedUntil (DD-MM-YYYY)', 'SnoozedUntil']);

  const cSendStart = colOf(['SendStartMail (true/false)', 'SendStartMail']);
  const cTo = colOf(['ClientTo (emails ; , : separated)', 'ClientTo']);
  const cCc = colOf(['ClientCC (emails ; , : separated)', 'ClientCC']);
  const cBcc = colOf(['ClientBCC (emails ; , : separated)', 'ClientBCC']);
  const cCcAssignee = colOf(['CcAssigneeOnClientStart (true/false)', 'CcAssigneeOnClientStart']);
  const cCcManager = colOf(['CcManagerOnClientStart (true/false)', 'CcManagerOnClientStart']);
  const cStartSub = colOf(['ClientStartSubject']);
  const cStartBody = colOf(['ClientStartBody']);

  const cSendComp = colOf(['SendClientCompletionMail (true/false)', 'SendClientCompletionMail']);
  const cCompTo = colOf(['CompletionTo (emails ; , : separated)', 'CompletionTo']);
  const cCompCc = colOf(['CompletionCC (emails ; , : separated)', 'CompletionCC']);
  const cCompBcc = colOf(['CompletionBCC (emails ; , : separated)', 'CompletionBCC']);
  const cCcAssigneeComp = colOf(['CcAssigneeOnCompletion (true/false)', 'CcAssigneeOnCompletion']);
  const cCcManagerComp = colOf(['CcManagerOnCompletion (true/false)', 'CcManagerOnCompletion']);
  const cCompSub = colOf(['ClientCompletionSubject']);
  const cCompBody = colOf(['ClientCompletionBody']);

  const startRow = 2;
  const lastRow = ws.lastRow ? ws.lastRow.number : 1;

  const window = await getCalendarWindow();

  let updatedCount = 0;
  let skippedCount = 0;
  const errors = [];

  // Limit to prevent abuse
  const maxUpdates = 800;

  for (let r = startRow; r <= lastRow; r++) {
    if (updatedCount + skippedCount >= maxUpdates) break;

    const row = ws.getRow(r);
    const taskId = String(getCellString(row.getCell(cTaskId)) || '').trim();
    if (!taskId) continue;

    const ref = db().collection('tasks').doc(taskId);
    const snap = await ref.get();
    if (!snap.exists) {
      skippedCount++;
      errors.push({ row: r, taskId, error: 'Task not found' });
      continue;
    }

    const old = snap.data();

    // Permission: associate can update only own tasks; and only safe fields
    if (!privileged) {
      if (old.assignedToUid !== user.uid) {
        skippedCount++;
        errors.push({ row: r, taskId, error: 'Not allowed (not your task)' });
        continue;
      }
    }

    const patch = {};
    const changed = [];

    // Helpers to set patch only if cell not blank
    const setIf = (field, value, allowEmptyString = false) => {
      if (value === undefined) return;
      if (!allowEmptyString && value === '') return;
      patch[field] = value;
      changed.push(field);
    };

    // Fields (privileged only for many)
    if (cTitle && privileged) setIf('title', getCellString(row.getCell(cTitle)));
    if (cCategory && privileged) setIf('category', getCellString(row.getCell(cCategory)));
    if (cType && privileged) setIf('type', getCellString(row.getCell(cType)));
    if (cPriority) setIf('priority', (getCellString(row.getCell(cPriority)) || '').toUpperCase());

    // Due date + trigger affects startDate + calendar patch (privileged only)
    let dueDateYmd = null;
    if (cDue && privileged) {
      const dueDmy = parseDmyCell(row.getCell(cDue));
      if (dueDmy) {
        try {
          dueDateYmd = dmyToYmd(dueDmy);
          patch.dueDateYmd = dueDateYmd;
          patch.dueDate = admin.firestore.Timestamp.fromDate(dateFromYmdIST(dueDateYmd));
          changed.push('dueDateYmd', 'dueDate');
        } catch (e) {
          skippedCount++;
          errors.push({ row: r, taskId, error: `Invalid due date: ${e.message}` });
          continue;
        }
      }
    }

    let triggerDaysBefore = null;
    if (cTrigger && privileged) {
      const v = getCellString(row.getCell(cTrigger));
      if (v) {
        const n = Math.max(0, parseInt(v, 10));
        triggerDaysBefore = n;
        patch.triggerDaysBefore = n;
        changed.push('triggerDaysBefore');
      }
    }

    // Reassign (privileged only)
    if (cAssigned && privileged) {
      const email = getCellString(row.getCell(cAssigned)).trim();
      if (email) {
        // best-effort lookup
        const q = await db().collection('users').where('emailLower', '==', email.toLowerCase()).limit(1).get();
        if (q.empty) {
          skippedCount++;
          errors.push({ row: r, taskId, error: 'AssignedToEmail not found in users' });
          continue;
        }
        patch.assignedToEmail = email;
        patch.assignedToUid = q.docs[0].id;
        changed.push('assignedToEmail', 'assignedToUid');
      }
    }

    // Status + notes (associates allowed, but COMPLETED only privileged)
    if (cStatus) {
      const st = getCellString(row.getCell(cStatus)).toUpperCase().trim();
      if (st) {
        // Associates can mark COMPLETED if the task is assigned to them (checked above)
        patch.status = st;
        changed.push('status');
        if (st === 'APPROVAL_PENDING') {
          patch.completedRequestedAt = admin.firestore.FieldValue.serverTimestamp();
          changed.push('completedRequestedAt');
        }
        if (st === 'COMPLETED') {
          patch.completedAt = admin.firestore.FieldValue.serverTimestamp();
          changed.push('completedAt');
        }
      }
    }

    if (cStatusNote) setIf('statusNote', getCellString(row.getCell(cStatusNote)), true);
    if (cDelayReason) setIf('delayReason', getCellString(row.getCell(cDelayReason)), true);
    if (cDelayNotes) setIf('delayNotes', getCellString(row.getCell(cDelayNotes)), true);

    // Snooze (associates allowed)
    if (cSnooze) {
      const snoozeDmy = parseDmyCell(row.getCell(cSnooze));
      if (snoozeDmy) {
        try {
          const y = dmyToYmd(snoozeDmy);
          patch.snoozedUntilYmd = y;
          changed.push('snoozedUntilYmd');
        } catch (e) {
          skippedCount++;
          errors.push({ row: r, taskId, error: `Invalid snooze date: ${e.message}` });
          continue;
        }
      }
    }

    // Mail fields (privileged only)
    if (privileged) {
      if (cSendStart) {
        const v = getCellString(row.getCell(cSendStart));
        if (v) { patch.sendClientStartMail = truthy(v); changed.push('sendClientStartMail'); }
      }
      if (cTo) { const v = getCellString(row.getCell(cTo)); if (v) { patch.clientToEmails = asEmailListLoose(v); changed.push('clientToEmails'); } }
      if (cCc) { const v = getCellString(row.getCell(cCc)); if (v) { patch.clientCcEmails = asEmailListLoose(v); changed.push('clientCcEmails'); } }
      if (cBcc) { const v = getCellString(row.getCell(cBcc)); if (v) { patch.clientBccEmails = asEmailListLoose(v); changed.push('clientBccEmails'); } }
      if (cCcAssignee) { const v = getCellString(row.getCell(cCcAssignee)); if (v) { patch.ccAssigneeOnClientStart = truthy(v); changed.push('ccAssigneeOnClientStart'); } }
      if (cCcManager) { const v = getCellString(row.getCell(cCcManager)); if (v) { patch.ccManagerOnClientStart = truthy(v); changed.push('ccManagerOnClientStart'); } }
      if (cStartSub) { const v = getCellString(row.getCell(cStartSub)); if (v) { patch.clientStartSubject = v; changed.push('clientStartSubject'); } }
      if (cStartBody) { const v = getCellString(row.getCell(cStartBody)); if (v) { patch.clientStartBody = v; changed.push('clientStartBody'); } }

      if (cSendComp) {
        const v = getCellString(row.getCell(cSendComp));
        if (v) { patch.sendClientCompletionMail = truthy(v); changed.push('sendClientCompletionMail'); }
      }
      if (cCompTo) { const v = getCellString(row.getCell(cCompTo)); if (v) { patch.completionToEmails = asEmailListLoose(v); changed.push('completionToEmails'); } }
      if (cCompCc) { const v = getCellString(row.getCell(cCompCc)); if (v) { patch.completionCcEmails = asEmailListLoose(v); changed.push('completionCcEmails'); } }
      if (cCompBcc) { const v = getCellString(row.getCell(cCompBcc)); if (v) { patch.completionBccEmails = asEmailListLoose(v); changed.push('completionBccEmails'); } }
      if (cCcAssigneeComp) { const v = getCellString(row.getCell(cCcAssigneeComp)); if (v) { patch.ccAssigneeOnCompletion = truthy(v); changed.push('ccAssigneeOnCompletion'); } }
      if (cCcManagerComp) { const v = getCellString(row.getCell(cCcManagerComp)); if (v) { patch.ccManagerOnCompletion = truthy(v); changed.push('ccManagerOnCompletion'); } }
      if (cCompSub) { const v = getCellString(row.getCell(cCompSub)); if (v) { patch.clientCompletionSubject = v; changed.push('clientCompletionSubject'); } }
      if (cCompBody) { const v = getCellString(row.getCell(cCompBody)); if (v) { patch.clientCompletionBody = v; changed.push('clientCompletionBody'); } }
    }

    if (!changed.length) {
      skippedCount++;
      continue;
    }

    // If dueDateYmd or triggerDaysBefore changed, recompute startDateYmd and patch calendar event
    const finalDueYmd = dueDateYmd || old.dueDateYmd;
    const finalTrigger = (triggerDaysBefore !== null && triggerDaysBefore !== undefined)
      ? triggerDaysBefore
      : (Number(old.triggerDaysBefore ?? 15));

    if (privileged && (dueDateYmd || triggerDaysBefore !== null)) {
      const startYmd = ymdIST(addDays(dateFromYmdIST(finalDueYmd), -finalTrigger));
      patch.startDateYmd = startYmd;
      patch.startDate = admin.firestore.Timestamp.fromDate(dateFromYmdIST(startYmd));
      changed.push('startDateYmd', 'startDate');

      // Patch calendar event
      try {
        const prefix = (old.status === 'COMPLETED') ? '[COMPLETED] ' : '';
        const summary = `${prefix}START: ${patch.title || old.title || 'Task'}`;
        const cn = String(old.clientNameSnapshot || '').trim() || String(old.clientId || '').trim();
        const extra = String(old.calendarDescription || '').trim();
        const descBase = `Client: ${cn}\nStart: ${startYmd}\nDue: ${finalDueYmd}\n`;
        const desc = extra ? `${descBase}\n${extra}` : descBase;
        const eventId = old.calendarEventId || old.calendarStartEventId || null;

        await patchEvent({
          eventId,
          whenYmd: startYmd,
          summary,
          description: desc,
          colorId: (old.status === 'COMPLETED') ? '2' : null,
          window
        });

        if (old.calendarDueEventId) {
          await patchEvent({
            eventId: old.calendarDueEventId,
            whenYmd: finalDueYmd,
            summary: `${prefix}DUE: ${patch.title || old.title || 'Task'}`,
            description: desc,
            colorId: (old.status === 'COMPLETED') ? '2' : null,
            window
          });
        }
      } catch (e) {
        errors.push({ row: r, taskId, error: `Calendar patch failed (ignored): ${e.message || String(e)}` });
      }
    }

    patch.updatedAt = admin.firestore.FieldValue.serverTimestamp();

    await ref.update(patch);

    await auditLog({
      taskId,
      action: 'TASK_OFFLINE_UPDATE',
      actorUid: user.uid,
      actorEmail: user.email,
      details: { fields: changed }
    });

    updatedCount++;
  }

  return json(event, 200, {
    ok: true,
    updatedCount,
    skippedCount,
    errors: errors.slice(0, 200)
  });
});