// netlify/functions/reports_taskHistoryPdf.js
const PDFDocument = require('pdfkit');
const { withCors, json, db, ymdToDmy, IST_TZ } = require('./_common');
const { requireUser } = require('./_auth');

function roleOf(user) {
  let r = String(user?.role || 'ASSOCIATE').toUpperCase().trim();
  if (r === 'WORKER') r = 'ASSOCIATE';
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
  const privileged = isPrivileged(role);

  const body = JSON.parse(event.body || '{}');
  const taskId = String(body.taskId || '').trim();
  if (!taskId) return json(event, 400, { ok:false, error:'taskId required' });

  const tSnap = await db().collection('tasks').doc(taskId).get();
  if (!tSnap.exists) return json(event, 404, { ok:false, error:'Task not found' });
  const task = tSnap.data();

  if (!privileged && task.assignedToUid !== user.uid) {
    return json(event, 403, { ok:false, error:'Not allowed' });
  }

  const cSnap = task.clientId ? await db().collection('clients').doc(task.clientId).get() : null;
  const client = cSnap && cSnap.exists ? (cSnap.data() || {}) : {};

  const aSnap = await db().collection('auditLogs').where('taskId', '==', taskId).get();
  const logs = aSnap.docs.map(d => d.data())
    .sort((x,y)=> (x.timestamp?.toMillis?.()||0) - (y.timestamp?.toMillis?.()||0));

  const comSnap = await db().collection('tasks').doc(taskId).collection('comments')
    .orderBy('createdAt', 'asc')
    .limit(400)
    .get();
  const comments = comSnap.docs.map(d => d.data());

  const attachments = Array.isArray(task.attachments) ? task.attachments : [];

  // PDF
  const doc = new PDFDocument({ size: 'A4', margin: 40 });
  const chunks = [];
  doc.on('data', (c) => chunks.push(c));
  const done = new Promise((resolve) => doc.on('end', resolve));

  doc.fontSize(16).text('Task History Report');
  doc.moveDown(0.2);

  doc.fontSize(11).text(task.title || 'Task');
  doc.fontSize(9).fillColor('#444')
    .text(`Client: ${client.name || ''}`)
    .text(`Start: ${ymdToDmy(task.startDateYmd)}   Due: ${ymdToDmy(task.dueDateYmd)}   Status: ${task.status || ''}`)
    .text(`Assignee: ${task.assignedToEmail || ''}`)
    .text(`Generated: ${new Date().toLocaleString('en-IN', { timeZone: IST_TZ })}`);
  doc.fillColor('#000');
  doc.moveDown(0.8);

  doc.fontSize(12).text('Timeline', { underline: true });
  doc.moveDown(0.2);
  if (!logs.length) {
    doc.fontSize(10).fillColor('#666').text('No timeline entries.');
    doc.fillColor('#000');
  } else {
    for (const a of logs) {
      const when = a.timestamp?.toDate?.()
        ? a.timestamp.toDate().toLocaleString('en-IN', { timeZone: IST_TZ })
        : '';
      doc.fontSize(10).text(`${when} — ${a.action || ''}`);
      if (a.actorEmail) doc.fontSize(9).fillColor('#444').text(a.actorEmail);
      doc.fillColor('#666').fontSize(9).text(JSON.stringify(a.details || {}));
      doc.fillColor('#000');
      doc.moveDown(0.4);
      if (doc.y > 760) doc.addPage();
    }
  }

  doc.moveDown(0.6);
  doc.fontSize(12).text('Comments', { underline: true });
  doc.moveDown(0.2);
  if (!comments.length) {
    doc.fontSize(10).fillColor('#666').text('No comments.');
    doc.fillColor('#000');
  } else {
    for (const c of comments) {
      const when = c.createdAt?.toDate?.()
        ? c.createdAt.toDate().toLocaleString('en-IN', { timeZone: IST_TZ })
        : '';
      const who = c.authorName || c.authorEmail || '';
      doc.fontSize(10).text(`${when} — ${who}`);
      doc.fontSize(9).fillColor('#444').text(String(c.text || ''), { width: 520 });
      doc.fillColor('#000');
      doc.moveDown(0.4);
      if (doc.y > 760) doc.addPage();
    }
  }

  doc.moveDown(0.6);
  doc.fontSize(12).text('Attachments', { underline: true });
  doc.moveDown(0.2);
  if (!attachments.length) {
    doc.fontSize(10).fillColor('#666').text('No attachments.');
    doc.fillColor('#000');
  } else {
    for (const a of attachments) {
      doc.fontSize(10).text(`${a.type || ''} — ${a.fileName || ''}`);
      if (a.driveWebViewLink) doc.fontSize(9).fillColor('#1a73e8').text(a.driveWebViewLink);
      doc.fillColor('#000');
      doc.moveDown(0.3);
      if (doc.y > 760) doc.addPage();
    }
  }

  doc.end();
  await done;

  const safeTitle = String(task.title || 'task').replace(/[^\w\- ]+/g, '').slice(0, 40).trim() || 'task';
  return json(event, 200, {
    ok: true,
    fileName: `task_history_${safeTitle}_${taskId}.pdf`,
    base64: Buffer.concat(chunks).toString('base64'),
    mime: 'application/pdf'
  });
});