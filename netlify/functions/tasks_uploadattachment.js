// netlify/functions/tasks_uploadattachment.js
const Busboy = require('busboy');
const { withCors, json, db, admin, drive, auditLog } = require('./_common');
const { requireUser } = require('./_auth');

function roleOf(user) {
  let r = String(user?.role || 'ASSOCIATE').toUpperCase().trim();
  if (r === 'WORKER') r = 'ASSOCIATE'; // compat
  return r || 'ASSOCIATE';
}
function isPrivileged(role) {
  return role === 'PARTNER' || role === 'MANAGER';
}

async function ensureRootFolder() {
  const d = drive();
  const name = process.env.DRIVE_ROOT_FOLDER_NAME || 'ComplianceManagement';
  const res = await d.files.list({
    q: `mimeType='application/vnd.google-apps.folder' and name='${name}' and trashed=false`,
    fields: 'files(id,name)',
    spaces: 'drive'
  });
  if (res.data.files && res.data.files.length) return res.data.files[0].id;

  const created = await d.files.create({
    requestBody: { name, mimeType: 'application/vnd.google-apps.folder' },
    fields: 'id'
  });
  return created.data.id;
}

async function ensureClientFolder(clientId) {
  const cRef = db().collection('clients').doc(clientId);
  const cSnap = await cRef.get();
  const client = cSnap.data() || {};

  if (client.driveFolderId) return client.driveFolderId;

  const rootId = await ensureRootFolder();
  const folderName = `${client.name || 'Client'}_${clientId}`;
  const d = drive();

  const created = await d.files.create({
    requestBody: {
      name: folderName,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [rootId]
    },
    fields: 'id'
  });

  await cRef.update({ driveFolderId: created.data.id });
  return created.data.id;
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const role = roleOf(user);

  const busboy = Busboy({ headers: event.headers });
  let taskId = null;
  let attachmentType = 'OTHER';
  let fileBuffer = Buffer.alloc(0);
  let fileName = 'file';
  let mimeType = 'application/octet-stream';

  busboy.on('field', (name, val) => {
    if (name === 'taskId') taskId = String(val || '').trim();
    if (name === 'type') attachmentType = String(val || 'OTHER').toUpperCase().trim();
  });

  busboy.on('file', (name, file, info) => {
    fileName = info.filename || 'file';
    mimeType = info.mimeType || 'application/octet-stream';
    file.on('data', (d) => { fileBuffer = Buffer.concat([fileBuffer, d]); });
  });

  const done = new Promise((resolve, reject) => {
    busboy.on('finish', resolve);
    busboy.on('error', reject);
  });

  const bodyBuf = event.isBase64Encoded ? Buffer.from(event.body, 'base64') : Buffer.from(event.body || '');
  busboy.end(bodyBuf);
  await done;

  if (!taskId) return json(event, 400, { ok:false, error:'taskId required' });
  if (!fileBuffer.length) return json(event, 400, { ok:false, error:'file missing' });

  const tRef = db().collection('tasks').doc(taskId);
  const tSnap = await tRef.get();
  if (!tSnap.exists) return json(event, 404, { ok:false, error:'Task not found' });
  const task = tSnap.data();

  const privileged = isPrivileged(role);
  const isAssignee = task.assignedToUid === user.uid;

  if (!privileged && !isAssignee) {
    return json(event, 403, { ok:false, error:'Not allowed' });
  }

  const folderId = await ensureClientFolder(task.clientId);

  const created = await drive().files.create({
    requestBody: { name: fileName, parents: [folderId] },
    media: { mimeType, body: Buffer.from(fileBuffer) },
    fields: 'id, webViewLink'
  });

  const attachment = {
    type: attachmentType,
    fileName,
    mimeType,
    driveFileId: created.data.id,
    driveWebViewLink: created.data.webViewLink,
    uploadedByUid: user.uid,
    uploadedAt: admin.firestore.FieldValue.serverTimestamp()
  };

  await tRef.update({
    attachments: admin.firestore.FieldValue.arrayUnion(attachment),
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  });

  await auditLog({
    taskId,
    action: 'ATTACHMENT_UPLOAD',
    actorUid: user.uid,
    actorEmail: user.email,
    details: { fileName, type: attachmentType }
  });

  return json(event, 200, { ok:true, attachment });
});