const Busboy = require('busboy');
const { json, db, admin, drive, auditLog } = require('./_common');
const { requireUser } = require('./_auth');

async function ensureRootFolder() {
  const d = drive();
  const name = process.env.DRIVE_ROOT_FOLDER_NAME || 'ComplianceManagement';

  // search existing folder by name
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
  const client = cSnap.data();
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

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: require('./_common').cors() };
  if (event.httpMethod !== 'POST') return json(405, { error: 'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const busboy = Busboy({ headers: event.headers });

  let taskId = null;
  let attachmentType = 'OTHER';
  let fileBuffer = Buffer.alloc(0);
  let fileName = 'file';
  let mimeType = 'application/octet-stream';

  busboy.on('field', (name, val) => {
    if (name === 'taskId') taskId = val;
    if (name === 'type') attachmentType = val.toUpperCase();
  });

  busboy.on('file', (name, file, info) => {
    fileName = info.filename;
    mimeType = info.mimeType;
    file.on('data', (d) => { fileBuffer = Buffer.concat([fileBuffer, d]); });
  });

  const done = new Promise((resolve, reject) => {
    busboy.on('finish', resolve);
    busboy.on('error', reject);
  });

  // Netlify sends body as base64 sometimes
  const body = event.isBase64Encoded ? Buffer.from(event.body, 'base64') : event.body;
  busboy.end(body);
  await done;

  if (!taskId) return json(400, { error: 'taskId required' });

  const tRef = db().collection('tasks').doc(taskId);
  const tSnap = await tRef.get();
  if (!tSnap.exists) return json(404, { error: 'Task not found' });

  const task = tSnap.data();
  const isPartner = user.role === 'PARTNER';
  const isAssignee = task.assignedToUid === user.uid;
  if (!isPartner && !isAssignee) return json(403, { error: 'Not allowed' });

  const folderId = await ensureClientFolder(task.clientId);

  const d = drive();
  const created = await d.files.create({
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

  await auditLog({ taskId, action: 'ATTACHMENT_UPLOAD', actorUid: user.uid, actorEmail: user.email, details: { fileName, type: attachmentType } });

  return json(200, { ok: true, attachment });
};
