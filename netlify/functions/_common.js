const admin = require('firebase-admin');
const { google } = require('googleapis');
const { Readable } = require('stream');

let inited = false;

function init() {
  if (inited) return;
  const sa = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
  admin.initializeApp({ credential: admin.credential.cert(sa) });
  inited = true;
}

function db() { init(); return admin.firestore(); }
function auth() { init(); return admin.auth(); }

function oauthClient() {
  const o = new google.auth.OAuth2(process.env.GOOGLE_CLIENT_ID, process.env.GOOGLE_CLIENT_SECRET);
  o.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
  return o;
}

function calendar() { return google.calendar({ version: 'v3', auth: oauthClient() }); }
function gmail() { return google.gmail({ version: 'v1', auth: oauthClient() }); }
function drive() { return google.drive({ version: 'v3', auth: oauthClient() }); }

/** CORS headers ALWAYS */
function corsHeaders(event) {
  const origin = event?.headers?.origin || event?.headers?.Origin || '*';
  return {
    'Access-Control-Allow-Origin': origin,
    'Vary': 'Origin',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, x-cron-secret',
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
    'Access-Control-Max-Age': '86400',
  };
}

function json(event, statusCode, body) {
  return {
    statusCode,
    headers: { ...corsHeaders(event), 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  };
}

/** Wrapper: OPTIONS + try/catch + always CORS */
function withCors(handler) {
  return async (event, context) => {
    if (event.httpMethod === 'OPTIONS') {
      return { statusCode: 204, headers: corsHeaders(event), body: '' };
    }
    try {
      const res = await handler(event, context);
      res.headers = { ...corsHeaders(event), ...(res.headers || {}) };
      return res;
    } catch (e) {
      console.error('Function crashed:', e);
      return json(event, 500, { ok:false, error: e.message || String(e) });
    }
  };
}

function ymd(d) {
  const x = new Date(d);
  x.setHours(0,0,0,0);
  const yyyy = x.getFullYear();
  const mm = String(x.getMonth()+1).padStart(2,'0');
  const dd = String(x.getDate()).padStart(2,'0');
  return `${yyyy}-${mm}-${dd}`;
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function addInterval(baseDate, recurrence, i) {
  const d = new Date(baseDate);
  if (i === 0) return d;

  if (recurrence === 'WEEKLY') d.setDate(d.getDate() + i * 7);
  else if (recurrence === 'MONTHLY') d.setMonth(d.getMonth() + i);
  else if (recurrence === 'QUARTERLY') d.setMonth(d.getMonth() + i * 3);
  else if (recurrence === 'HALF_YEARLY') d.setMonth(d.getMonth() + i * 6);
  else if (recurrence === 'YEARLY') d.setFullYear(d.getFullYear() + i);
  else d.setDate(d.getDate() + i);

  return d;
}

async function auditLog({ taskId, action, actorUid, actorEmail, details }) {
  await db().collection('auditLogs').add({
    taskId: taskId || null,
    action,
    actorUid: actorUid || null,
    actorEmail: actorEmail || null,
    timestamp: admin.firestore.FieldValue.serverTimestamp(),
    details: details || {}
  });
}

function buildRawEmail({ from, to, cc = [], bcc = [], subject, html }) {
  const safeFrom = from || process.env.BOT_FROM || process.env.BOT_EMAIL || '';
  if (!safeFrom) throw new Error('BOT_FROM or BOT_EMAIL env var is missing');

  const lines = [
    `From: ${safeFrom}`,
    `To: ${to.join(', ')}`,
    cc.length ? `Cc: ${cc.join(', ')}` : null,
    bcc.length ? `Bcc: ${bcc.join(', ')}` : null,
    `Subject: ${subject}`,
    `Date: ${new Date().toUTCString()}`,
    'MIME-Version: 1.0',
    'Content-Type: text/html; charset="UTF-8"',
    '',
    html || ''
  ];

  const msg = lines.filter(x => x !== null).join('\r\n');

  // Gmail wants base64url (no + /) and usually without trailing '=' padding
  return Buffer.from(msg)
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

async function sendEmail({ to, cc = [], bcc = [], subject, html }) {
  if (!to?.length) return;

  const g = gmail();
  const raw = buildRawEmail({
    from: process.env.BOT_FROM,
    to, cc, bcc, subject, html,
  });

  await g.users.messages.send({
    userId: 'me',
    requestBody: { raw }
  });
}
async function sendEmail({ to, cc = [], bcc = [], subject, html }) {
  if (!to?.length) return;
  const g = gmail();
  const raw = buildRawEmail({
    from: process.env.BOT_FROM,
    to, cc, bcc, subject, html,
  });
  await g.users.messages.send({ userId: 'me', requestBody: { raw } });
}

async function driveUpload({ folderId, filename, mimeType, buffer }) {
  const d = drive();
  const res = await d.files.create({
    requestBody: { name: filename, parents: [folderId] },
    media: { mimeType, body: Readable.from(buffer) },
    fields: 'id, webViewLink, size, mimeType, name'
  });
  return res.data;
}

module.exports = {
  admin, db, auth, calendar, gmail, drive,
  withCors, json,
  ymd, addDays, addInterval,
  auditLog, sendEmail,
  driveUpload
};