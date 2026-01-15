const admin = require('firebase-admin');
const { google } = require('googleapis');

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
  else d.setDate(d.getDate() + i); // AD_HOC fallback
  return d;
}

function cors() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, x-cron-secret',
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS'
  };
}

function json(statusCode, body) {
  return { statusCode, headers: { ...cors(), 'Content-Type': 'application/json' }, body: JSON.stringify(body) };
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

// Build raw MIME email (supports attachments)
function buildRawEmail({ from, to, cc = [], bcc = [], subject, html, attachments = [] }) {
  const boundary = `----cm_${Date.now()}`;
  const headers = [
    `From: ${from}`,
    `To: ${to.join(', ')}`,
    cc.length ? `Cc: ${cc.join(', ')}` : '',
    bcc.length ? `Bcc: ${bcc.join(', ')}` : '',
    `Subject: ${subject}`,
    'MIME-Version: 1.0'
  ].filter(Boolean);

  if (!attachments.length) {
    const msg = [
      ...headers,
      'Content-Type: text/html; charset="UTF-8"',
      '',
      html
    ].join('\n');
    return Buffer.from(msg).toString('base64').replace(/\+/g,'-').replace(/\//g,'_');
  }

  const parts = [];
  parts.push(`--${boundary}`);
  parts.push('Content-Type: text/html; charset="UTF-8"\n');
  parts.push(html + '\n');

  for (const a of attachments) {
    parts.push(`--${boundary}`);
    parts.push(`Content-Type: ${a.mimeType}; name="${a.filename}"`);
    parts.push('Content-Transfer-Encoding: base64');
    parts.push(`Content-Disposition: attachment; filename="${a.filename}"\n`);
    parts.push(a.contentBase64 + '\n');
  }

  parts.push(`--${boundary}--`);

  const msg = [
    ...headers,
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
    '',
    parts.join('\n')
  ].join('\n');

  return Buffer.from(msg).toString('base64').replace(/\+/g,'-').replace(/\//g,'_');
}

async function sendEmail({ to, cc = [], bcc = [], subject, html, attachments = [] }) {
  if (!to?.length) return;
  const g = gmail();
  const raw = buildRawEmail({
    from: process.env.BOT_FROM,
    to, cc, bcc, subject, html, attachments
  });

  await g.users.messages.send({ userId: 'me', requestBody: { raw } });
}

module.exports = {
  admin, db, auth, calendar, drive, gmail,
  ymd, addDays, addInterval,
  json, cors, auditLog, sendEmail
};
