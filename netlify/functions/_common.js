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

// ===== Google OAuth =====
function oauthClient() {
  const o = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET
  );
  o.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
  return o;
}

function calendar() { return google.calendar({ version: 'v3', auth: oauthClient() }); }
function gmail() { return google.gmail({ version: 'v1', auth: oauthClient() }); }
function drive() { return google.drive({ version: 'v3', auth: oauthClient() }); }

// ===== CORS =====
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
      return json(event, 500, { ok: false, error: e.message || String(e) });
    }
  };
}

// ===== Timezone-safe date helpers (IST) =====
const IST_TZ = 'Asia/Kolkata';

/**
 * YYYY-MM-DD in a specific timezone (IST).
 * Uses Intl so it is not dependent on server timezone (UTC on Netlify).
 */
function ymdInTZ(date, timeZone) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(date);
  const m = Object.fromEntries(parts.map(p => [p.type, p.value]));
  return `${m.year}-${m.month}-${m.day}`;
}

function ymdIST(date) {
  return ymdInTZ(date, IST_TZ);
}

/**
 * Parse DD-MM-YYYY -> YYYY-MM-DD (strict)
 */
function dmyToYmd(dmyStr) {
  const s = String(dmyStr || '').trim();
  const m = /^(\d{2})-(\d{2})-(\d{4})$/.exec(s);
  if (!m) throw new Error(`Invalid date format "${s}". Expected DD-MM-YYYY`);
  const dd = m[1], mm = m[2], yyyy = m[3];
  const iso = `${yyyy}-${mm}-${dd}`;
  // validate actual date
  const dt = new Date(`${iso}T00:00:00Z`);
  if (Number.isNaN(dt.getTime())) throw new Error(`Invalid date: ${s}`);
  const check = dt.toISOString().slice(0, 10);
  if (check !== iso) throw new Error(`Invalid date: ${s}`);
  return iso;
}

/**
 * YYYY-MM-DD -> DD-MM-YYYY
 */
function ymdToDmy(ymdStr) {
  if (!ymdStr) return '';
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymdStr).trim());
  if (!m) return '';
  return `${m[3]}-${m[2]}-${m[1]}`;
}

function addDays(date, days) {
  const d = new Date(date);
  d.setUTCDate(d.getUTCDate() + days);
  return d;
}

/**
 * Creates a Date that represents 00:00 at IST for a given YYYY-MM-DD.
 * (Internally it's a UTC timestamp, but day math stays consistent for IST)
 */
function dateFromYmdIST(ymdStr) {
  // explicit +05:30 offset
  return new Date(`${ymdStr}T00:00:00+05:30`);
}

// Recurrence step calculator (based on Date math)
function addInterval(baseDate, recurrence, i) {
  const d = new Date(baseDate);
  if (i === 0) return d;

  const r = String(recurrence || 'AD_HOC').toUpperCase();

  if (r === 'DAILY') d.setUTCDate(d.getUTCDate() + i);
  else if (r === 'WEEKLY') d.setUTCDate(d.getUTCDate() + i * 7);
  else if (r === 'BIWEEKLY') d.setUTCDate(d.getUTCDate() + i * 14);
  else if (r === 'MONTHLY') d.setUTCMonth(d.getUTCMonth() + i);
  else if (r === 'BIMONTHLY') d.setUTCMonth(d.getUTCMonth() + i * 2);
  else if (r === 'QUARTERLY') d.setUTCMonth(d.getUTCMonth() + i * 3);
  else if (r === 'HALF_YEARLY') d.setUTCMonth(d.getUTCMonth() + i * 6);
  else if (r === 'YEARLY') d.setUTCFullYear(d.getUTCFullYear() + i);
  else d.setUTCDate(d.getUTCDate() + i);

  return d;
}

// ===== Settings: calendar window =====
async function getCalendarWindow() {
  // Defaults
  const defaults = { startHH: 10, endHH: 12, timeZone: IST_TZ };
  try {
    const snap = await db().collection('settings').doc('calendar').get();
    if (!snap.exists) return defaults;
    const s = snap.data() || {};
    return {
      startHH: Number.isFinite(Number(s.startHH)) ? Number(s.startHH) : defaults.startHH,
      endHH: Number.isFinite(Number(s.endHH)) ? Number(s.endHH) : defaults.endHH,
      timeZone: s.timeZone || defaults.timeZone,
    };
  } catch {
    return defaults;
  }
}

function calTimeRange(ymdStr, startHH = 10, endHH = 12, timeZone = IST_TZ) {
  return {
    start: { dateTime: `${ymdStr}T${String(startHH).padStart(2, '0')}:00:00`, timeZone },
    end: { dateTime: `${ymdStr}T${String(endHH).padStart(2, '0')}:00:00`, timeZone },
  };
}

// ===== Audit =====
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

// ===== Email helpers =====
function asEmailList(x) {
  if (Array.isArray(x)) return x.map(s => String(s).trim()).filter(Boolean);
  if (typeof x === 'string') return x.split(/[;,:]/).map(s => s.trim()).filter(Boolean);
  return [];
}

function uniqEmails(arr) {
  const out = [];
  const seen = new Set();
  for (const e of (arr || [])) {
    const v = String(e || '').trim();
    if (!v) continue;
    const k = v.toLowerCase();
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(v);
  }
  return out;
}

function escapeHtml(s) {
  return String(s || '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function normalizeBodyToHtml(body) {
  const str = String(body || '');
  // If it already contains HTML tags, trust it as HTML
  if (/<[a-z][\s\S]*>/i.test(str)) return str;
  // else treat as plaintext and convert newlines to <br>
  return escapeHtml(str).replace(/\r?\n/g, '<br>');
}

function renderTemplate(str, vars) {
  let out = String(str || '');
  for (const [k, v] of Object.entries(vars || {})) {
    out = out.replaceAll(`{{${k}}}`, v == null ? '' : String(v));
  }
  return out;
}

function buildRawEmail({ from, to, cc = [], bcc = [], subject, html, inReplyTo, references }) {
  const safeFrom = from || process.env.BOT_FROM || process.env.BOT_EMAIL || '';
  if (!safeFrom) throw new Error('BOT_FROM or BOT_EMAIL env var is missing');

  const lines = [
    `From: ${safeFrom}`,
    `To: ${to.join(', ')}`,
    cc.length ? `Cc: ${cc.join(', ')}` : null,
    bcc.length ? `Bcc: ${bcc.join(', ')}` : null,
    `Subject: ${subject}`,
    `Date: ${new Date().toUTCString()}`,
    inReplyTo ? `In-Reply-To: ${inReplyTo}` : null,
    references ? `References: ${references}` : null,
    'MIME-Version: 1.0',
    'Content-Type: text/html; charset="UTF-8"',
    '',
    html || ''
  ];

  const msg = lines.filter(x => x !== null).join('\r\n');

  return Buffer.from(msg)
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

async function getRfcMessageIdHeader(gmailMessageId) {
  if (!gmailMessageId) return null;
  const g = gmail();
  const res = await g.users.messages.get({
    userId: 'me',
    id: gmailMessageId,
    format: 'metadata',
    metadataHeaders: ['Message-ID'],
  });
  const hdrs = res.data?.payload?.headers || [];
  const h = hdrs.find(x => String(x.name || '').toLowerCase() === 'message-id');
  return h?.value || null;
}

/**
 * Sends a NEW email (new thread if no threadId passed).
 * Returns { gmailId, threadId, rfcMessageId }
 */
async function sendEmail({ to, cc = [], bcc = [], subject, html }) {
  const toU = uniqEmails(to);
  const ccU = uniqEmails(cc);
  const bccU = uniqEmails(bcc);
  if (!toU.length) return null;

  const g = gmail();
  const raw = buildRawEmail({
    from: process.env.BOT_FROM,
    to: toU,
    cc: ccU,
    bcc: bccU,
    subject,
    html: normalizeBodyToHtml(html),
  });

  const sent = await g.users.messages.send({
    userId: 'me',
    requestBody: { raw }
  });

  const gmailId = sent.data?.id || null;
  const threadId = sent.data?.threadId || null;
  const rfcMessageId = gmailId ? await getRfcMessageIdHeader(gmailId) : null;

  return { gmailId, threadId, rfcMessageId };
}

/**
 * Sends a reply in an existing Gmail thread (mail trail).
 * Requires threadId. For best threading include inReplyTo/references = RFC Message-ID header value.
 */
async function sendEmailReply({ threadId, inReplyTo, references, to, cc = [], bcc = [], subject, html }) {
  const toU = uniqEmails(to);
  const ccU = uniqEmails(cc);
  const bccU = uniqEmails(bcc);
  if (!toU.length) return null;
  if (!threadId) {
    // fallback: send as new mail
    return sendEmail({ to: toU, cc: ccU, bcc: bccU, subject, html });
  }

  const g = gmail();
  const raw = buildRawEmail({
    from: process.env.BOT_FROM,
    to: toU,
    cc: ccU,
    bcc: bccU,
    subject,
    html: normalizeBodyToHtml(html),
    inReplyTo,
    references: references || inReplyTo || null,
  });

  const sent = await g.users.messages.send({
    userId: 'me',
    requestBody: { raw, threadId }
  });

  return {
    gmailId: sent.data?.id || null,
    threadId: sent.data?.threadId || threadId,
    rfcMessageId: sent.data?.id ? await getRfcMessageIdHeader(sent.data.id) : null
  };
}

// ===== Drive upload helper (safe: uses stream) =====
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
  IST_TZ,
  ymdInTZ, ymdIST, dmyToYmd, ymdToDmy,
  dateFromYmdIST, addDays, addInterval,
  getCalendarWindow, calTimeRange,
  auditLog,
  asEmailList, uniqEmails,
  renderTemplate,
  sendEmail, sendEmailReply,
  driveUpload
};