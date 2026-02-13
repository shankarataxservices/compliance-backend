// netlify/functions/_common.js
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
function getAllowedOrigin(event) {
  const origin = event?.headers?.origin || event?.headers?.Origin || '';
  const allow = String(process.env.CORS_ALLOW_ORIGINS || '').trim();
  if (!allow) return origin || '*';
  const allowed = allow.split(',').map(s => s.trim()).filter(Boolean);
  if (!origin) return allowed[0] || '*';
  return allowed.includes(origin) ? origin : (allowed[0] || '*');
}
function corsHeaders(event) {
  const origin = getAllowedOrigin(event);
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

function ymdInTZ(date, timeZone) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone, year: 'numeric', month: '2-digit', day: '2-digit',
  }).formatToParts(date);
  const m = Object.fromEntries(parts.map(p => [p.type, p.value]));
  return `${m.year}-${m.month}-${m.day}`;
}
function ymdIST(date) { return ymdInTZ(date, IST_TZ); }

function dmyToYmd(dmyStr) {
  const s = String(dmyStr || '').trim();
  const m = /^(\d{2})-(\d{2})-(\d{4})$/.exec(s);
  if (!m) throw new Error(`Invalid date format "${s}". Expected DD-MM-YYYY`);
  const dd = m[1], mm = m[2], yyyy = m[3];
  const iso = `${yyyy}-${mm}-${dd}`;
  const dt = new Date(`${iso}T00:00:00Z`);
  if (Number.isNaN(dt.getTime())) throw new Error(`Invalid date: ${s}`);
  const check = dt.toISOString().slice(0, 10);
  if (check !== iso) throw new Error(`Invalid date: ${s}`);
  return iso;
}
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
function dateFromYmdIST(ymdStr) {
  return new Date(`${ymdStr}T00:00:00+05:30`);
}
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

// NEW: token expansion (supports \\n tokens stored in Firestore)
function expandNewlineTokens(s) {
  let x = String(s ?? '');
  x = x.replaceAll('\\r\\n', '\n');
  x = x.replaceAll('\\n', '\n');
  return x;
}
function normalizeBodyToHtml(body) {
  let str = expandNewlineTokens(body);
  // If it already contains HTML tags, assume rich text
  if (/<[a-z][\s\S]*>/i.test(str)) return str;
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
  const subj = String(subject || '').replace(/\r?\n/g, ' ').trim();
  const lines = [
    `From: ${safeFrom}`,
    `To: ${to.join(', ')}`,
    cc.length ? `Cc: ${cc.join(', ')}` : null,
    bcc.length ? `Bcc: ${bcc.join(', ')}` : null,
    `Subject: ${subj}`,
    `Date: ${new Date().toUTCString()}`,
    inReplyTo ? `In-Reply-To: ${inReplyTo}` : null,
    references ? `References: ${references}` : null,
    'MIME-Version: 1.0',
    'Content-Type: text/html; charset="UTF-8"',
    '',
    html || ''
  ];
  const msg = lines.filter(x => x !== null).join('\r\n');
  return Buffer.from(msg).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

async function getMessageHeaders(gmailMessageId, headerNames = ['Message-ID', 'References']) {
  if (!gmailMessageId) return {};
  const g = gmail();
  const res = await g.users.messages.get({
    userId: 'me',
    id: gmailMessageId,
    format: 'metadata',
    metadataHeaders: headerNames,
  });
  const hdrs = res.data?.payload?.headers || [];
  const out = {};
  for (const h of hdrs) {
    const name = String(h.name || '').toLowerCase();
    out[name] = h.value || '';
  }
  return out;
}

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

  const sent = await g.users.messages.send({ userId: 'me', requestBody: { raw } });
  const gmailId = sent.data?.id || null;
  const threadId = sent.data?.threadId || null;

  let rfcMessageId = null;
  let references = null;
  try {
    if (gmailId) {
      const hdr = await getMessageHeaders(gmailId, ['Message-ID', 'References']);
      rfcMessageId = hdr['message-id'] || null;
      references = hdr['references'] || null;
    }
  } catch (e) {
    console.warn('Could not fetch headers. Continuing.', e?.message || e);
  }

  return { gmailId, threadId, rfcMessageId, references };
}

async function sendEmailReply({ threadId, inReplyTo, references, to, cc = [], bcc = [], subject, html }) {
  const toU = uniqEmails(to);
  const ccU = uniqEmails(cc);
  const bccU = uniqEmails(bcc);
  if (!toU.length) return null;

  if (!threadId) {
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

  const sent = await g.users.messages.send({ userId: 'me', requestBody: { raw, threadId } });
  const gmailId = sent.data?.id || null;
  const outThreadId = sent.data?.threadId || threadId;

  let rfcMessageId = null;
  let outReferences = null;
  try {
    if (gmailId) {
      const hdr = await getMessageHeaders(gmailId, ['Message-ID', 'References']);
      rfcMessageId = hdr['message-id'] || null;
      outReferences = hdr['references'] || null;
    }
  } catch (e) {
    console.warn('Could not fetch headers for reply. Continuing.', e?.message || e);
  }

  return { gmailId, threadId: outThreadId, rfcMessageId, references: outReferences || references || null };
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

// Helper: Create single START event
async function createStartCalendarEvent({
  title,
  clientId,
  clientName,
  startDateYmd,
  dueDateYmd,
  window,
  calendarDescription
}) {
  const cal = calendar();
  const range = calTimeRange(startDateYmd, window.startHH, window.endHH, window.timeZone);

  const base =
    `Client: ${String(clientName || '').trim() || String(clientId || '').trim()}\n` +
    `Start: ${ymdToDmy(startDateYmd)}\n` +
    `Due: ${ymdToDmy(dueDateYmd)}\n`;

  const extra = String(calendarDescription || '').trim();
  const description = extra ? `${base}\n${extra}` : base;

  const res = await cal.events.insert({
    calendarId: 'primary',
    sendUpdates: 'none',
    requestBody: {
      summary: `START: ${title}`,
      description,
      ...range
    }
  });
  return { calendarEventId: res.data.id, calendarHtmlLink: res.data.htmlLink || null };
}

function buildGoogleCalendarTemplateUrl({ title, startYmd, startHH, endHH, timeZone, details }) {
  const ymdToCompact = (s) => String(s).replaceAll('-', '');
  const hh2 = (h) => String(h).padStart(2, '0');
  const start = `${ymdToCompact(startYmd)}T${hh2(startHH)}0000`;
  const end = `${ymdToCompact(startYmd)}T${hh2(endHH)}0000`;
  const params = new URLSearchParams({
    action: 'TEMPLATE',
    text: title || 'Compliance Task',
    dates: `${start}/${end}`,
    ctz: timeZone || 'Asia/Kolkata',
    details: details || ''
  });
  return `https://calendar.google.com/calendar/render?${params.toString()}`;
}

// ===== Recipient logic (start + completion) =====
async function getManagerEmailForAssignee(assignedToUid) {
  if (!assignedToUid) return null;
  try {
    const snap = await db().collection('users').doc(assignedToUid).get();
    if (!snap.exists) return null;
    const u = snap.data() || {};
    return u.managerEmail || null;
  } catch {
    return null;
  }
}

/**
 * Resolve recipients for START email
 * - If sendClientStartMail === false: we DO NOT use client primary or task.clientToEmails
 * - But we still allow internal CC/BCC (client master + task overrides)
 * - Gmail requires To, so we promote first CC/BCC into To when needed
 */
async function resolveStartRecipients({ client, task }) {
  const sendClientStartMail = (task.sendClientStartMail !== false);

  const taskTo = Array.isArray(task.clientToEmails) ? task.clientToEmails : [];
  const taskCc = Array.isArray(task.clientCcEmails) ? task.clientCcEmails : [];
  const taskBcc = Array.isArray(task.clientBccEmails) ? task.clientBccEmails : [];

  const baseClientTo = (client.primaryEmail ? [client.primaryEmail] : []);
  const baseClientCc = Array.isArray(client.ccEmails) ? client.ccEmails : [];
  const baseClientBcc = Array.isArray(client.bccEmails) ? client.bccEmails : [];

  let to = [];
  if (sendClientStartMail) {
    to = taskTo.length ? taskTo : baseClientTo;
  } else {
    to = []; // explicitly none
  }

  const cc = uniqEmails([...(baseClientCc || []), ...(taskCc || [])]);
  const bcc = uniqEmails([...(baseClientBcc || []), ...(taskBcc || [])]);

  // Optional CC assignee/manager on start
  if (task.ccAssigneeOnClientStart === true && task.assignedToEmail) cc.push(task.assignedToEmail);
  if (task.ccManagerOnClientStart === true && task.assignedToUid) {
    const mgr = await getManagerEmailForAssignee(task.assignedToUid);
    if (mgr) cc.push(mgr);
  }

  to = uniqEmails(to);
  const ccU = uniqEmails(cc);
  const bccU = uniqEmails(bcc);

  // Promote CC/BCC into To if To is empty
  if (!to.length) {
    if (ccU.length) to = [ccU[0]];
    else if (bccU.length) to = [bccU[0]];
  }

  return { to: uniqEmails(to), cc: ccU, bcc: bccU, sendClientStartMail };
}

/**
 * Resolve recipients for COMPLETION email (reply-all behaviour + overrides)
 * Overrides are completionToEmails/completionCcEmails/completionBccEmails.
 * If completionToEmails empty -> fallback to start logic's To (client primary / task start To)
 * Always include internal trail toggles: ccAssigneeOnCompletion / ccManagerOnCompletion.
 */
async function resolveCompletionRecipients({ client, task }) {
  const compTo = Array.isArray(task.completionToEmails) ? task.completionToEmails : [];
  const compCc = Array.isArray(task.completionCcEmails) ? task.completionCcEmails : [];
  const compBcc = Array.isArray(task.completionBccEmails) ? task.completionBccEmails : [];

  let { to, cc, bcc } = await resolveStartRecipients({ client, task });
  // For completion we should not respect sendClientStartMail; completion is separate flag
  // We'll fallback to start-based recipients only when completion override missing.

  const finalTo = compTo.length ? uniqEmails(compTo) : uniqEmails(to);
  const finalCc = uniqEmails([...(cc || []), ...(compCc || [])]);
  const finalBcc = uniqEmails([...(bcc || []), ...(compBcc || [])]);

  if (task.ccAssigneeOnCompletion === true && task.assignedToEmail) finalCc.push(task.assignedToEmail);
  if (task.ccManagerOnCompletion === true && task.assignedToUid) {
    const mgr = await getManagerEmailForAssignee(task.assignedToUid);
    if (mgr) finalCc.push(mgr);
  }

  const toU = uniqEmails(finalTo);
  const ccU = uniqEmails(finalCc);
  const bccU = uniqEmails(finalBcc);

  // Promote CC/BCC into To if To is empty (Gmail requirement)
  let toOut = toU;
  if (!toOut.length) {
    if (ccU.length) toOut = [ccU[0]];
    else if (bccU.length) toOut = [bccU[0]];
  }

  return { to: uniqEmails(toOut), cc: ccU, bcc: bccU };
}

// ===== Start mail send helper (shared by create/import) =====
async function trySendStartMailImmediately({ task, client, window }) {
  if (!task.clientStartSubject && !task.clientStartBody) return null;

  const recipients = await resolveStartRecipients({ client, task });
  if (!recipients.to.length) return null;

  const addToCalendarUrl = buildGoogleCalendarTemplateUrl({
    title: `START: ${task.title || 'Task'}`,
    startYmd: task.startDateYmd,
    startHH: window.startHH,
    endHH: window.endHH,
    timeZone: window.timeZone,
    details:
      `Client: ${client.name || ''}\n` +
      `Task: ${task.title || ''}\n` +
      `Start: ${ymdToDmy(task.startDateYmd)}\n` +
      `Due: ${ymdToDmy(task.dueDateYmd)}\n`
  });

  const vars = {
    clientName: client.name || '',
    taskTitle: task.title || '',
    startDate: ymdToDmy(task.startDateYmd),
    dueDate: ymdToDmy(task.dueDateYmd),
    addToCalendarUrl
  };

  const subject = renderTemplate(task.clientStartSubject || `We started {{taskTitle}}`, vars);
  const baseBody = renderTemplate(
    task.clientStartBody || `Dear {{clientName}},\n\nWe started work on {{taskTitle}}.\nDue: {{dueDate}}\n\nRegards,\nCompliance Team`,
    vars
  );
  const appended = `${baseBody}\n\n---\nAdd to your Google Calendar:\n${addToCalendarUrl}`;

  const mailRes = await sendEmail({
    to: recipients.to,
    cc: recipients.cc,
    bcc: recipients.bcc,
    subject,
    html: appended
  });

  return {
    clientStartMailSent: true,
    clientStartMailSentAt: admin.firestore.FieldValue.serverTimestamp(),
    clientStartGmailThreadId: mailRes?.threadId || null,
    clientStartGmailId: mailRes?.gmailId || null,
    clientStartRfcMessageId: mailRes?.rfcMessageId || null,
    clientStartReferences: mailRes?.references || null
  };
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
  driveUpload,
  createStartCalendarEvent,
  trySendStartMailImmediately,
  expandNewlineTokens,
  // NEW exports
  resolveStartRecipients,
  resolveCompletionRecipients,
  getManagerEmailForAssignee,
  buildGoogleCalendarTemplateUrl
};