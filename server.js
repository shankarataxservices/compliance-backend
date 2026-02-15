// server.js
require('dotenv').config();
const express = require('express');

const app = express();

/**
 * Local Netlify Functions runner (Express).
 *
 * This version is updated to match your CURRENT backend file list + endpoints,
 * including the new exports/reports/offline update endpoints and the role migration helper.
 *
 * NOTE about casing:
 * - Some of your filenames are camelCase (users_setManager.js etc.)
 * - Your UI calls lowercase endpoints (/users_setmanager etc.)
 * - Netlify deploy routing is usually forgiving; Express is not.
 * - So we map both casings via FUNCTION_ALIASES.
 */

// Import your existing functions (no rewrite)
const FUNCTIONS = {
  // core
  ping: require('./netlify/functions/ping'),

  // clients
  clients_create: require('./netlify/functions/clients_create'),
  clients_update: require('./netlify/functions/clients_update'),

  // tasks
  tasks_createone: require('./netlify/functions/tasks_createone'),
  tasks_updatestatus: require('./netlify/functions/tasks_updatestatus'),
  tasks_updatetask: require('./netlify/functions/tasks_updatetask'),
  tasks_delete: require('./netlify/functions/tasks_delete'),
  tasks_bulkUpdate: require('./netlify/functions/tasks_bulkUpdate'),
  tasks_uploadattachment: require('./netlify/functions/tasks_uploadattachment'),
  tasks_addcomment: require('./netlify/functions/tasks_addcomment'),

  // imports
  tasks_bulkimportxlsx: require('./netlify/functions/tasks_bulkimportxlsx'),
  tasks_bulkimportcsv: require('./netlify/functions/tasks_bulkimportcsv'), // legacy compat (partner/manager)

  // offline update upload
  tasks_bulkupdate_from_xlsx: require('./netlify/functions/tasks_bulkupdate_from_xlsx'),

  // series tools
  series_rebuild: require('./netlify/functions/series_rebuild'),
  series_reassign: require('./netlify/functions/series_reassign'),

  // settings
  settings_get: require('./netlify/functions/settings_get'),
  settings_update: require('./netlify/functions/settings_update'),
  settings_calendar_get: require('./netlify/functions/settings_calendar_get'),
  settings_calendar_update: require('./netlify/functions/settings_calendar_update'),

  // security setting (associate import password)
  // filename in your backend: settings_workerimportpassword_set.js
  settings_workerimportpassword_set: require('./netlify/functions/settings_workerimportpassword_set'),

    // users
  users_list: require('./netlify/functions/users_list'),
  users_setrole: require('./netlify/functions/users_setrole'),
  
  // NEW: Required for signup initialization
  users_initself: require('./netlify/functions/users_initself'),

  // filenames in your backend folder (camelCase):
  users_setManager: require('./netlify/functions/users_setManager'),
  users_setDisplayName: require('./netlify/functions/users_setDisplayName'),

  // exports (XLSX)
  exports_myClientsTemplateXlsx: require('./netlify/functions/exports_myClientsTemplateXlsx'),
  exports_quickXlsx: require('./netlify/functions/exports_quickXlsx'),
  exports_firmRangeWithHistoryXlsx: require('./netlify/functions/exports_firmRangeWithHistoryXlsx'),
  exports_clientHistoryXlsx: require('./netlify/functions/exports_clientHistoryXlsx'),
  exports_taskHistoryXlsx: require('./netlify/functions/exports_taskHistoryXlsx'),
  exports_tasksUpdateTemplateXlsx: require('./netlify/functions/exports_tasksUpdateTemplateXlsx'),
  exports_tasksExportForUpdateXlsx: require('./netlify/functions/exports_tasksExportForUpdateXlsx'),

  // reports (PDF)
  reports_firmRangePdf: require('./netlify/functions/reports_firmRangePdf'),
  reports_clientHistoryPdf: require('./netlify/functions/reports_clientHistoryPdf'),
  reports_taskHistoryPdf: require('./netlify/functions/reports_taskHistoryPdf'),
  reports_dailyDigestPdf: require('./netlify/functions/reports_dailyDigestPdf'),
  reports_monthlyPdf: require('./netlify/functions/reports_monthlyPdf'),

  // cron jobs
  jobs_daily: require('./netlify/functions/jobs_daily'), // alias/no-op
  jobs_daily5am: require('./netlify/functions/jobs_daily5am'),
  jobs_client0815: require('./netlify/functions/jobs_client0815'),
  jobs_monthlysummary: require('./netlify/functions/jobs_monthlysummary'),

  // migration helper
  roles_migrate_worker_to_associate: require('./netlify/functions/roles_migrate_worker_to_associate'),
};

/**
 * URL path aliases for compatibility:
 * - UI calls lowercase endpoints
 * - some handlers are camelCase keys here
 * - some UI endpoint names differ in casing for workerImportPassword
 */
const FUNCTION_ALIASES = {
  // team endpoints (UI uses lowercase)
  users_setmanager: 'users_setManager',
  users_setdisplayname: 'users_setDisplayName',

  // worker import password endpoint (UI calls /settings_workerImportPassword_set)
  settings_workerImportPassword_set: 'settings_workerimportpassword_set',
  settings_workerimportpassword_set: 'settings_workerimportpassword_set',

  // bulk update endpoint (UI calls /tasks_bulkUpdate)
  tasks_bulkupdate: 'tasks_bulkUpdate',

  // offline update upload: accept alternative casings
  tasks_bulkupdate_from_xlsx: 'tasks_bulkupdate_from_xlsx',
  tasks_bulkupdate_from_Xlsx: 'tasks_bulkupdate_from_xlsx',

  // role migration helper: accept alternative casing
  roles_migrate_worker_to_ASSOCIATE: 'roles_migrate_worker_to_associate',
};

function isJson(req) {
  return (req.headers['content-type'] || '').includes('application/json');
}

app.use((req, res, next) => {
  const origin = req.headers.origin || '*';

  res.setHeader('Access-Control-Allow-Origin', origin);
  res.setHeader('Vary', 'Origin');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, x-cron-secret');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Max-Age', '86400');

  if (req.method === 'OPTIONS') return res.status(204).send('');
  next();
});

app.all('/.netlify/functions/:name', (req, res) => {
  const rawName = String(req.params.name || '');

  const resolvedName =
    FUNCTION_ALIASES[rawName] ||
    FUNCTION_ALIASES[rawName.toLowerCase()] ||
    rawName;

  const fn = FUNCTIONS[resolvedName];
  if (!fn) {
    return res.status(404).json({
      ok: false,
      error: `No such function: ${rawName}`,
      hint: `Resolved name: ${resolvedName}`,
      available: Object.keys(FUNCTIONS).sort(),
    });
  }

  const chunks = [];
  req.on('data', (c) => chunks.push(c));

  req.on('end', async () => {
    const buf = Buffer.concat(chunks);

    const event = {
      httpMethod: req.method,
      headers: req.headers,
      queryStringParameters: req.query,
      path: req.path,
      isBase64Encoded: !isJson(req),
      body: isJson(req) ? buf.toString('utf8') : buf.toString('base64'),
    };

    try {
      const out = await fn.handler(event, {});
      res.status(out.statusCode || 200);
      for (const [k, v] of Object.entries(out.headers || {})) res.setHeader(k, v);
      res.send(out.body || '');
    } catch (e) {
      console.error(e);
      res.status(500).json({ ok: false, error: e.message || String(e) });
    }
  });
});

const PORT = process.env.PORT || 8888;
app.listen(PORT, '0.0.0.0', () => console.log('Backend running on port', PORT));