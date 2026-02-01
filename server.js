require('dotenv').config();
const express = require('express');

const app = express();

// Import your existing functions (no rewrite)
const FUNCTIONS = {
  // core
  ping: require('./netlify/functions/ping'),

  // auth/roles/settings
  settings_get: require('./netlify/functions/settings_get'),
  settings_update: require('./netlify/functions/settings_update'),
  settings_calendar_get: require('./netlify/functions/settings_calendar_get'),
  settings_calendar_update: require('./netlify/functions/settings_calendar_update'),

  users_list: require('./netlify/functions/users_list'),
  users_setrole: require('./netlify/functions/users_setrole'),
  users_setmanager: require('./netlify/functions/users_setmanager'),

  // clients
  clients_create: require('./netlify/functions/clients_create'),

  // tasks
  tasks_createone: require('./netlify/functions/tasks_createone'),
  tasks_updatestatus: require('./netlify/functions/tasks_updatestatus'),
  tasks_updatetask: require('./netlify/functions/tasks_updatetask'),
  tasks_delete: require('./netlify/functions/tasks_delete'),

  // imports
  tasks_bulkimportcsv: require('./netlify/functions/tasks_bulkimportcsv'),
  tasks_bulkimportxlsx: require('./netlify/functions/tasks_bulkimportxlsx'),

  // attachments
  tasks_uploadattachment: require('./netlify/functions/tasks_uploadattachment'),

  // exports
  exports_clientHistoryXlsx: require('./netlify/functions/exports_clientHistoryXlsx'),
  exports_myClientsTemplateXlsx: require('./netlify/functions/exports_myClientsTemplateXlsx'),
  exports_firmRangeWithHistoryXlsx: require('./netlify/functions/exports_firmRangeWithHistoryXlsx'),
  exports_quickXlsx: require('./netlify/functions/exports_quickXlsx'),

  // series tools
  series_rebuild: require('./netlify/functions/series_rebuild'),
  series_reassign: require('./netlify/functions/series_reassign'),

  // cron jobs
  jobs_daily: require('./netlify/functions/jobs_daily'),           // alias/no-op
  jobs_daily5am: require('./netlify/functions/jobs_daily5am'),     // internal digest
  jobs_client0815: require('./netlify/functions/jobs_client0815'), // client mails
  jobs_monthlysummary: require('./netlify/functions/jobs_monthlysummary'),
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

  // IMPORTANT: respond to preflight here
  if (req.method === 'OPTIONS') {
    return res.status(204).send('');
  }
  next();
});

app.all('/.netlify/functions/:name', (req, res) => {
  const fn = FUNCTIONS[req.params.name];
  if (!fn) return res.status(404).json({ ok: false, error: 'No such function' });

  const chunks = [];
  req.on('data', (c) => chunks.push(c));

  req.on('end', async () => {
    const buf = Buffer.concat(chunks);

    // Build a Netlify-style event object
    const event = {
      httpMethod: req.method,
      headers: req.headers,
      queryStringParameters: req.query,
      path: req.path,
      // JSON should be plain string; uploads should be base64
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