// netlify/functions/exports_myClientsTemplateXlsx.js
const { withCors, json, db } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');
const ExcelJS = require('exceljs');

function col(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok: false, error: 'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;

  const { user } = authRes;
  const p = requirePartner(event, user);
  if (p.error) return p.error;

  // load clients
  const snap = await db().collection('clients').limit(400).get();
  const clients = snap.docs.map(d => ({ id: d.id, ...d.data() }));

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Compliance Management';
  wb.created = new Date();

  const ws = wb.addWorksheet('Import');
  const lists = wb.addWorksheet('Lists');

  // Dropdown lists
  const categories = ['GST', 'TDS', 'Income Tax', 'ROC', 'Accounting', 'Audit', 'Other'];
  const recurrences = ['AD_HOC', 'DAILY', 'WEEKLY', 'BIWEEKLY', 'MONTHLY', 'BIMONTHLY', 'QUARTERLY', 'HALF_YEARLY', 'YEARLY'];
  const types = ['FILING', 'REVIEW', 'PAYMENT', 'FOLLOW_UP', 'CLIENT_PENDING'];
  const priorities = ['HIGH', 'MEDIUM', 'LOW'];

  lists.getCell('A1').value = 'Categories';
  categories.forEach((v, i) => (lists.getCell(`A${i + 2}`).value = v));

  lists.getCell('B1').value = 'Recurrences';
  recurrences.forEach((v, i) => (lists.getCell(`B${i + 2}`).value = v));

  lists.getCell('C1').value = 'Types';
  types.forEach((v, i) => (lists.getCell(`C${i + 2}`).value = v));

  lists.getCell('D1').value = 'Priorities';
  priorities.forEach((v, i) => (lists.getCell(`D${i + 2}`).value = v));

  lists.getCell('E1').value = 'Clients';
  const clientNames = clients
    .map(c => String(c.name || '').trim())
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
  clientNames.forEach((v, i) => (lists.getCell(`E${i + 2}`).value = v));

  lists.columns.forEach(c => (c.width = 24));
  lists.getColumn(5).width = 42;

  // Main sheet headers
  const headers = [
    'Title',
    'Client',
    'ClientEmail',
    'DueDate (DD-MM-YYYY)',
    'Category',
    'Type (you can type custom)',
    'Recurrence',
    'GenerateCount',
    'TriggerDays',
    'AssignedToEmail',
    'Priority',

    'SendStartMail (true/false)',
    'ClientTo (emails ; , : separated)',
    'ClientCC (emails ; , : separated)',
    'ClientBCC (emails ; , : separated)',
    'ClientStartSubject',
    'ClientStartBody',

    'SendClientCompletionMail (true/false)',
    'CompletionTo (emails ; , : separated)',
    'CompletionCC (emails ; , : separated)',
    'CompletionBCC (emails ; , : separated)',
    'CcAssigneeOnClientStart (true/false)',
    'CcManagerOnClientStart (true/false)',
    'CcAssigneeOnCompletion (true/false)',
    'CcManagerOnCompletion (true/false)',
    'ClientCompletionSubject',
    'ClientCompletionBody',
    
    // NEW: Google Calendar event description (optional)
    'GoogleCalendarDescription',
    ];
  ws.addRow(headers);
  ws.getRow(1).font = { bold: true };
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  const widths = [
 34, 32, 28, 20, 18, 24, 16, 14, 12, 22, 12,
 18, 28, 28, 28, 30, 44,
 26, 28, 28, 28, 28, 28, 28, 28, 30, 44,
 46, // GoogleCalendarDescription
];
  ws.columns = headers.map((h, i) => ({ header: h, width: widths[i] || 22 }));

  // sample due date = today + 30 days
  const pad = (n) => String(n).padStart(2, '0');
  const d = new Date();
  d.setDate(d.getDate() + 30);
  const dueDmy = `${pad(d.getDate())}-${pad(d.getMonth() + 1)}-${d.getFullYear()}`;

  // sample rows
  const sampleRows = Math.min(30, clientNames.length);
  for (let i = 0; i < sampleRows; i++) {
    const cname = clientNames[i] || '';
    ws.addRow([
      `Sample Task for ${cname || 'Client'}`,
      cname,
      '',
      dueDmy,
      'Other',
      'FOLLOW_UP',
      'AD_HOC',
      1,
      7,
      '',
      'MEDIUM',

      'true',
      '',
      '',
      '',
      'We started working on {{taskTitle}}',
      `Dear {{clientName}},\n\nWe started work on {{taskTitle}}.\nDue: {{dueDate}}.\n\nAdd to calendar: {{addToCalendarUrl}}\n\nRegards,\nYour Firm`,

      'true',
      '',
      '',
      '',
      'false',
      'false',
      'false',
      'false',
      'Completed: {{taskTitle}}',
      `Dear {{clientName}},\n\nWe have completed {{taskTitle}}.\nCompleted at: {{completedAt}}\n\nRegards,\nYour Firm`,
      
      '', // GoogleCalendarDescription (optional)
      ]);  }

  // Validations
  const maxRows = Math.max(150, sampleRows + 80);
  const catRange = `Lists!$A$2:$A$${categories.length + 1}`;
  const recRange = `Lists!$B$2:$B$${recurrences.length + 1}`;
  const typeRange = `Lists!$C$2:$C$${types.length + 1}`;
  const priRange = `Lists!$D$2:$D$${priorities.length + 1}`;
  const clientsRange = `Lists!$E$2:$E$${clientNames.length + 1}`;

  const COL_CLIENT = headers.indexOf('Client') + 1;
  const COL_CATEGORY = headers.indexOf('Category') + 1;
  const COL_TYPE = headers.indexOf('Type (you can type custom)') + 1;
  const COL_REC = headers.indexOf('Recurrence') + 1;
  const COL_PRIORITY = headers.indexOf('Priority') + 1;

  const COL_SENDSTART = headers.indexOf('SendStartMail (true/false)') + 1;
  const COL_SENDCOMP = headers.indexOf('SendClientCompletionMail (true/false)') + 1;

  const COL_CC_ASSIGNEE = headers.indexOf('CcAssigneeOnClientStart (true/false)') + 1;
  const COL_CC_MANAGER = headers.indexOf('CcManagerOnClientStart (true/false)') + 1;
  const COL_CC_ASSIGNEE_COMP = headers.indexOf('CcAssigneeOnCompletion (true/false)') + 1;
  const COL_CC_MANAGER_COMP = headers.indexOf('CcManagerOnCompletion (true/false)') + 1;

  const tfFormula = `"true,false"`;

  for (let r = 2; r <= maxRows; r++) {
    ws.getCell(`${col(COL_CLIENT)}${r}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [clientsRange],
      showErrorMessage: false
    };
    ws.getCell(`${col(COL_CATEGORY)}${r}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [catRange],
      showErrorMessage: true
    };
    ws.getCell(`${col(COL_REC)}${r}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [recRange],
      showErrorMessage: true
    };
    ws.getCell(`${col(COL_TYPE)}${r}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [typeRange],
      showErrorMessage: false
    };
    ws.getCell(`${col(COL_PRIORITY)}${r}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [priRange],
      showErrorMessage: true
    };

    ws.getCell(`${col(COL_SENDSTART)}${r}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [tfFormula],
      showErrorMessage: true
    };
    ws.getCell(`${col(COL_SENDCOMP)}${r}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [tfFormula],
      showErrorMessage: true
    };
    ws.getCell(`${col(COL_CC_ASSIGNEE)}${r}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [tfFormula],
      showErrorMessage: true
    };
    ws.getCell(`${col(COL_CC_MANAGER)}${r}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [tfFormula],
      showErrorMessage: true
    };
    ws.getCell(`${col(COL_CC_ASSIGNEE_COMP)}${r}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [tfFormula],
      showErrorMessage: true
    };
    ws.getCell(`${col(COL_CC_MANAGER_COMP)}${r}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [tfFormula],
      showErrorMessage: true
    };
  }

  const buf = await wb.xlsx.writeBuffer();
  return json(event, 200, {
    ok: true,
    fileName: `Client_Tasks_Import_Template.xlsx`,
    base64: Buffer.from(buf).toString('base64')
  });
});