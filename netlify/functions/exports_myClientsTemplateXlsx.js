const { withCors, json, db } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');
const ExcelJS = require('exceljs');

function col(n) {
  // 1 -> A, 2 -> B ...
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

  // load clients (limit for template)
  const snap = await db().collection('clients').limit(200).get();
  const clients = snap.docs.map(d => ({ id: d.id, ...d.data() }));

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Compliance Management';
  wb.created = new Date();

  const ws = wb.addWorksheet('Import');
  const lists = wb.addWorksheet('Lists');

  // Dropdown lists
  const categories = ['GST', 'TDS', 'Income Tax', 'ROC', 'Accounting', 'Audit', 'Other'];
  const recurrences = ['AD_HOC', 'DAILY', 'WEEKLY', 'BIWEEKLY', 'MONTHLY', 'BIMONTHLY', 'QUARTERLY', 'HALF_YEARLY', 'YEARLY'];
  const types = ['FILING', 'REVIEW', 'PAYMENT', 'FOLLOW_UP', 'CLIENT_PENDING']; // you can still type custom

  lists.getCell('A1').value = 'Categories';
  categories.forEach((v, i) => (lists.getCell(`A${i + 2}`).value = v));

  lists.getCell('B1').value = 'Recurrences';
  recurrences.forEach((v, i) => (lists.getCell(`B${i + 2}`).value = v));

  lists.getCell('C1').value = 'Types';
  types.forEach((v, i) => (lists.getCell(`C${i + 2}`).value = v));

  lists.columns.forEach(c => (c.width = 22));

  // Main sheet headers (must match tasks_bulkimportxlsx.js parser)
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

    'ClientTo (emails ; , : separated)',
    'ClientCC (emails ; , : separated)',
    'ClientBCC (emails ; , : separated)',

    'ClientStartSubject',
    'ClientStartBody',

    'SendClientCompletionMail (true/false)',
    'ClientCompletionSubject',
    'ClientCompletionBody',

    'CcAssigneeOnClientStart (true/false)',
    'CcManagerOnClientStart (true/false)'
  ];

  ws.addRow(headers);
  ws.getRow(1).font = { bold: true };
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  const widths = [34, 32, 28, 20, 18, 24, 16, 14, 12, 22, 28, 28, 28, 28, 44, 26, 30, 44, 28, 28];
  ws.columns = headers.map((h, i) => ({ header: h, width: widths[i] || 22 }));

  // sample due date = today + 30 days
  const pad = (n) => String(n).padStart(2, '0');
  const d = new Date();
  d.setDate(d.getDate() + 30);
  const dueDmy = `${pad(d.getDate())}-${pad(d.getMonth() + 1)}-${d.getFullYear()}`;

  const sampleRows = Math.min(50, clients.length);
  for (let i = 0; i < sampleRows; i++) {
    const c = clients[i];
    ws.addRow([
      `Sample Task for ${c.name || 'Client'}`,
      c.name || '',
      c.primaryEmail || '',
      dueDmy,
      'Other',
      'FOLLOW_UP',
      'AD_HOC',
      1,
      7,
      '',

      '', // ClientTo
      '', // ClientCC
      '', // ClientBCC

      'We started working on {{taskTitle}}',
      'Dear {{clientName}},\n\nWe started work on {{taskTitle}}.\nDue: {{dueDate}}.\n\nAdd to calendar: {{addToCalendarUrl}}\n\nRegards,\nYour Firm',

      'true',
      'Completed: {{taskTitle}}',
      'Dear {{clientName}},\n\nWe have completed {{taskTitle}}.\nCompleted at: {{completedAt}}\n\nRegards,\nYour Firm',

      'false',
      'false'
    ]);
  }

  // Apply dropdown validations for a reasonable range
  const maxRows = Math.max(100, sampleRows + 50);

  const catRange = `Lists!$A$2:$A$${categories.length + 1}`;
  const recRange = `Lists!$B$2:$B$${recurrences.length + 1}`;
  const typeRange = `Lists!$C$2:$C$${types.length + 1}`;

  const COL_CATEGORY = headers.indexOf('Category') + 1;
  const COL_TYPE = headers.indexOf('Type (you can type custom)') + 1;
  const COL_REC = headers.indexOf('Recurrence') + 1;

  const COL_SENDCOMP = headers.indexOf('SendClientCompletionMail (true/false)') + 1;
  const COL_CC_ASSIGNEE = headers.indexOf('CcAssigneeOnClientStart (true/false)') + 1;
  const COL_CC_MANAGER = headers.indexOf('CcManagerOnClientStart (true/false)') + 1;

  for (let r = 2; r <= maxRows; r++) {
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

    // Type dropdown but allow custom typing (no error popup)
    ws.getCell(`${col(COL_TYPE)}${r}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [typeRange],
      showErrorMessage: false
    };

    const tfFormula = `"true,false"`;

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
  }

  const buf = await wb.xlsx.writeBuffer();

  return json(event, 200, {
    ok: true,
    fileName: `my_clients_import_template.xlsx`,
    base64: Buffer.from(buf).toString('base64')
  });
});