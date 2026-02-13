// netlify/functions/exports_tasksUpdateTemplateXlsx.js
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
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;

  const { user } = authRes;
  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Compliance Management';
  wb.created = new Date();

  const ws = wb.addWorksheet('Update');
  const lists = wb.addWorksheet('Lists');

  const statuses = ['PENDING','IN_PROGRESS','CLIENT_PENDING','APPROVAL_PENDING','COMPLETED'];
  const categories = ['GST', 'TDS', 'INCOME_TAX', 'ROC', 'ACCOUNTING', 'AUDIT', 'OTHER'];
  const priorities = ['HIGH','MEDIUM','LOW'];
  const tf = ['true','false'];

  lists.getCell('A1').value = 'Statuses';
  statuses.forEach((v, i) => (lists.getCell(`A${i + 2}`).value = v));
  lists.getCell('B1').value = 'Categories';
  categories.forEach((v, i) => (lists.getCell(`B${i + 2}`).value = v));
  lists.getCell('C1').value = 'Priorities';
  priorities.forEach((v, i) => (lists.getCell(`C${i + 2}`).value = v));
  lists.getCell('D1').value = 'TrueFalse';
  tf.forEach((v, i) => (lists.getCell(`D${i + 2}`).value = v));
  lists.columns.forEach(c => (c.width = 22));

  const headers = [
    'TaskId', // REQUIRED for update
    'Title',
    'Category',
    'Type',
    'Priority',
    'DueDate (DD-MM-YYYY)',
    'TriggerDays',
    'AssignedToEmail',
    'Status',
    'StatusNote',
    'DelayReason',
    'DelayNotes',
    'SnoozedUntil (DD-MM-YYYY)',
    'SendStartMail (true/false)',
    'ClientTo (emails ; , : separated)',
    'ClientCC (emails ; , : separated)',
    'ClientBCC (emails ; , : separated)',
    'CcAssigneeOnClientStart (true/false)',
    'CcManagerOnClientStart (true/false)',
    'ClientStartSubject',
    'ClientStartBody',
    'SendClientCompletionMail (true/false)',
    'CompletionTo (emails ; , : separated)',
    'CompletionCC (emails ; , : separated)',
    'CompletionBCC (emails ; , : separated)',
    'CcAssigneeOnCompletion (true/false)',
    'CcManagerOnCompletion (true/false)',
    'ClientCompletionSubject',
    'ClientCompletionBody',
  ];

  ws.addRow(headers);
  ws.getRow(1).font = { bold: true };
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  const widths = [
    26, 36, 18, 20, 12, 20, 12, 26, 18, 36, 16, 32, 22,
    18, 28, 28, 28, 22, 22, 30, 44,
    26, 28, 28, 28, 26, 26, 30, 44
  ];
  ws.columns = headers.map((h, i) => ({ header: h, width: widths[i] || 22 }));

  // Add a sample row (TaskId empty intentionally)
  ws.addRow([
    '', 'Example: GSTR-3B Filing', 'GST', 'FILING', 'MEDIUM',
    '20-02-2026', '15', 'worker@firm.com',
    'PENDING', '', '', '', '',
    'true', '', '', '',
    'false', 'false',
    'We started {{taskTitle}}', 'Dear {{clientName}},\\n\\nWe started work on {{taskTitle}}.',
    'true', '', '', '',
    'false', 'false',
    'Completed: {{taskTitle}}', 'Dear {{clientName}},\\n\\nCompleted: {{taskTitle}} at {{completedAt}}.'
  ]);

  // Data validations
  const maxRows = 400;
  const statusRange = `Lists!$A$2:$A$${statuses.length + 1}`;
  const catRange = `Lists!$B$2:$B$${categories.length + 1}`;
  const priRange = `Lists!$C$2:$C$${priorities.length + 1}`;
  const tfRange = `Lists!$D$2:$D$${tf.length + 1}`;

  const COL_CATEGORY = headers.indexOf('Category') + 1;
  const COL_PRIORITY = headers.indexOf('Priority') + 1;
  const COL_STATUS = headers.indexOf('Status') + 1;
  const COL_SENDSTART = headers.indexOf('SendStartMail (true/false)') + 1;
  const COL_SENDCOMP = headers.indexOf('SendClientCompletionMail (true/false)') + 1;
  const COL_CC_ASSIGNEE = headers.indexOf('CcAssigneeOnClientStart (true/false)') + 1;
  const COL_CC_MANAGER = headers.indexOf('CcManagerOnClientStart (true/false)') + 1;
  const COL_CC_ASSIGNEE_COMP = headers.indexOf('CcAssigneeOnCompletion (true/false)') + 1;
  const COL_CC_MANAGER_COMP = headers.indexOf('CcManagerOnCompletion (true/false)') + 1;

  for (let r = 2; r <= maxRows; r++) {
    ws.getCell(`${col(COL_STATUS)}${r}`).dataValidation = { type:'list', allowBlank:true, formulae:[statusRange], showErrorMessage:false };
    ws.getCell(`${col(COL_CATEGORY)}${r}`).dataValidation = { type:'list', allowBlank:true, formulae:[catRange], showErrorMessage:false };
    ws.getCell(`${col(COL_PRIORITY)}${r}`).dataValidation = { type:'list', allowBlank:true, formulae:[priRange], showErrorMessage:false };

    ws.getCell(`${col(COL_SENDSTART)}${r}`).dataValidation = { type:'list', allowBlank:true, formulae:[tfRange], showErrorMessage:false };
    ws.getCell(`${col(COL_SENDCOMP)}${r}`).dataValidation = { type:'list', allowBlank:true, formulae:[tfRange], showErrorMessage:false };

    ws.getCell(`${col(COL_CC_ASSIGNEE)}${r}`).dataValidation = { type:'list', allowBlank:true, formulae:[tfRange], showErrorMessage:false };
    ws.getCell(`${col(COL_CC_MANAGER)}${r}`).dataValidation = { type:'list', allowBlank:true, formulae:[tfRange], showErrorMessage:false };
    ws.getCell(`${col(COL_CC_ASSIGNEE_COMP)}${r}`).dataValidation = { type:'list', allowBlank:true, formulae:[tfRange], showErrorMessage:false };
    ws.getCell(`${col(COL_CC_MANAGER_COMP)}${r}`).dataValidation = { type:'list', allowBlank:true, formulae:[tfRange], showErrorMessage:false };
  }

  const buf = await wb.xlsx.writeBuffer();
  return json(event, 200, {
    ok: true,
    fileName: 'Tasks_Update_Template.xlsx',
    base64: Buffer.from(buf).toString('base64')
  });
});