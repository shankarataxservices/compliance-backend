const { withCors, json, db, ymdIST, addDays, dateFromYmdIST, ymdToDmy, IST_TZ } = require('./_common');
const { requireUser, requirePartner } = require('./_auth');
const ExcelJS = require('exceljs');

function tsToIstString(ts) {
  if (!ts || !ts.toDate) return '';
  return ts.toDate().toLocaleString('en-IN', { timeZone: IST_TZ });
}

async function loadClientsMap(clientIds) {
  const ids = [...new Set((clientIds || []).filter(Boolean))];
  const map = new Map();
  if (!ids.length) return map;
  const refs = ids.map(id => db().collection('clients').doc(id));
  const snaps = await db().getAll(...refs);
  snaps.forEach(s => { if (s.exists) map.set(s.id, s.data()); });
  return map;
}

exports.handler = withCors(async (event) => {
  if (event.httpMethod !== 'POST') return json(event, 405, { ok:false, error:'POST only' });

  const authRes = await requireUser(event);
  if (authRes.error) return authRes.error;
  const { user } = authRes;

  const p = requirePartner(event, user);
  if (p.error) return p.error;

  const body = JSON.parse(event.body || '{}');
  const mode = String(body.mode || '').toUpperCase().trim();
  if (!mode) return json(event, 400, { ok:false, error:'mode required' });

  const todayYmd = ymdIST(new Date());
  let fromYmd = todayYmd;
  let toYmd = todayYmd;

  if (mode === 'NEXT_7') toYmd = ymdIST(addDays(dateFromYmdIST(todayYmd), 7));
  else if (mode === 'NEXT_15') toYmd = ymdIST(addDays(dateFromYmdIST(todayYmd), 15));
  else if (mode === 'NEXT_30') toYmd = ymdIST(addDays(dateFromYmdIST(todayYmd), 30));
  else if (mode === 'OVERDUE') { /* handled below */ }
  else if (mode === 'APPROVAL_PENDING') { /* handled below */ }
  else return json(event, 400, { ok:false, error:'Invalid mode' });

  let q = db().collection('tasks');

  if (mode === 'OVERDUE') {
    q = q.where('status', 'in', ['PENDING','IN_PROGRESS','CLIENT_PENDING','APPROVAL_PENDING'])
         .where('dueDateYmd', '<', todayYmd);
  } else if (mode === 'APPROVAL_PENDING') {
    q = q.where('status', '==', 'APPROVAL_PENDING');
  } else {
    // upcoming window (exclude completed for operational list)
    q = q.where('status', 'in', ['PENDING','IN_PROGRESS','CLIENT_PENDING','APPROVAL_PENDING'])
         .where('dueDateYmd', '>=', fromYmd)
         .where('dueDateYmd', '<=', toYmd);
  }

  const snap = await q.limit(800).get();
  const tasks = snap.docs.map(d => ({ id:d.id, ...d.data() }));
  const clientsMap = await loadClientsMap(tasks.map(t => t.clientId));

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Compliance Management';
  wb.created = new Date();

  const ws = wb.addWorksheet('Tasks');
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  ws.addRow([
    'TaskId','Client','Title','Category','Type',
    'Start (DD-MM-YYYY)','Due (DD-MM-YYYY)',
    'Status','Assignee',
    'StatusNote','DelayReason','DelayNotes',
    'CompletedAt (IST)',
    'CalendarEventId'
  ]);
  ws.getRow(1).font = { bold: true };
  ws.columns.forEach(c => { c.width = 22; });
  ws.getColumn(2).width = 30;
  ws.getColumn(3).width = 42;

  tasks.sort((a,b)=>String(a.dueDateYmd||'').localeCompare(String(b.dueDateYmd||'')));

  for (const t of tasks) {
    const c = clientsMap.get(t.clientId) || {};
    ws.addRow([
      t.id,
      c.name || t.clientId || '',
      t.title || '',
      t.category || '',
      t.type || '',
      ymdToDmy(t.startDateYmd),
      ymdToDmy(t.dueDateYmd),
      t.status || '',
      t.assignedToEmail || '',
      t.statusNote || '',
      t.delayReason || '',
      t.delayNotes || '',
      tsToIstString(t.completedAt),
      t.calendarEventId || t.calendarStartEventId || ''
    ]);
  }

  const buf = await wb.xlsx.writeBuffer();

  const label =
    mode === 'OVERDUE' ? `overdue_asof_${ymdToDmy(todayYmd)}` :
    mode === 'APPROVAL_PENDING' ? `approval_pending_${ymdToDmy(todayYmd)}` :
    `${ymdToDmy(fromYmd)}_to_${ymdToDmy(toYmd)}`;

  return json(event, 200, {
    ok: true,
    fileName: `quick_${mode.toLowerCase()}_${label}.xlsx`,
    count: tasks.length,
    base64: Buffer.from(buf).toString('base64')
  });
});