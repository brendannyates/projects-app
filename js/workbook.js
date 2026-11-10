// Excel workbook export/import with ExcelJS (vendor/exceljs.min.js, loaded as a global).
import { STATUSES, PRIORITIES, PROJECT_STATUSES, uid, weekStart, addDays, pad } from './core.js';
import { normRemind } from './ops.js';

const ROWS = 1000;          // rows prepared for validation and the timeline
const WEEKS = 30;           // timeline width
const C = {                 // tokens.css, light mode, as ARGB
  ink: 'FF1B211F', muted: 'FF66706B', faint: 'FFA3ABA6', line: 'FFE9EBE6', quiet: 'FFF7F8F5',
  accent: 'FF1E6B59', accentSoft: 'FFDDEDE6', warn: 'FF8F5400', warnSoft: 'FFFAE9CF',
  danger: 'FFB3261E', dangerSoft: 'FFF9DCD9', white: 'FFFFFFFF', done: 'FFC9CFCB'
};
const FONT = 'Segoe UI';

// Excel dates: build from UTC parts so no timezone shift.
const toXl = s => { if (!s) return null; const [y, m, d] = s.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
function fromXl(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'object' && !(v instanceof Date)) v = v.result ?? v.text ?? '';
  if (v instanceof Date) return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())}`;
  if (typeof v === 'number') return fromXl(new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000));
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) return `${m[3].length === 2 ? '20' + m[3] : m[3]}-${pad(m[1])}-${pad(m[2])}`;
  return '';
}
function text(v) {
  if (v == null) return '';
  if (v instanceof Date) return fromXl(v);
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(r => r.text).join('');
    if ('result' in v) return text(v.result);
    if (v.text) return String(v.text);
    return '';
  }
  return String(v).trim();
}
function remindXl(v) {
  if (v instanceof Date) return `${fromXl(v)} ${pad(v.getUTCHours())}:${pad(v.getUTCMinutes())}`;
  if (typeof v === 'number') { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v * 1440) * 60000); return remindXl(d); }
  return normRemind(text(v));
}
const pick = (v, list, def) => list.find(x => x.toLowerCase() === text(v).toLowerCase()) || def;

function styleHeader(row) {
  row.height = 22;
  row.eachCell(c => {
    c.font = { name: FONT, bold: true, color: { argb: C.ink }, size: 10.5 };
    c.alignment = { vertical: 'middle' };
  });
}

export async function buildWorkbook(data, today) {
  const XL = window.ExcelJS;
  const wb = new XL.Workbook();
  wb.creator = 'Projects app';
  wb.created = new Date();

  const pName = Object.fromEntries(data.projects.map(p => [p.id, p.name]));
  const pOrder = Object.fromEntries(data.projects.map((p, i) => [p.id, i]));
  const tasks = [...data.tasks].sort((a, b) =>
    (pOrder[a.projectId] ?? 1e6) - (pOrder[b.projectId] ?? 1e6) ||
    (a.start || a.due || '9999').localeCompare(b.start || b.due || '9999') ||
    a.title.localeCompare(b.title));

  // Sheets in tab order.
  const tl = wb.addWorksheet('Timeline', { views: [{ state: 'frozen', xSplit: 5, ySplit: 3, showGridLines: false }] });
  const ts = wb.addWorksheet('Tasks', { views: [{ state: 'frozen', ySplit: 1 }] });
  const ps = wb.addWorksheet('Projects', { views: [{ state: 'frozen', ySplit: 1 }] });
  const rm = wb.addWorksheet('Read me', { views: [{ showGridLines: false }] });

  // ---- Tasks ----
  ts.columns = [
    { key: 'id', width: 13 }, { key: 'project', width: 28 }, { key: 'title', width: 42 },
    { key: 'status', width: 10 }, { key: 'priority', width: 10 }, { key: 'start', width: 12 },
    { key: 'due', width: 12 }, { key: 'doneOn', width: 12 }, { key: 'remind', width: 16 }, { key: 'reminded', width: 10 },
    { key: 'notes', width: 60 }
  ];
  const taskRows = tasks.map(t => [t.id, pName[t.projectId] || '', t.title, t.status, t.priority,
    toXl(t.start), toXl(t.due), toXl(t.doneOn), t.remind || '', t.reminded ? 'Yes' : '', t.notes]);
  ts.addTable({
    name: 'tblTasks', ref: 'A1', headerRow: true,
    style: { theme: 'TableStyleLight1', showRowStripes: false },
    columns: ['ID', 'Project', 'Task', 'Status', 'Priority', 'Start', 'Due', 'Done on', 'Remind at', 'Reminded', 'Notes'].map(name => ({ name, filterButton: true })),
    rows: taskRows.length ? taskRows : [['', '', '', '', '', null, null, null, '', '', '']]
  });
  styleHeader(ts.getRow(1));
  ['F', 'G', 'H'].forEach(col => { ts.getColumn(col).numFmt = 'mmm d, yyyy'; });
  ts.getColumn('A').font = { name: FONT, color: { argb: C.faint }, size: 9 };
  ts.getColumn('K').alignment = { wrapText: true, vertical: 'top' };
  ts.dataValidations.add(`B2:B${ROWS}`, { type: 'list', allowBlank: true, formulae: [`Projects!$B$2:$B$${ROWS}`] });
  ts.dataValidations.add(`D2:D${ROWS}`, { type: 'list', allowBlank: true, formulae: [`"${STATUSES.join(',')}"`] });
  ts.dataValidations.add(`E2:E${ROWS}`, { type: 'list', allowBlank: true, formulae: [`"${PRIORITIES.join(',')}"`] });
  ts.addConditionalFormatting({
    ref: `A2:K${ROWS}`,
    rules: [
      { type: 'expression', priority: 1, formulae: ['$D2="Done"'], style: { font: { color: { argb: C.faint } } } },
      { type: 'expression', priority: 2, formulae: ['AND(ISNUMBER($G2),$G2<TODAY())'],
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.dangerSoft } }, font: { color: { argb: C.danger } } } },
      { type: 'expression', priority: 3, formulae: ['AND(ISNUMBER($G2),$G2=TODAY())'],
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.warnSoft } } } }
    ]
  });

  // ---- Projects ----
  ps.columns = [
    { key: 'id', width: 13 }, { key: 'name', width: 32 }, { key: 'status', width: 10 },
    { key: 'start', width: 12 }, { key: 'due', width: 12 }, { key: 'notes', width: 44 },
    { key: 'open', width: 11 }, { key: 'done', width: 11 }, { key: 'pct', width: 10 }
  ];
  const projRows = data.projects.map((p, i) => {
    const r = i + 2;
    return [p.id, p.name, p.status, toXl(p.start), toXl(p.due), p.notes,
      { formula: `COUNTIFS(Tasks!$B:$B,B${r},Tasks!$D:$D,"<>Done",Tasks!$C:$C,"<>")` },
      { formula: `COUNTIFS(Tasks!$B:$B,B${r},Tasks!$D:$D,"Done")` },
      { formula: `IF(G${r}+H${r}=0,"",H${r}/(G${r}+H${r}))` }];
  });
  ps.addTable({
    name: 'tblProjects', ref: 'A1', headerRow: true,
    style: { theme: 'TableStyleLight1', showRowStripes: false },
    columns: ['ID', 'Project', 'Status', 'Start', 'Due', 'Notes', 'Open tasks', 'Done tasks', 'Progress'].map(name => ({ name, filterButton: true })),
    rows: projRows.length ? projRows : [['', '', '', null, null, '', '', '', '']]
  });
  styleHeader(ps.getRow(1));
  ['D', 'E'].forEach(col => { ps.getColumn(col).numFmt = 'mmm d, yyyy'; });
  ps.getColumn('I').numFmt = '0%';
  ps.getColumn('A').font = { name: FONT, color: { argb: C.faint }, size: 9 };
  ps.dataValidations.add(`C2:C${ROWS}`, { type: 'list', allowBlank: true, formulae: [`"${PROJECT_STATUSES.join(',')}"`] });

  // ---- Timeline (Gantt). Follows the Tasks sheet row by row; bars are conditional formats. ----
  const dated = data.tasks.flatMap(t => [t.start, t.due]).filter(Boolean).sort();
  let first = weekStart(addDays(today, -14));
  if (dated[0] && dated[0] < first) first = weekStart(dated[0]);
  if (addDays(first, WEEKS * 7) < today) first = weekStart(addDays(today, -14));

  tl.getColumn(1).width = 24; tl.getColumn(2).width = 34;
  tl.getColumn(3).width = 9; tl.getColumn(4).width = 9; tl.getColumn(5).width = 8;
  for (let w = 0; w < WEEKS; w++) tl.getColumn(6 + w).width = 6.5;

  tl.getCell('A1').value = 'Timeline';
  tl.getCell('A1').font = { name: FONT, size: 16, bold: true, color: { argb: C.ink } };
  tl.getCell('B1').value = 'First week starts (a Sunday). Change it to scroll:';
  tl.getCell('B1').alignment = { horizontal: 'right' };
  tl.getCell('B1').font = { name: FONT, size: 9, color: { argb: C.muted } };
  tl.mergeCells('C1:D1');
  const anchor = tl.getCell('C1');
  anchor.value = toXl(first); anchor.numFmt = 'mmm d, yyyy';
  anchor.font = { name: FONT, bold: true, color: { argb: C.accent } };
  anchor.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.accentSoft } };
  anchor.dataValidation = { type: 'date', operator: 'greaterThan', formulae: [new Date(Date.UTC(2000, 0, 1))],
    showErrorMessage: true, error: 'Enter a date. Weeks run Sunday to Saturday.' };
  tl.getCell('F1').value = 'Bars: green planned, red overdue, grey done. Shaded column is this week.';
  tl.getCell('F1').font = { name: FONT, size: 9, color: { argb: C.muted } };

  const hdr = tl.getRow(3);
  ['Project', 'Task', 'Start', 'Due', 'Status'].forEach((h, i) => { hdr.getCell(i + 1).value = h; });
  for (let w = 0; w < WEEKS; w++) {
    const col = 6 + w, L = colLetter(col), prev = colLetter(col - 1);
    // Sunday of the week the anchor falls in, so a non-Sunday anchor still lines up.
    tl.getCell(3, col).value = { formula: w === 0 ? '$C$1-WEEKDAY($C$1)+1' : `${prev}3+7` };
    tl.getCell(3, col).numFmt = 'mmm d';
    tl.getCell(3, col).alignment = { horizontal: 'center', textRotation: 90 };
    tl.getCell(2, col).value = { formula: w === 0 ? `TEXT(${L}3,"mmm")` : `IF(MONTH(${L}3)<>MONTH(${prev}3),TEXT(${L}3,"mmm"),"")` };
    tl.getCell(2, col).font = { name: FONT, size: 9, bold: true, color: { argb: C.muted } };
  }
  hdr.height = 44;
  hdr.eachCell(c => { c.font = { name: FONT, bold: true, size: 9.5, color: { argb: C.ink } }; c.border = { bottom: { style: 'thin', color: { argb: C.faint } } }; });
  hdr.getCell(1).alignment = hdr.getCell(2).alignment = { vertical: 'bottom' };

  const last = 3 + ROWS - 1;
  for (let i = 0; i < ROWS - 1; i++) {
    const r = 4 + i, src = i + 2, row = tl.getRow(r);
    const ref = c => `IF(Tasks!$C${src}="","",IF(Tasks!${c}${src}="","",Tasks!${c}${src}))`;
    row.getCell(1).value = { formula: ref('B') };
    row.getCell(2).value = { formula: ref('C') };
    row.getCell(3).value = { formula: ref('F') };
    row.getCell(4).value = { formula: ref('G') };
    row.getCell(5).value = { formula: ref('D') };
    row.getCell(3).numFmt = row.getCell(4).numFmt = 'mmm d';
    for (let c = 1; c <= 5; c++) row.getCell(c).font = { name: FONT, size: 9.5, color: { argb: c === 2 ? C.ink : C.muted } };
  }
  const gantt = `F4:${colLetter(5 + WEEKS)}${last}`;
  const inSpan = 'AND(ISNUMBER($D4),F$3<=$D4,F$3+6>=IF(ISNUMBER($C4),$C4,$D4))';
  const fill = argb => ({ fill: { type: 'pattern', pattern: 'solid', bgColor: { argb } } });
  tl.addConditionalFormatting({
    ref: gantt,
    rules: [
      { type: 'expression', priority: 1, formulae: [`AND($E4="Done",${inSpan})`], style: fill(C.done) },
      { type: 'expression', priority: 2, formulae: [`AND($D4<TODAY(),${inSpan})`], style: fill(C.danger) },
      { type: 'expression', priority: 3, formulae: [inSpan], style: fill(C.accent) },
      { type: 'expression', priority: 4, formulae: ['AND(TODAY()>=F$3,TODAY()<=F$3+6)'], style: fill(C.warnSoft) },
      { type: 'expression', priority: 5, formulae: ['AND($A4<>"",$A4<>$A3)'], style: { border: { top: { style: 'thin', color: { argb: C.faint } } } } }
    ]
  });
  tl.addConditionalFormatting({
    ref: `F3:${colLetter(5 + WEEKS)}3`,
    rules: [{ type: 'expression', priority: 6, formulae: ['AND(TODAY()>=F$3,TODAY()<=F$3+6)'], style: { ...fill(C.warnSoft), font: { color: { argb: C.warn }, bold: true } } }]
  });
  tl.addConditionalFormatting({
    ref: `A4:E${last}`,
    rules: [
      { type: 'expression', priority: 7, formulae: ['AND($A4<>"",$A4<>$A3)'], style: { border: { top: { style: 'thin', color: { argb: C.faint } } } } },
      { type: 'expression', priority: 8, formulae: ['$E4="Done"'], style: { font: { color: { argb: C.faint } } } }
    ]
  });

  // ---- Read me ----
  rm.getColumn(1).width = 100;
  const lines = [
    ['Projects workbook', 16, true],
    [`Saved from the Projects app on ${new Date().toLocaleString()}.`, 10, false],
    ['', 10, false],
    ['How this works', 12, true],
    ['The app keeps your data on this PC. This workbook is a copy you can open, edit and share.', 10, false],
    ['To bring edits back into the app: Workbook > Load from workbook. That replaces the app data with this file (you can undo once).', 10, false],
    ['Keep the ID column. It is how the app matches rows. New rows can leave ID blank.', 10, false],
    ['Tasks: one row per task. Project must match a name on the Projects sheet (the dropdown helps). Status: To do, Doing, Done.', 10, false],
    ['Remind at: a date and time like 2026-11-12 09:00. The app shows a notification then, while it is open.', 10, false],
    ['Projects: Start and Due are optional. Open tasks, Done tasks and Progress are formulas.', 10, false],
    ['Timeline: follows the Tasks sheet row by row. Change the date in C1 to scroll. Weeks run Sunday to Saturday.', 10, false],
    ['', 10, false],
    ['Close this file in Excel before the app saves to it, or the save will fail.', 10, true]
  ];
  lines.forEach(([v, size, bold], i) => {
    const c = rm.getCell(i + 1, 1); c.value = v;
    c.font = { name: FONT, size, bold, color: { argb: C.ink } }; c.alignment = { wrapText: true };
  });

  wb.views = [{ activeTab: 0 }];
  return wb.xlsx.writeBuffer();
}

function colLetter(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

function readSheet(ws) {
  if (!ws) return [];
  const head = {};
  ws.getRow(1).eachCell((c, i) => { head[text(c.value).toLowerCase()] = i; });
  const rows = [];
  ws.eachRow((row, r) => {
    if (r === 1) return;
    const get = name => { const i = head[name]; return i ? row.getCell(i).value : null; };
    rows.push(get);
  });
  return rows;
}

export async function readWorkbook(buffer) {
  const wb = new window.ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ps = wb.getWorksheet('Projects'), ts = wb.getWorksheet('Tasks');
  if (!ts) throw new Error('No Tasks sheet');

  const projects = [], byName = {};
  for (const get of readSheet(ps)) {
    const name = text(get('project'));
    if (!name) continue;
    const p = { id: text(get('id')) || uid(), name, status: pick(get('status'), PROJECT_STATUSES, 'Active'),
      start: fromXl(get('start')), due: fromXl(get('due')), notes: text(get('notes')) };
    projects.push(p); byName[name.toLowerCase()] = p;
  }
  const tasks = [];
  for (const get of readSheet(ts)) {
    const title = text(get('task'));
    if (!title) continue;
    const pn = text(get('project'));
    let p = pn ? byName[pn.toLowerCase()] : null;
    if (pn && !p) { p = { id: uid(), name: pn, status: 'Active', start: '', due: '', notes: '' }; projects.push(p); byName[pn.toLowerCase()] = p; }
    const status = pick(get('status'), STATUSES, 'To do');
    tasks.push({ id: text(get('id')) || uid(), projectId: p ? p.id : '', title, status,
      priority: pick(get('priority'), PRIORITIES, 'Normal'), start: fromXl(get('start')), due: fromXl(get('due')),
      doneOn: status === 'Done' ? fromXl(get('done on')) : '', remind: remindXl(get('remind at')), reminded: /^y/i.test(text(get('reminded'))) ? 'Yes' : '',
      notes: text(get('notes')) });
  }
  return { projects, tasks };
}
