// Pure logic: dates and derived state. No DOM, no storage.
// Dates are local 'YYYY-MM-DD' strings. Weeks run Sunday to Saturday.

export const STATUSES = ['To do', 'Doing', 'Done'];
export const PRIORITIES = ['Low', 'Normal', 'High'];
export const PROJECT_STATUSES = ['Active', 'On hold', 'Done'];

const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const MONTH = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const DOW = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
const DAY = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];

export const pad = n => String(n).padStart(2, '0');
export const iso = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
export const parse = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
export const todayIso = () => iso(new Date());
export const addDays = (s, n) => { const d = parse(s); d.setDate(d.getDate() + n); return iso(d); };
export const diffDays = (a, b) => Math.round((parse(b) - parse(a)) / 86400000); // b - a
export const weekStart = s => addDays(s, -parse(s).getDay());
export const isDate = s => typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s);

export const fmtShort = s => { const d = parse(s); return `${MON[d.getMonth()]} ${d.getDate()}`; };
export const fmtDay = s => { const d = parse(s); return `${DOW[d.getDay()]} ${MON[d.getMonth()]} ${d.getDate()}`; };
export const fmtLong = s => { const d = parse(s); return `${DAY[d.getDay()]}, ${MONTH[d.getMonth()]} ${d.getDate()}`; };
export const dow = s => DOW[parse(s).getDay()];
export const monthName = s => MONTH[parse(s).getMonth()];
export const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;

export const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

// done | late (overdue) | soon (due today) | track (in progress) | open
export function taskState(t, today) {
  if (t.status === 'Done') return 'done';
  if (t.due && t.due < today) return 'late';
  if (t.due === today) return 'soon';
  if (t.status === 'Doing') return 'track';
  return 'open';
}

export function dueText(t, today) {
  if (t.status === 'Done') return t.doneOn ? `Done ${fmtShort(t.doneOn)}` : 'Done';
  if (!t.due) return 'No due date';
  const n = diffDays(today, t.due);
  if (n < 0) return `${plural(-n, 'day')} overdue`;
  if (n === 0) return 'Due today';
  if (n === 1) return 'Due tomorrow';
  if (n < 7) return `Due ${dow(t.due)}`;
  return `Due ${fmtShort(t.due)}`;
}

// One sentence for the detail card lead.
export function taskLead(t, today) {
  const due = dueText(t, today);
  if (t.status === 'Done') return due;
  if (t.status === 'Doing') return `In progress. ${due}.`;
  if (t.start && t.start > today) return `Starts ${fmtShort(t.start)}. ${due}.`;
  return `${due}.`;
}

export function nextAction(t) {
  if (t.status === 'To do') return { label: 'Start task', status: 'Doing' };
  if (t.status === 'Doing') return { label: 'Mark done', status: 'Done' };
  return { label: 'Reopen task', status: 'To do' };
}

export function setStatus(t, status, today) {
  t.status = status;
  t.doneOn = status === 'Done' ? today : '';
  return t;
}

const rank = { late: 0, soon: 1, track: 2, open: 3, done: 4 };
const prio = { High: 0, Normal: 1, Low: 2 };
export function sortTasks(list, today) {
  return [...list].sort((a, b) =>
    rank[taskState(a, today)] - rank[taskState(b, today)] ||
    (a.due || '9999').localeCompare(b.due || '9999') ||
    (prio[a.priority] ?? 1) - (prio[b.priority] ?? 1) ||
    a.title.localeCompare(b.title));
}

export function projectStats(p, tasks, today) {
  const mine = tasks.filter(t => t.projectId === p.id);
  const done = mine.filter(t => t.status === 'Done').length;
  const states = mine.map(t => taskState(t, today));
  const late = states.filter(s => s === 'late').length;
  const soon = states.filter(s => s === 'soon').length;
  const doing = states.filter(s => s === 'track').length;
  const open = mine.length - done;
  let state = 'open';
  if (p.status === 'Done') state = 'done';
  else if (late || (p.due && p.due < today && open)) state = 'late';
  else if (soon || (p.due && diffDays(today, p.due) <= 3 && open)) state = 'soon';
  else if (doing) state = 'track';
  // Span for the timeline: project dates if set, else the spread of its tasks.
  const dates = mine.flatMap(t => [t.start, t.due]).filter(Boolean).sort();
  const start = p.start || dates[0] || '';
  const end = p.due || dates[dates.length - 1] || '';
  return { total: mine.length, done, open, late, soon, doing, state, start, end };
}

export function projectStatusText(p, s, today) {
  if (p.status === 'Done') return 'Done';
  if (p.status === 'On hold') return 'On hold';
  if (s.late) return `${s.late} overdue`;
  if (p.due && p.due < today && s.open) return `Past due ${fmtShort(p.due)}`;
  if (s.soon) return `${s.soon} due today`;
  if (p.due) return `Due ${fmtShort(p.due)}`;
  if (!s.total) return 'No tasks yet';
  return s.open ? `${s.open} open` : 'All done';
}

// Today view groups. Each open task lands in exactly one group.
export function todayGroups(tasks, today, nowText = '') {
  const open = tasks.filter(t => t.status !== 'Done' && !(nowText && t.snooze && t.snooze > nowText));
  const horizon = addDays(today, 7);
  const g = { late: [], soon: [], doing: [], next: [], later: [] };
  for (const t of open) {
    const s = taskState(t, today);
    if (s === 'late') g.late.push(t);
    else if (s === 'soon') g.soon.push(t);
    else if (s === 'track') g.doing.push(t);
    else if (t.due && t.due <= horizon) g.next.push(t);
    else g.later.push(t);
  }
  for (const k in g) g[k] = sortTasks(g[k], today);
  return g;
}

export function todaySubline(g) {
  const parts = [];
  if (g.late.length) parts.push({ text: `${g.late.length} overdue`, loud: true });
  parts.push({ text: `${g.soon.length} due today` });
  if (g.doing.length) parts.push({ text: `${g.doing.length} in progress` });
  parts.push({ text: `${g.next.length} in the next 7 days` });
  return parts;
}

export function weekDays(ws) { return Array.from({ length: 7 }, (_, i) => addDays(ws, i)); }

// Bar geometry for the timeline. A task with only a due date is a one-day marker.
export function barSpan(start, end) {
  if (!end && !start) return null;
  const s = start || end, e = end || start;
  return s <= e ? { s, e } : { s: e, e: s };
}
