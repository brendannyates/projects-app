import * as C from './core.js';
import * as S from './store.js';
import { buildWorkbook, readWorkbook } from './workbook.js';
import { applyOps, parseInbox, addNote, normRemind, stampText, dueReminders } from './ops.js';

const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

const st = {
  data: S.load(),
  settings: S.loadSettings(),
  today: C.todayIso(),
  view: 'today',
  selTask: null,
  selProject: null,
  weekOf: C.weekStart(C.todayIso()),
  projFilter: 'active',
  handle: null,
  persisted: 'unknown',
  sync: 'ok',
  installPrompt: null
};
const FSA = 'showSaveFilePicker' in window;
const STANDALONE = matchMedia('(display-mode: standalone), (display-mode: window-controls-overlay), (display-mode: fullscreen)').matches;
const PX = { day: 32, week: 12, quarter: 4 };

const proj = id => st.data.projects.find(p => p.id === id);
const task = id => st.data.tasks.find(t => t.id === id);
const item = id => st.data.items.find(i => i.id === id);
const tickBtn = t => `<button class="tick ${t.status === 'Done' ? 'on' : t.status === 'Doing' ? 'doing' : ''}" data-act="toggle" data-id="${t.id}" aria-label="${esc(t.status === 'Done' ? `Reopen ${t.title}` : `Mark ${t.title} done`)}">✓</button>`;
const pname = id => (id && proj(id)?.name) || 'No project';

// ---------- saving ----------
// Every record gets an `updated` stamp when it changes, so edits made here can be merged
// with edits another device saved to the workbook in the meantime.
let baseline = new Map();
const recJson = r => { const { updated, ...rest } = r; return JSON.stringify(rest); };
function primeStamps() { baseline = new Map([...st.data.projects, ...st.data.tasks, ...st.data.items, ...st.data.digest].map(r => [r.id, recJson(r)])); }
function stamp() {
  const now = Date.now(), seen = new Map();
  for (const r of [...st.data.projects, ...st.data.tasks, ...st.data.items, ...st.data.digest]) {
    const j = recJson(r);
    if (baseline.get(r.id) !== j) r.updated = now;
    seen.set(r.id, j);
  }
  const changed = seen.size !== baseline.size || [...seen].some(([id, j]) => baseline.get(id) !== j);
  baseline = seen;
  return changed;
}

let syncTimer = null;
function commit(msg, undo) {
  if (stamp() && st.handle) st.data.meta.dirty = true;
  S.save(st.data);
  render();
  if (msg) toast(msg, undo);
  if (st.handle) { clearTimeout(syncTimer); syncTimer = setTimeout(() => syncNow(), 300); }
}
function snapshotFor(fn) { const before = JSON.stringify(st.data); return () => { st.data = S.normalise(JSON.parse(before)); fn?.(); commit('Undone'); }; }

let toastTimer;
function toast(msg, undo) {
  const el = $('#toast');
  el.innerHTML = esc(msg) + (undo ? ' <button type="button">Undo</button>' : '');
  if (undo) el.querySelector('button').onclick = () => { el.classList.remove('show'); undo(); };
  el.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), undo ? 6000 : 2800);
}

// ---------- rendering ----------
function render() {
  document.querySelectorAll('.rail .item[data-view]').forEach(b => { if (b.dataset.view === st.view) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current'); });
  document.querySelectorAll('.view').forEach(v => { v.hidden = v.id !== 'v-' + st.view; });
  const g = C.todayGroups(st.data.tasks, st.today, stampText());
  const c = $('#cToday'); c.hidden = !g.late.length; c.textContent = g.late.length;
  $('#cBook').hidden = !needsBackup();
  ({ today: renderToday, week: renderWeek, projects: renderProjects, timeline: renderTimeline, workbook: renderWorkbook })[st.view]();
  renderConn();
}

function renderConn() {
  const dot = $('#connDot'), txt = $('#connText'), ban = $('#syncBanner');
  let warn = false, line, banner = '';
  const name = st.handle?.name || 'the workbook';
  if (S.DEMO) { warn = true; line = 'Demo data'; }
  else if (!st.handle) { warn = st.data.tasks.length > 0; line = 'On this PC only'; }
  else if (st.sync === 'needperm') {
    warn = true; line = 'Workbook not connected';
    banner = `<span>Reconnect to ${esc(name)} to load the latest from OneDrive. When Edge asks, choose <b>Allow on every visit</b> so it stops asking.</span><button class="btn sm primary" data-act="reconnect">Reconnect workbook</button>`;
  } else if (st.sync === 'locked') {
    warn = true; line = 'Waiting for Excel to close';
    banner = `<span>Changes aren't in ${esc(name)} yet because it's open in Excel. Close it there and they save within 30 sec. Your changes are kept here.</span>`;
  } else if (st.sync === 'error') {
    warn = true; line = "Couldn't reach the workbook";
    banner = `<span>Couldn't reach ${esc(name)}. Check OneDrive is running and the file hasn't moved. Your changes are kept here.</span><button class="btn sm ghost" data-act="syncnow">Try again</button><button class="btn sm mute" data-act="wbopen">Find the file</button>`;
  } else if (st.data.meta.dirty) line = 'Saving to OneDrive...';
  else line = st.data.meta.lastSync ? `Synced ${relTime(new Date(st.data.meta.lastSync).toISOString())}` : 'Connecting...';
  dot.className = warn ? 'cdot warn' : 'cdot';
  $('#syncBtn').classList.toggle('warn', !!st.handle && ['needperm', 'locked', 'error'].includes(st.sync));
  $('#syncBtn').title = st.handle ? `Sync now. ${line}` : 'Sync: connect a workbook first';
  txt.textContent = line;
  ban.hidden = !banner; ban.innerHTML = banner;
}
function relTime(isoTs) {
  const d = new Date(isoTs), day = C.iso(d);
  const t = `${d.getHours()}:${C.pad(d.getMinutes())}`;
  if (day === st.today) return `at ${t}`;
  if (day === C.addDays(st.today, -1)) return `yesterday ${t}`;
  return `${C.fmtShort(day)}`;
}
function needsBackup() {
  if (S.DEMO) return false;
  if (st.handle) return ['needperm', 'locked', 'error'].includes(st.sync);
  return st.data.tasks.length > 0;
}

function taskRow(t, opts = {}) {
  const s = C.taskState(t, st.today);
  const bits = [];
  if (opts.showProject !== false) bits.push(esc(pname(t.projectId)));
  if (t.priority === 'High') bits.push('<span class="hi">High priority</span>');
  if (t.start && t.start > st.today && t.status !== 'Done') bits.push(`Starts ${C.fmtShort(t.start)}`);
  if (t.remind && !t.reminded && t.status !== 'Done') bits.push(`<span class="rem">Reminder ${esc(remindText(t.remind))}</span>`);
  const tick = t.status === 'Done' ? 'on' : t.status === 'Doing' ? 'doing' : '';
  const label = t.status === 'Done' ? `Reopen ${t.title}` : `Mark ${t.title} done`;
  return `<div class="trow${st.selTask === t.id ? ' sel' : ''}" data-state="${s}">
    <button class="tick ${tick}" data-act="toggle" data-id="${t.id}" aria-label="${esc(label)}">✓</button>
    <button class="tmain" data-act="select" data-id="${t.id}"><b>${esc(t.title)}</b><small>${bits.join(' · ') || '&nbsp;'}</small></button>
    <span class="st">${esc(C.dueText(t, st.today))}</span>
  </div>`;
}

// ----- Today -----
function renderToday() {
  const now = new Date();
  $('#clockT').textContent = `${now.getHours()}:${C.pad(now.getMinutes())}`;
  $('#clockD').textContent = C.fmtLong(st.today);
  const g = C.todayGroups(st.data.tasks, st.today, stampText());
  const rems = st.data.tasks.filter(t => t.remind && !t.reminded && t.status !== 'Done' && t.remind.startsWith(st.today)).length;
  $('#todaySub').innerHTML = [...C.todaySubline(g), ...(rems ? [{ text: `${C.plural(rems, 'reminder')} today` }] : [])].map(p => p.loud ? `<strong>${esc(p.text)}</strong>` : esc(p.text)).join(', ');

  const doneToday = st.data.tasks.filter(t => t.status === 'Done' && t.doneOn === st.today);
  const groups = [
    ['Overdue', g.late], ['Due today', g.soon, 'Nothing due today.'], ['In progress', g.doing],
    ['Next 7 days', g.next], ['Later or no date', g.later, null, 6], ['Done today', doneToday]
  ];
  let html = '';
  if (!st.data.tasks.length) html = `<div class="card group"><p class="empty">No tasks yet. Choose New task or press N to add your first one.</p></div>`;
  else for (const [title, list, emptyMsg, cap] of groups) {
    if (!list.length && !emptyMsg) continue;
    const shown = cap && !st.showAllLater ? list.slice(0, cap) : list;
    html += `<div class="card group"><div class="gh"><h2>${title}</h2><span>${list.length || ''}</span></div>
      ${list.length ? shown.map(t => taskRow(t)).join('') : `<p class="empty">${emptyMsg}</p>`}
      ${cap && list.length > cap && !st.showAllLater ? `<button class="btn sm ghost" data-act="more">Show all ${list.length}</button>` : ''}</div>`;
  }
  const snoozed = st.data.tasks.filter(t => t.status !== 'Done' && t.snooze && t.snooze > stampText()).sort((a, b) => a.snooze.localeCompare(b.snooze));
  if (snoozed.length) html += `<div class="card group"><div class="gh"><h2>Snoozed</h2><span>${snoozed.length}</span></div>
    ${snoozed.map(t => `<div class="trow" data-state="open">${tickBtn(t)}<button class="tmain" data-act="select" data-id="${t.id}"><b>${esc(t.title)}</b><small>${esc(pname(t.projectId))}</small></button><span class="st">Back ${esc(remindText(t.snooze))}</span></div>`).join('')}</div>`;
  $('#todayLists').innerHTML = digestCard() + html;

  if (!st.selTask || !task(st.selTask)) st.selTask = (g.late[0] || g.soon[0] || g.doing[0] || g.next[0] || g.later[0] || {}).id || null;
  $('#todayDetail').innerHTML = taskDetail(st.selTask && task(st.selTask));
  // re-mark selection after choosing default
  document.querySelectorAll('#todayLists .trow').forEach(r => r.classList.toggle('sel', r.querySelector('.tmain').dataset.id === st.selTask));
}

function refBlock(p) {
  if (!p.refUrl) return `<p class="hint" style="margin:6px 0 0"><button class="btn sm ghost" data-act="editproj" data-id="${p.id}" style="padding:0 6px">Add a reference file</button> Link the latest build sheet or flow page.</p>`;
  const web = /^https?:\/\//i.test(p.refUrl);
  return `<div class="ref"><span class="muted">Reference</span>
    ${web ? `<a href="${esc(p.refUrl)}" target="_blank" rel="noopener">${esc(p.refName || 'Open file')}</a>` : `<span>${esc(p.refName || p.refUrl)}</span><button class="btn sm ghost" data-act="copyref" data-id="${p.id}">Copy path</button>`}</div>`;
}

function digestCard() {
  const items = st.data.digest.filter(g => g.date === st.today);
  const older = st.data.digest.filter(g => g.kind === 'action' && g.status === 'new' && g.date < st.today && g.date >= C.addDays(st.today, -3));
  const meetings = items.filter(g => g.kind === 'meeting').sort((a, b) => a.time.localeCompare(b.time));
  const actions = [...items.filter(g => g.kind === 'action' && g.status === 'new'), ...older];
  const handled = items.filter(g => g.kind === 'action' && g.status !== 'new').length;
  if (!meetings.length && !actions.length) return '';
  const meet = g => `<div class="dgm"><span class="dgt">${esc(g.time.replace(/^0/, ''))}</span><div><b>${esc(g.title)}</b>
      ${g.notes ? `<p class="notes">${esc(g.notes)}</p>` : ''}</div>
      ${g.status === 'new' ? `<button class="btn sm ghost" data-act="dgadd" data-id="${g.id}">Add prep to my day</button>` : `<span class="st">${g.status === 'added' ? 'Added' : ''}</span>`}</div>`;
  const act = g => `<div class="dga"><div><b>${esc(g.title)}</b><small>${esc([g.source, g.project, g.due ? C.dueText({ status: 'To do', due: g.due }, st.today) : '', g.date < st.today ? `from ${C.fmtShort(g.date)}` : ''].filter(Boolean).join(' · '))}</small>
      ${g.notes ? `<p class="notes">${esc(g.notes)}</p>` : ''}</div>
      <span class="dgbtn"><button class="btn sm ghost" data-act="dgadd" data-id="${g.id}">Add to my day</button><button class="btn sm mute" data-act="dgdismiss" data-id="${g.id}">Dismiss</button></span></div>`;
  return `<div class="card group digest"><div class="gh"><h2>Morning digest</h2><span>${actions.length ? `${actions.length} to review` : 'All reviewed'}${handled ? `, ${handled} handled` : ''}${actions.length > 3 ? ` <button class="btn sm mute" data-act="dgdismissall">Dismiss all</button>` : ''}</span></div>
    ${meetings.length ? `<div class="dgsec"><h3>On your calendar</h3>${meetings.map(meet).join('')}</div>` : ''}
    ${actions.length ? `<div class="dgsec"><h3>From your email</h3>${actions.map(act).join('')}</div>` : ''}</div>`;
}

function notesSection(kind, rec) {
  return `<div class="section"><div class="sh"><h2>Notes</h2><span>${rec.notes ? '' : 'None yet'}</span></div>
    ${rec.notes ? `<p class="notes">${esc(rec.notes)}</p>` : ''}
    <form class="notebox" data-act="addnote" data-kind="${kind}" data-id="${rec.id}">
      <input name="text" placeholder="Add a note" aria-label="Add a note" maxlength="1000" required>
      <button class="btn sm ghost" type="submit">Add note</button></form></div>`;
}
const remindText = r => { if (!r) return ''; const [d, t] = r.split(' '); return `${d === st.today ? 'Today' : C.fmtDay(d)}, ${t.replace(/^0/, '')}`; };

function taskDetail(t) {
  if (!t) return `<p class="muted" style="margin:0">Nothing selected. Pick a task to see it here.</p>`;
  const s = C.taskState(t, st.today), na = C.nextAction(t);
  const fact = (k, v) => `<dt>${k}</dt><dd>${v}</dd>`;
  return `<div class="dh"><div><div class="name">${esc(t.title)}</div>
      <div class="meta">${t.projectId ? `<button class="btn sm ghost" style="padding:0;min-height:0" data-act="openproj" data-id="${t.projectId}">${esc(pname(t.projectId))}</button>` : 'No project'}</div></div>
      <button class="btn sm mute" data-act="edit" data-id="${t.id}">Edit</button></div>
    <p class="lead" data-state="${s}">${esc(C.taskLead(t, st.today))}</p>
    <dl class="facts">
      ${fact('Status', esc(t.status))}
      ${fact('Priority', t.priority === 'High' ? '<span class="hi">High</span>' : esc(t.priority))}
      ${fact('Start', t.start ? C.fmtDay(t.start) : '<span class="muted">Not set</span>')}
      ${fact('Due', t.due ? C.fmtDay(t.due) : '<span class="muted">Not set</span>')}
      ${t.doneOn ? fact('Done', C.fmtDay(t.doneOn)) : ''}
      ${t.snooze && t.snooze > stampText() ? fact('Snoozed', `until ${remindText(t.snooze)}`) : ''}
      ${t.remind ? fact('Reminder', `<span class="rem">${remindText(t.remind)}</span>${t.reminded ? ' <span class="muted">(sent)</span>' : ''}`) : ''}
    </dl>
    <button class="btn primary" data-act="advance" data-id="${t.id}" data-status="${na.status}">${na.label}</button>
    ${t.status !== 'Done' ? `<div class="actions">
      <button class="btn sm ghost" data-act="tomorrow" data-id="${t.id}">Move to tomorrow</button>
      <button class="btn sm ghost" data-act="week" data-id="${t.id}">Push 1 week</button>
      ${t.snooze && t.snooze > stampText() ? `<button class="btn sm ghost" data-act="unsnooze" data-id="${t.id}">Unsnooze</button>` : `<button class="btn sm ghost" data-act="snooze" data-id="${t.id}">Snooze</button>`}
      ${t.remind ? `<button class="btn sm ghost" data-act="clearremind" data-id="${t.id}">Clear reminder</button>` : `<button class="btn sm ghost" data-act="edit" data-id="${t.id}">Set reminder</button>`}
      <button class="btn sm mute" data-act="delete" data-id="${t.id}">Delete</button></div>`
      : `<div class="actions"><button class="btn sm mute" data-act="delete" data-id="${t.id}">Delete</button></div>`}
    ${notesSection('task', t)}`;
}

// ----- Week -----
function renderWeek() {
  const days = C.weekDays(st.weekOf), end = days[6];
  const inWeek = st.data.tasks.filter(t => t.due >= st.weekOf && t.due <= end);
  const done = inWeek.filter(t => t.status === 'Done').length;
  const label = st.weekOf === C.weekStart(st.today) ? 'This week' : `Week of ${C.fmtShort(st.weekOf)}`;
  $('#v-week h1').textContent = label;
  $('#weekSub').textContent = `${C.fmtShort(st.weekOf)} to ${C.fmtShort(end)}. ${C.plural(inWeek.length, 'task')} due, ${done} done.`;

  const late = st.data.tasks.filter(t => t.status !== 'Done' && t.due && t.due < st.weekOf && t.due < st.today);
  $('#weekLate').innerHTML = late.length && st.weekOf <= st.today && end >= st.today
    ? `<div class="latestrip"><b>Overdue from earlier weeks</b>${C.sortTasks(late, st.today).map(chip).join('')}</div>` : '';

  $('#weekGrid').innerHTML = days.map(d => {
    const list = C.sortTasks(st.data.tasks.filter(t => t.due === d), st.today);
    const cls = d === st.today ? 'today' : d < st.today ? 'past' : '';
    return `<div class="day ${cls}" data-day="${d}">
      <div class="dhd"><b>${C.dow(d)}</b><span>${C.fmtShort(d)}</span></div>
      ${list.map(chip).join('')}
      <button class="daddbtn" data-act="addday" data-day="${d}">+ Add</button></div>`;
  }).join('');
}
function chip(t) {
  const s = C.taskState(t, st.today);
  return `<div class="chip" draggable="true" data-state="${s}" data-id="${t.id}">${tickBtn(t)}<button class="cbody" data-act="edit" data-id="${t.id}">${esc(t.title)}<small>${esc(pname(t.projectId))}${t.priority === 'High' ? ' · High' : ''}${s === 'late' ? ' · ' + esc(C.dueText(t, st.today)) : ''}</small></button></div>`;
}

// ----- Projects -----
function projectList() {
  const f = st.projFilter;
  return st.data.projects.filter(p => f === 'all' || (f === 'active' && p.status === 'Active') || (f === 'hold' && p.status === 'On hold') || (f === 'done' && p.status === 'Done'));
}
function renderProjects() {
  const all = st.data.projects, active = all.filter(p => p.status === 'Active');
  const stats = active.map(p => C.projectStats(p, st.data.tasks, st.today));
  const late = stats.filter(s => s.state === 'late').length;
  $('#projSub').innerHTML = `${active.length} active${late ? `, <strong>${late} with overdue work</strong>` : ''}, ${all.length - active.length} on hold or done`;
  document.querySelectorAll('#projFilter button').forEach(b => b.setAttribute('aria-pressed', b.dataset.f === st.projFilter));

  const rank = { late: 0, soon: 1, track: 2, open: 3, done: 4 };
  const list = projectList().map(p => ({ p, s: C.projectStats(p, st.data.tasks, st.today) }))
    .sort((a, b) => rank[a.s.state] - rank[b.s.state] || (a.p.due || '9999').localeCompare(b.p.due || '9999'));
  const loose = st.data.tasks.filter(t => !t.projectId);
  let html = list.map(({ p, s }) => projRow(p.id, p.name, s, C.projectStatusText(p, s, st.today), p)).join('');
  if (loose.length && (st.projFilter === 'active' || st.projFilter === 'all')) {
    const s = C.projectStats({ id: '', status: 'Active' }, st.data.tasks, st.today);
    html += projRow('', 'No project', s, s.late ? `${s.late} overdue` : `${s.open} open`, null);
  }
  $('#projList').innerHTML = html || `<p class="empty">No projects here. Choose New project to start one.</p>`;

  if (st.selProject === null || (st.selProject && !proj(st.selProject))) st.selProject = list[0]?.p.id ?? (loose.length ? '' : null);
  document.querySelectorAll('.prow').forEach(r => r.setAttribute('aria-pressed', r.dataset.id === st.selProject));
  $('#projDetail').innerHTML = projectDetail(st.selProject);
}
function projRow(id, name, s, status, p) {
  const sub = [p?.status === 'On hold' ? 'On hold' : '', p?.start || p?.due ? [p.start && C.fmtShort(p.start), p.due && C.fmtShort(p.due)].filter(Boolean).join(' to ') : ''].filter(Boolean).join(' · ');
  const its = st.data.items.filter(i => i.projectId === id && id), chk = its.filter(i => i.done).length;
  const all = s.total + its.length, pct = all ? Math.round((s.done + chk) / all * 100) : 0;
  const counts = [s.total ? `${s.open} open ${s.open === 1 ? 'task' : 'tasks'}` : '', its.length ? `${chk} of ${its.length} checked` : ''].filter(Boolean).join(' · ') || 'Nothing yet';
  return `<button class="prow" data-state="${s.state}" data-act="selproj" data-id="${id}" aria-label="${esc(name)}, ${esc(status)}">
    <span><b>${esc(name)}</b><small>${sub || counts}</small></span>
    <span class="prog"><span class="bar"><i style="width:${pct}%"></i></span>${counts}</span>
    <span class="st">${esc(status)}</span></button>`;
}
function projectDetail(id) {
  if (id === null) return `<p class="muted" style="margin:0">Pick a project to see its tasks.</p>`;
  const p = id ? proj(id) : null;
  const s = C.projectStats(p || { id: '', status: 'Active' }, st.data.tasks, st.today);
  const tasks = C.sortTasks(st.data.tasks.filter(t => t.projectId === id), st.today);
  const open = tasks.filter(t => t.status !== 'Done'), done = tasks.filter(t => t.status === 'Done');
  const mini = t => { const ts = C.taskState(t, st.today); return `<div class="mini" data-state="${ts}">${tickBtn(t)}<button data-act="edit" data-id="${t.id}">${esc(t.title)}</button><span class="st">${esc(C.dueText(t, st.today))}</span></div>`; };
  const its = id ? st.data.items.filter(i => i.projectId === id).sort((a, b) => (a.done ? 1 : 0) - (b.done ? 1 : 0)) : [];
  const chk = its.filter(i => i.done).length;
  const checklist = !p ? '' : `<div class="section"><div class="sh"><h2>Checklist</h2><span>${its.length ? `${chk} of ${its.length} checked` : 'The parts of this project'}</span></div>
      ${its.map(i => `<div class="citem${i.done ? ' done' : ''}">
        <button class="tick ${i.done ? 'on' : ''}" data-act="toggleitem" data-id="${i.id}" aria-label="${esc(i.done ? `Uncheck ${i.title}` : `Check off ${i.title}`)}">✓</button>
        <span class="ct" title="${esc(i.notes)}">${esc(i.title)}${i.notes ? `<small>${esc(i.notes.split('\n').pop())}</small>` : ''}</span>
        <button class="x sm" data-act="itemtask" data-id="${i.id}" title="Make this one of my tasks" aria-label="Make ${esc(i.title)} a task">+</button>
        <button class="x sm" data-act="delitem" data-id="${i.id}" title="Remove from checklist" aria-label="Remove ${esc(i.title)}">×</button></div>`).join('')}
      <form class="quick" data-act="additem" data-id="${p.id}"><input name="item" placeholder="Add a checklist item" aria-label="New checklist item" required><button class="btn sm ghost" type="submit">Add</button></form></div>`;
  const lead = p ? C.projectStatusText(p, s, st.today) : `${s.open} open`;
  return `<div class="dh"><div><div class="name">${esc(p ? p.name : 'No project')}</div>
      <div class="meta">${p ? `${esc(p.status)}${p.start ? ' · Starts ' + C.fmtShort(p.start) : ''}${p.due ? ' · Due ' + C.fmtShort(p.due) : ''}` : 'Tasks not filed under a project'}</div></div>
      ${p ? `<button class="btn sm mute" data-act="editproj" data-id="${p.id}">Edit</button>` : ''}</div>
    <p class="lead" data-state="${s.state}">${esc(lead)}</p>
    ${p ? refBlock(p) : ''}
    <div class="section"><div class="sh"><h2>My tasks</h2><span>${open.length} open</span></div>
    <form class="quick" data-act="quickadd" data-id="${id}">
      <input name="title" placeholder="Add a task" aria-label="New task title" required>
      <input name="due" type="date" aria-label="Due date" style="flex:0 0 138px">
      <button class="btn sm ghost" type="submit">Add</button>
    </form>
      <div style="margin-top:8px">${open.length ? open.map(mini).join('') : '<p class="muted" style="margin:0;font-size:14px">No open tasks.</p>'}
      ${done.length ? `<details class="donelist"><summary>${done.length} done</summary>${done.map(mini).join('')}</details>` : ''}</div></div>
    ${checklist}
    ${p ? `<div class="actions" style="margin-top:16px">
      ${p.status !== 'Done' ? `<button class="btn sm ghost" data-act="projdone" data-id="${p.id}">Mark project done</button>` : `<button class="btn sm ghost" data-act="projactive" data-id="${p.id}">Reopen project</button>`}
      ${p.status === 'Active' ? `<button class="btn sm mute" data-act="projhold" data-id="${p.id}">Put on hold</button>` : ''}
      <button class="btn sm mute" data-act="delproj" data-id="${p.id}">Delete</button></div>` : ''}
    ${p ? notesSection('project', p) : ''}`;
}

// ----- Timeline -----
let tlRange = null;
function renderTimeline() {
  const z = st.settings.timelineZoom, px = PX[z];
  document.querySelectorAll('#tlZoom button').forEach(b => b.setAttribute('aria-pressed', b.dataset.z === z));
  $('#tlDone').checked = st.settings.showDone;
  const showDone = st.settings.showDone;

  const projects = st.data.projects.filter(p => showDone || p.status !== 'Done');
  const groups = projects.map(p => ({ p, s: C.projectStats(p, st.data.tasks, st.today) }));
  groups.push({ p: null, s: null });
  const visible = t => (showDone || t.status !== 'Done') && (t.start || t.due);
  const undated = st.data.tasks.filter(t => t.status !== 'Done' && !t.start && !t.due).length;

  const dates = [st.today, ...st.data.tasks.filter(visible).flatMap(t => [t.start, t.due]), ...projects.flatMap(p => [p.start, p.due])].filter(Boolean).sort();
  const from = C.weekStart(C.addDays(dates[0] < C.addDays(st.today, -14) ? dates[0] : st.today, -14));
  const lastD = dates[dates.length - 1];
  const to = C.addDays(C.weekStart(lastD > C.addDays(st.today, 56) ? lastD : C.addDays(st.today, 56)), 13);
  const days = C.diffDays(from, to) + 1, width = days * px;
  tlRange = { from, px };
  const x = d => C.diffDays(from, d) * px;

  const grid = `background-image:repeating-linear-gradient(90deg,var(--surface-2) 0 1px,transparent 1px ${px * 7}px);`;
  // header: months and weeks
  let head = '';
  for (let d = from; d <= to; d = C.addDays(d, 1)) {
    if (d.endsWith('-01') || d === from) head += `<span class="mlab" style="left:${x(d)}px">${C.monthName(d)} ${d.slice(0, 4)}</span>`;
  }
  const wStep = z === 'quarter' ? 4 : 1;
  let wi = 0;
  for (let d = from; d <= to; d = C.addDays(d, 7), wi++) {
    const now = st.today >= d && st.today <= C.addDays(d, 6);
    if (wi % wStep === 0 || now) head += `<span class="wlab${now ? ' now' : ''}" style="left:${x(d)}px">${z === 'day' ? C.fmtShort(d) : d.slice(8).replace(/^0/, '')}</span>`;
  }
  if (z === 'day') for (let d = from; d <= to; d = C.addDays(d, 1)) if (C.parse(d).getDay()) head += `<span class="wlab" style="left:${x(d)}px;color:var(--faint)">${d.slice(8).replace(/^0/, '')}</span>`;

  let rows = `<div class="tlrow tlhead"><div class="tllabel">Project and task</div><div class="track" style="width:${width}px">${head}</div></div>`;
  let any = false;
  for (const { p, s } of groups) {
    const tasks = st.data.tasks.filter(t => t.projectId === (p ? p.id : '') && visible(t))
      .sort((a, b) => (a.start || a.due).localeCompare(b.start || b.due));
    if (!p && !tasks.length) continue;
    any = true;
    const status = p ? C.projectStatusText(p, s, st.today) : '';
    let span = '';
    if (p && s.start && s.end) {
      const b = C.barSpan(s.start, s.end);
      span = `<span class="pspan" data-state="${s.state}" style="left:${x(b.s)}px;width:${(C.diffDays(b.s, b.e) + 1) * px}px"></span>`;
    }
    if (p?.due) span += `<span class="pdue" title="Project due ${C.fmtShort(p.due)}" style="left:${x(p.due) + px / 2 - 7}px"></span>`;
    rows += `<div class="tlrow proj" data-state="${p ? s.state : 'open'}"><div class="tllabel"><button data-act="${p ? 'openproj' : 'openloose'}" data-id="${p ? p.id : ''}">${esc(p ? p.name : 'No project')}</button><span class="lst">${esc(status)}</span></div>
      <div class="track" style="width:${width}px;${grid}">${span}</div></div>`;
    for (const t of tasks) {
      const b = C.barSpan(t.start, t.due), ts = C.taskState(t, st.today);
      const single = !t.start || !t.due;
      const left = x(b.s), w = (C.diffDays(b.s, b.e) + 1) * px;
      const bar = single
        ? `<span class="tbar diamond" data-state="${ts}" data-id="${t.id}" style="left:${left + px / 2 - 8}px" title="${esc(t.title)}: ${esc(C.dueText(t, st.today))}"></span>`
        : `<span class="tbar" data-state="${ts}" data-id="${t.id}" style="left:${left}px;width:${w}px" title="${esc(t.title)}: ${C.fmtShort(b.s)} to ${C.fmtShort(b.e)}"><span class="grip"></span></span>`;
      const tagLeft = single ? left + px / 2 + 14 : left + w + 8;
      rows += `<div class="tlrow" data-state="${ts}"><div class="tllabel">${tickBtn(t)}<button data-act="edit" data-id="${t.id}">${esc(t.title)}</button><span class="lst">${esc(C.dueText(t, st.today))}</span></div>
        <div class="track" style="width:${width}px;${grid}">${bar}<span class="btag" style="left:${tagLeft}px">${ts === 'late' ? esc(C.dueText(t, st.today)) : t.due ? C.fmtShort(t.due) : ''}</span></div></div>`;
    }
  }
  if (!any) rows += `<p class="empty">Nothing on the timeline yet. Give a task a start or due date and it shows up here.</p>`;
  const wrap = $('#tlWrap'), keep = wrap.scrollLeft, first = !wrap.dataset.ready;
  $('#tl').style.width = 280 + width + 'px';
  $('#tl').innerHTML = rows + `<div class="nowline" style="left:${280 + x(st.today) + px / 2}px" title="Today"></div>`;
  const placed = st.data.tasks.filter(visible).length;
  $('#tlSub').textContent = `${C.plural(placed, 'task')} placed${undated ? `, ${undated} without dates not shown` : ''}. Weeks run Sunday to Saturday.`;
  if (first || wrap.dataset.zoom !== z) { wrap.dataset.ready = 1; wrap.dataset.zoom = z; jumpToday(); }
  else wrap.scrollLeft = keep;
}
function jumpToday() {
  if (!tlRange) return;
  const wrap = $('#tlWrap');
  wrap.scrollLeft = Math.max(0, C.diffDays(tlRange.from, st.today) * tlRange.px - (wrap.clientWidth - 280) / 3);
}

// Drag bars: move, or resize from the right edge. A click without movement opens the editor.
function initTimelineDrag() {
  let drag = null;
  $('#tl').addEventListener('pointerdown', e => {
    const bar = e.target.closest('.tbar');
    if (!bar || e.button !== 0) return;
    const t = task(bar.dataset.id);
    drag = { bar, t, x0: e.clientX, mode: e.target.classList.contains('grip') ? 'resize' : 'move',
      left: parseFloat(bar.style.left), width: bar.offsetWidth, delta: 0 };
    bar.setPointerCapture(e.pointerId);
    bar.classList.add('dragging');
    e.preventDefault();
  });
  $('#tl').addEventListener('pointermove', e => {
    if (!drag) return;
    const px = tlRange.px;
    drag.delta = Math.round((e.clientX - drag.x0) / px);
    if (drag.mode === 'move') drag.bar.style.left = drag.left + drag.delta * px + 'px';
    else drag.bar.style.width = Math.max(px, drag.width + drag.delta * px) + 'px';
  });
  const end = () => {
    if (!drag) return;
    const { t, delta, mode, bar } = drag; drag = null;
    bar.classList.remove('dragging');
    if (!delta) { openTask(t.id); return; }
    const undo = snapshotFor();
    if (mode === 'move') {
      if (t.start) t.start = C.addDays(t.start, delta);
      if (t.due) t.due = C.addDays(t.due, delta);
    } else {
      const nd = C.addDays(t.due, delta);
      t.due = nd < t.start ? t.start : nd;
    }
    commit(`${t.title}: ${t.start ? C.fmtShort(t.start) + ' to ' : 'due '}${C.fmtShort(t.due || t.start)}`, undo);
  };
  $('#tl').addEventListener('pointerup', end);
  $('#tl').addEventListener('pointercancel', end);
}

// ----- Workbook -----
const XLSX_TYPE = { description: 'Excel workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } };
function renderWorkbook() {
  const m = st.data.meta;
  $('#bookSub').textContent = st.handle
    ? 'Your projects live in this workbook. Every PC pointed at it sees the same data.'
    : 'Keep your projects in a workbook in OneDrive so every PC with the app sees the same data.';
  let b = '';
  if (S.DEMO) {
    b = `<h2>Workbook</h2><p class="muted">OneDrive sync is off in demo mode, so demo data never reaches your real workbook. You can still download a copy.</p>
      <div class="btnrow"><button class="btn ghost" data-act="wbcopy">Download a copy</button></div>`;
  } else if (!FSA) {
    b = `<h2>Workbook</h2><p class="muted">This browser can't open files directly, so it can't keep a workbook in sync. Open the app in Edge or Chrome on a PC. Here you can still download a copy or load one.</p>
      <div class="btnrow"><button class="btn ghost" data-act="wbcopy">Download a copy</button><button class="btn ghost" data-act="wbload">Load from workbook</button></div>`;
  } else if (st.handle) {
    const status = {
      needperm: '<span class="hi">Not connected. Reconnect to load the latest.</span>',
      locked: '<span class="hi">Open in Excel. Changes save once you close it.</span>',
      error: '<span class="hi">Couldn\'t reach the file.</span>'
    }[st.sync] || (m.dirty ? 'Saving...' : m.lastSync ? `<span class="ok">Synced ${relTime(new Date(m.lastSync).toISOString())}</span>` : 'Connecting...');
    b = `<h2>Synced workbook</h2><p class="big">${esc(st.handle.name)}</p><p style="margin:0">${status}</p>
      <div class="btnrow">
        ${st.sync === 'needperm' ? '<button class="btn primary" data-act="reconnect">Reconnect workbook</button>' : '<button class="btn ghost" data-act="syncnow">Sync now</button>'}
        <button class="btn ghost" data-act="wbcopy">Download a copy</button>
        <button class="btn mute" data-act="wbopen">Switch workbook</button>
        <button class="btn mute" data-act="wbunlink">Stop syncing</button>
      </div>
      <div class="section"><div class="sh"><h2>Claude inbox</h2><span>${st.inbox ? 'Connected' : 'Not connected'}</span></div>
        ${st.inbox
          ? `<p class="notes">${esc(st.inbox.name)}. ${m.inboxLast ? `Last applied ${C.plural(m.inboxLast.n, 'change')} ${relTime(new Date(m.inboxLast.at).toISOString())}.` : 'Nothing from Claude yet.'}</p>
             <div class="btnrow"><button class="btn sm mute" data-act="inboxunlink">Disconnect</button></div>`
          : `<p class="notes">Lets Claude add and update tasks from chat or voice on any device, including your phone. Claude writes changes to a text file next to the workbook; the app applies them and clears it.</p>
             <div class="btnrow"><button class="btn ghost" data-act="inboxlink">Connect the inbox</button></div>
             <p class="hint">Save it in the same OneDrive folder as the workbook, named Projects App Inbox.txt.</p>`}
      </div>
      <div class="section"><div class="sh"><h2>On your other devices</h2></div>
        <p class="notes">PC: open the app, go to Workbook, choose <b>Open my workbook</b> and pick ${esc(st.handle.name)} in your OneDrive folder.
Phone or tablet: open ${esc(st.handle.name)} in the Excel or OneDrive app and edit the Tasks sheet. Your PCs pick up the changes when the app opens.
Close the file in Excel on a PC when you're done, or the app can't save to it.</p></div>`;
  } else {
    const has = st.data.tasks.length || st.data.projects.length;
    b = `<h2>Not synced yet</h2>
      <p class="muted">Your data is only on this PC. In the file picker, choose your OneDrive folder (often <b>OneDrive - Mindful Health Solutions</b>).</p>
      ${has ? `<div class="callout">${C.plural(st.data.tasks.length, 'task')} live only in this browser. If its data is cleared, they're gone.</div>` : ''}
      <div class="btnrow">
        ${has ? '<button class="btn primary" data-act="wbnew">Create a workbook in OneDrive</button><button class="btn ghost" data-act="wbopen">Open my workbook</button>'
              : '<button class="btn primary" data-act="wbopen">Open my workbook</button><button class="btn ghost" data-act="wbnew">Create a workbook in OneDrive</button>'}
      </div>
      <p class="hint">Open my workbook: you already made one on another PC. The app shows what's in it. Create: saves what's here to a new file.</p>`;
  }
  if (S.hasSnapshot()) b += `<p class="hint"><button class="btn sm ghost" data-act="undoload">Put back my earlier data</button> Restores what this PC had before you opened a workbook, and stops syncing.</p>`;
  $('#bookCard').innerHTML = b;

  $('#appCard').innerHTML = `<h2>App</h2>
    ${opt('autoFullscreen', 'Go full screen when the app opens', 'Installed app only. Browsers need one click or key press first, so it switches on your first click.')}
    <div class="optrow"><span>Install as a desktop app<small>${STANDALONE ? 'Running as an installed app.' : st.installPrompt ? 'Adds it to the Start menu and opens in its own window.' : 'In Edge: the install icon in the address bar, or ... > Apps > Install this site as an app.'}</small></span>
      ${!STANDALONE && st.installPrompt ? '<button class="btn sm ghost" data-act="install">Install</button>' : ''}</div>`;

  const P = st.data.projects.length, T = st.data.tasks.length, open = st.data.tasks.filter(t => t.status !== 'Done').length;
  $('#storeCard').innerHTML = `<h2>${st.handle ? 'In the workbook' : 'Stored on this PC'}</h2>
    <dl class="facts"><dt>Projects</dt><dd>${P}</dd><dt>Tasks</dt><dd>${T} (${open} open)</dd>
    ${st.handle ? '' : `<dt>Protection</dt><dd>${st.persisted === 'persistent' ? '<span class="ok">Kept unless you clear it</span>' : '<span class="hi">May be cleared</span>'}</dd>`}</dl>
    <p class="hint">${st.handle ? 'This PC also keeps a working copy, so the app opens instantly and works offline. Changes made offline save to the workbook when OneDrive is back.' : 'Clearing browser data for this site deletes everything here.'}</p>`;
}
function opt(key, label, help, disabled) {
  return `<div class="optrow"><span>${label}<small>${help}</small></span>
    <label class="switch"><input type="checkbox" data-opt="${key}" ${st.settings[key] && !disabled ? 'checked' : ''} ${disabled ? 'disabled' : ''} aria-label="${esc(label)}"><span></span></label></div>`;
}

// ---------- OneDrive workbook sync ----------
// The workbook is the shared copy. This PC keeps a working copy in localStorage.
// meta.fileModified: the file's lastModified when we last read or wrote it.
// meta.lastSync: when this PC last matched the file. meta.dirty: local edits not yet written.
let syncing = false, syncAgain = false;
function setSync(s) { st.sync = s; if (st.view === 'workbook') render(); else { renderConn(); $('#cBook').hidden = !needsBackup(); } }

async function ensurePerm(h, ask) {
  const o = { mode: 'readwrite' };
  if ((await h.queryPermission(o)) === 'granted') return true;
  return ask && (await h.requestPermission(o)) === 'granted';
}
async function makeBuffer() {
  if (!window.ExcelJS) throw new Error('ExcelJS not loaded');
  return buildWorkbook(st.data, st.today);
}

// File wins, except records edited here since the last sync. Never synced: everything here wins.
function merge(file, local, lastSync) {
  const pick = (fromFile, fromLocal) => {
    const out = new Map(fromFile.map(r => [r.id, r]));
    for (const r of fromLocal) if (!lastSync || (r.updated || 0) > lastSync) out.set(r.id, r);
    return [...out.values()];
  };
  return S.normalise({ projects: pick(file.projects, local.projects), tasks: pick(file.tasks, local.tasks), items: pick(file.items || [], local.items || []), digest: pick(file.digest || [], local.digest || []), meta: local.meta });
}

async function syncNow({ ask = false } = {}) {
  if (!st.handle || S.DEMO) return;
  if (syncing) { syncAgain = true; return; }
  syncing = true;
  $('#syncBtn').classList.add('spin');
  let phase = 'read';
  try {
    if (!(await ensurePerm(st.handle, ask))) { setSync('needperm'); return; }
    const meta = st.data.meta;
    const file = await st.handle.getFile();
    let write = !!meta.dirty;
    if (!file.size) write = true;
    else if (file.lastModified !== meta.fileModified) {
      const got = await readWorkbook(await file.arrayBuffer());
      const before = st.data.tasks.length;
      st.data = meta.dirty ? merge(got, st.data, meta.lastSync) : S.normalise({ ...got, meta });
      st.data.meta = meta;
      meta.fileModified = file.lastModified;
      primeStamps();
      if (!meta.dirty) { meta.lastSync = Date.now(); S.save(st.data); render(); }
      else toast('Merged changes from another device with yours');
      if (meta.lastSync && before !== st.data.tasks.length && !meta.dirty) console.info('loaded from workbook');
    }
    if (await processInbox(ask)) write = true;
    if (write) {
      const buf = await makeBuffer();
      phase = 'write';
      const w = await st.handle.createWritable();
      await w.write(buf);
      await w.close();
      meta.fileModified = (await st.handle.getFile()).lastModified;
      meta.dirty = false;
      meta.lastSync = Date.now();
      meta.lastWorkbookSave = new Date().toISOString();
    } else if (!meta.dirty) meta.lastSync = Date.now();
    S.save(st.data);
    setSync('ok');
    render();
  } catch (e) {
    console.error('workbook sync failed', e);
    // Reading worked but writing didn't: almost always Excel holding the file open.
    setSync(phase === 'write' && e.name !== 'NotFoundError' ? 'locked' : 'error');
  } finally {
    syncing = false;
    $('#syncBtn').classList.remove('spin');
    if (syncAgain) { syncAgain = false; syncNow(); }
  }
}

async function linkHandle(h, { fileWins }) {
  if (!(await ensurePerm(h, true))) { toast('The app needs permission to edit that file. Try again and choose Allow.'); return; }
  st.handle = h; await S.setHandle(h);
  const m = st.data.meta;
  m.fileModified = -1;
  m.dirty = !fileWins;          // create: write what's here; open: read the file
  m.lastSync = fileWins ? m.lastSync : 0;
  st.sync = 'ok';
  await syncNow();
}

async function createWorkbook() {
  try {
    const h = await window.showSaveFilePicker({ suggestedName: 'Projects.xlsx', startIn: 'documents', types: [XLSX_TYPE] });
    await linkHandle(h, { fileWins: false });
    if (st.sync === 'ok') toast(`Saved to ${h.name}. Changes save there from now on.`);
  } catch (e) { if (e.name !== 'AbortError') { console.error(e); toast("Couldn't create the workbook there. Try another folder."); } }
}

async function openWorkbook() {
  let h;
  try { [h] = await window.showOpenFilePicker({ startIn: 'documents', types: [XLSX_TYPE] }); }
  catch (e) { if (e.name !== 'AbortError') console.error(e); return; }
  try {
    const got = await readWorkbook(await (await h.getFile()).arrayBuffer());
    const here = st.data.tasks.length;
    if (here && !st.handle && !confirm(`Show what's in ${h.name}? It has ${C.plural(got.tasks.length, 'task')}. The ${C.plural(here, 'task')} on this PC are set aside; you can put them back from the Workbook screen.`)) return;
    if (here && !st.handle) S.snapshot(st.data);
    st.selTask = null; st.selProject = null;
    await linkHandle(h, { fileWins: true });
    if (st.sync === 'ok') toast(`Opened ${h.name}. ${C.plural(st.data.tasks.length, 'task')}.`);
  } catch (e) {
    console.error('workbook open failed', e);
    toast("Couldn't read that workbook. It needs a Tasks sheet with a header row. Nothing was changed.");
  }
}

async function unlink() {
  st.handle = null; st.sync = 'ok';
  Object.assign(st.data.meta, { dirty: false, fileModified: null, lastSync: 0 });
  await S.setHandle(null); S.save(st.data); render();
  toast('Stopped syncing. The workbook stays in OneDrive; this PC keeps its copy.');
}

async function downloadCopy() {
  try {
    const buf = await makeBuffer();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    a.download = `Projects ${st.today}.xlsx`; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  } catch (e) { console.error(e); toast("Couldn't build the workbook. Reload the app and try again."); }
}

// Fallback for browsers without file access: one-off load that replaces local data.
async function loadWorkbook() {
  const file = await new Promise(res => { const i = $('#filePick'); i.value = ''; i.onchange = () => res(i.files[0]); i.click(); });
  if (!file) return;
  try {
    const got = await readWorkbook(await file.arrayBuffer());
    if (!confirm(`Replace the data here with ${file.name}? It has ${C.plural(got.tasks.length, 'task')}. You can undo this once.`)) return;
    S.snapshot(st.data);
    st.data = S.normalise({ ...got, meta: st.data.meta });
    commit(`Loaded ${C.plural(got.tasks.length, 'task')} from ${file.name}`);
  } catch (e) {
    console.error('workbook load failed', e);
    toast("Couldn't read that workbook. It needs a Tasks sheet with a header row. Nothing was changed.");
  }
}

// ---------- Claude inbox (voice and chat) ----------
// Claude appends one JSON change per line to a text file next to the workbook; the app applies and clears it.
const INBOX_HEADER = '# Projects App inbox. Claude adds one JSON change per line. The app applies them and clears this file.\n';
function inboxProblem(why, raw) {
  return { id: C.uid(), projectId: '', title: `Check inbox item: ${why}`.slice(0, 140), status: 'To do', priority: 'Normal',
    start: '', due: st.today, doneOn: '', remind: '', reminded: '', notes: String(raw).slice(0, 2000), updated: Date.now() };
}
async function processInbox(ask = false) {
  if (!st.inbox) return false;
  try {
    if (!(await ensurePerm(st.inbox, ask))) return false;
    const text = await (await st.inbox.getFile()).text();
    const { ops, bad } = parseInbox(text);
    if (!ops.length && !bad.length) return false;
    const meta = st.data.meta;
    const seen = new Set(meta.appliedOps || []);
    const fresh = ops.filter(o => !o.oid || !seen.has(o.oid));
    const res = applyOps(st.data, fresh, st.today);
    for (const f of res.failed) st.data.tasks.push(inboxProblem(f.why, JSON.stringify(f.op)));
    for (const b of bad) st.data.tasks.push(inboxProblem('line is not valid JSON', b));
    meta.appliedOps = [...seen, ...fresh.map(o => o.oid).filter(Boolean)].slice(-300);
    // Clear what we applied. Keep anything Claude appended while we were working.
    const again = await (await st.inbox.getFile()).text();
    const rest = again.startsWith(text) ? again.slice(text.length) : again === text ? '' : null;
    if (rest !== null) { const w = await st.inbox.createWritable(); await w.write(INBOX_HEADER + rest); await w.close(); }
    meta.dirty = true;
    meta.inboxLast = { at: Date.now(), n: res.applied.length, failed: res.failed.length + bad.length };
    primeStamps();
    S.save(st.data);
    const fails = res.failed.length + bad.length;
    toast(`Applied ${C.plural(res.applied.length, 'change')} from Claude${fails ? `. ${fails} couldn't be applied; see Today.` : ''}`);
    return true;
  } catch (e) { console.error('inbox failed', e); return false; }
}
async function linkInbox() {
  try {
    const h = await window.showSaveFilePicker({ suggestedName: 'Projects App Inbox.txt', startIn: st.handle || 'documents',
      types: [{ description: 'Text file', accept: { 'text/plain': ['.txt'] } }] });
    const f = await h.getFile();
    if (!f.size) { const w = await h.createWritable(); await w.write(INBOX_HEADER); await w.close(); }
    st.inbox = h; await S.setHandle(h, 'inbox');
    toast(`Inbox connected: ${h.name}`);
    await syncNow({ ask: true });
    render();
  } catch (e) { if (e.name !== 'AbortError') { console.error(e); toast("Couldn't use that file for the inbox. Try another name or folder."); } }
}

// ---------- reminders ----------
function askNotify() { if ('Notification' in window && Notification.permission === 'default') Notification.requestPermission(); }
function checkReminders() {
  const due = dueReminders(st.data.tasks, stampText());
  if (!due.length) return;
  for (const t of due) {
    t.reminded = 'Yes';
    if (t.snooze) t.snooze = '';
    if ('Notification' in window && Notification.permission === 'granted') {
      try {
        const n = new Notification(t.title, { body: `${pname(t.projectId)} · ${C.dueText(t, st.today)}`, tag: t.id, icon: 'icons/icon-192.png', requireInteraction: true });
        n.onclick = () => { window.focus(); st.selTask = t.id; go('today'); n.close(); };
      } catch (e) { console.warn('notification', e); }
    }
  }
  commit(due.length === 1 ? `Reminder: ${due[0].title}` : `${due.length} reminders: ${due.map(t => t.title).join('; ')}`);
}
setInterval(checkReminders, 20000);

// ---------- theme ----------
const darkMQ = matchMedia('(prefers-color-scheme: dark)');
function applyTheme() {
  const th = st.settings.theme || 'system', root = document.documentElement;
  if (th === 'system') root.removeAttribute('data-theme'); else root.dataset.theme = th;
  const dark = th === 'dark' || (th === 'system' && darkMQ.matches);
  root.toggleAttribute('data-dark', dark);
  const b = $('#themeBtn');
  b.setAttribute('aria-label', dark ? 'Switch to light mode' : 'Switch to dark mode');
  b.title = dark ? 'Light mode' : 'Dark mode';
}
darkMQ.addEventListener('change', applyTheme);
$('#themeBtn').onclick = () => {
  st.settings.theme = document.documentElement.hasAttribute('data-dark') ? 'light' : 'dark';
  S.saveSettings(st.settings); applyTheme();
};
applyTheme();

$('#syncBtn').onclick = async () => {
  if (!st.handle) { go('workbook'); toast('Connect a workbook first, then this keeps it in sync.'); return; }
  await syncNow({ ask: true });
  if (st.sync === 'ok') toast(`Up to date with ${st.handle.name}`);
};

// ---------- dialogs ----------
const dlg = $('#dlg'), form = $('#dlgForm');
function field(label, html) { return `<div class="field"><label>${label}</label>${html}</div>`; }
function options(list, val) { return list.map(o => `<option${o === val ? ' selected' : ''}>${esc(o)}</option>`).join(''); }

function openTask(id, preset = {}) {
  const t = id ? task(id) : { title: '', projectId: preset.projectId ?? (st.view === 'projects' && st.selProject ? st.selProject : ''), status: 'To do', priority: 'Normal', start: '', due: preset.due ?? st.today, notes: '' };
  const projOpts = `<option value="">No project</option>` + st.data.projects.filter(p => p.status !== 'Done' || p.id === t.projectId)
    .map(p => `<option value="${p.id}"${p.id === t.projectId ? ' selected' : ''}>${esc(p.name)}</option>`).join('') + `<option value="__new">New project...</option>`;
  form.innerHTML = `<div class="mh"><h2>${id ? 'Edit task' : 'New task'}</h2><button class="x" value="cancel" aria-label="Close">×</button></div>
    ${field('Task', `<input name="title" required maxlength="200" value="${esc(t.title)}">`)}
    ${field('Project', `<select name="projectId">${projOpts}</select>`)}
    <div class="field" id="newProjField" hidden><label>New project name</label><input name="newProject" maxlength="120"></div>
    <div class="row2">${field('Status', `<select name="status">${options(C.STATUSES, t.status)}</select>`)}${field('Priority', `<select name="priority">${options(C.PRIORITIES, t.priority)}</select>`)}</div>
    <div class="row2">${field('Start (optional)', `<input type="date" name="start" value="${t.start}">`)}${field('Due', `<input type="date" name="due" value="${t.due}">`)}</div>
    ${field('Remind me (optional)', `<input type="datetime-local" name="remind" value="${(t.remind || '').replace(' ', 'T')}">`)}
    ${field('Notes', `<textarea name="notes" maxlength="8000">${esc(t.notes)}</textarea>`)}
    <div class="err" id="dlgErr"></div>
    <div class="mfoot">${id ? `<button class="btn sm danger" value="delete" type="submit" formnovalidate>Delete task</button>` : '<span></span>'}
      <span><button class="btn mute" value="cancel" formnovalidate>Cancel</button> <button class="btn primary" value="save">${id ? 'Save task' : 'Add task'}</button></span></div>`;
  const sel = form.projectId;
  sel.onchange = () => { $('#newProjField').hidden = sel.value !== '__new'; if (sel.value === '__new') form.newProject.focus(); };
  form.onsubmit = e => {
    const action = e.submitter?.value;
    if (action === 'cancel') return;
    if (action === 'delete') { removeTask(id); return; }
    const f = Object.fromEntries(new FormData(form));
    if (f.start && f.due && f.start > f.due) { e.preventDefault(); $('#dlgErr').textContent = 'Start is after the due date. Change one of them.'; return; }
    if (f.projectId === '__new' && !f.newProject.trim()) { e.preventDefault(); $('#dlgErr').textContent = 'Name the new project, or pick an existing one.'; return; }
    const undo = snapshotFor();
    let projectId = f.projectId;
    if (projectId === '__new') { const p = { id: C.uid(), name: f.newProject.trim(), status: 'Active', start: '', due: '', notes: '' }; st.data.projects.push(p); projectId = p.id; }
    const target = id ? task(id) : { id: C.uid(), doneOn: '', remind: '', reminded: '' };
    const wasDone = target.status === 'Done';
    const rem = normRemind(f.remind);
    if (rem !== (target.remind || '')) { target.remind = rem; target.reminded = ''; if (rem) askNotify(); }
    Object.assign(target, { title: f.title.trim(), projectId, priority: f.priority, start: f.start, due: f.due, notes: f.notes.trim() });
    if (target.status !== f.status || !id) C.setStatus(target, f.status, st.today);
    if (wasDone && f.status === 'Done') target.doneOn = target.doneOn || st.today;
    if (!id) { st.data.tasks.push(target); st.selTask = target.id; }
    commit(id ? 'Task saved' : `Added ${target.title}`, undo);
  };
  dlg.showModal();
  form.title.focus();
}

function openProject(id) {
  const p = id ? proj(id) : { name: '', status: 'Active', start: '', due: '', notes: '' };
  form.innerHTML = `<div class="mh"><h2>${id ? 'Edit project' : 'New project'}</h2><button class="x" value="cancel" aria-label="Close">×</button></div>
    ${field('Project', `<input name="name" required maxlength="120" value="${esc(p.name)}">`)}
    ${field('Status', `<select name="status">${options(C.PROJECT_STATUSES, p.status)}</select>`)}
    <div class="row2">${field('Start (optional)', `<input type="date" name="start" value="${p.start}">`)}${field('Due (optional)', `<input type="date" name="due" value="${p.due}">`)}</div>
    <div class="row2">${field('Reference file (optional)', `<input name="refName" maxlength="120" placeholder="Build sheet" value="${esc(p.refName || '')}">`)}${field('Link or path', `<input name="refUrl" maxlength="600" placeholder="https://...sharepoint.com/..." value="${esc(p.refUrl || '')}">`)}</div>
    ${field('Notes', `<textarea name="notes" maxlength="4000">${esc(p.notes)}</textarea>`)}
    <div class="err" id="dlgErr"></div>
    <div class="mfoot">${id ? `<button class="btn sm danger" value="delete" formnovalidate>Delete project</button>` : '<span></span>'}
      <span><button class="btn mute" value="cancel" formnovalidate>Cancel</button> <button class="btn primary" value="save">${id ? 'Save project' : 'Add project'}</button></span></div>`;
  form.onsubmit = e => {
    const action = e.submitter?.value;
    if (action === 'cancel') return;
    if (action === 'delete') { removeProject(id); return; }
    const f = Object.fromEntries(new FormData(form));
    if (f.start && f.due && f.start > f.due) { e.preventDefault(); $('#dlgErr').textContent = 'Start is after the due date. Change one of them.'; return; }
    const undo = snapshotFor();
    const target = id ? p : { id: C.uid() };
    Object.assign(target, { name: f.name.trim(), status: f.status, start: f.start, due: f.due, notes: f.notes.trim(), refName: f.refName.trim(), refUrl: f.refUrl.trim() });
    if (!id) { st.data.projects.push(target); st.selProject = target.id; if (st.projFilter !== 'all' && target.status !== 'Active') st.projFilter = 'all'; }
    commit(id ? 'Project saved' : `Added ${target.name}`, undo);
  };
  dlg.showModal();
  form.name.focus();
}

function openSnooze(id) {
  const t = task(id), now = new Date();
  const at = (h, addDays = 0) => { const d = new Date(); d.setDate(d.getDate() + addDays); d.setHours(h, 0, 0, 0); return d; };
  const plus = hrs => new Date(now.getTime() + hrs * 3600000);
  const mon = at(9, ((8 - now.getDay()) % 7) || 7);
  const opts = [['In 1 hour', plus(1)], ['In 3 hours', plus(3)], ['Tomorrow 9:00', at(9, 1)], [`Monday 9:00`, mon]];
  form.innerHTML = `<div class="mh"><h2>Snooze ${esc(t.title)}</h2><button class="x" value="cancel" aria-label="Close">×</button></div>
    <p class="muted" style="margin:-6px 0 14px">It leaves Today and comes back with a reminder.</p>
    <div class="snz">${opts.map(([l, d], i) => `<button class="btn ghost" value="opt${i}">${l}<small>${C.fmtDay(C.iso(d))}, ${d.getHours()}:${C.pad(d.getMinutes())}</small></button>`).join('')}</div>
    ${field('Or pick a date and time', `<input type="datetime-local" name="until" value="${stampText(at(9, 1)).replace(' ', 'T')}">`)}
    <div class="mfoot"><span></span><span><button class="btn mute" value="cancel" formnovalidate>Cancel</button> <button class="btn primary" value="pick">Snooze until then</button></span></div>`;
  form.onsubmit = e => {
    const v = e.submitter?.value;
    if (v === 'cancel') return;
    const until = v === 'pick' ? normRemind(form.until.value) : stampText(opts[Number(v.slice(3))][1]);
    if (!until) { e.preventDefault(); return; }
    const undo = snapshotFor();
    t.snooze = until; t.remind = until; t.reminded = '';
    askNotify();
    commit(`Snoozed until ${remindText(until)}`, undo);
  };
  dlg.showModal();
}

function removeTask(id) {
  const t = task(id); if (!t) return;
  const undo = snapshotFor();
  st.data.tasks = st.data.tasks.filter(x => x.id !== id);
  if (st.selTask === id) st.selTask = null;
  commit(`Deleted ${t.title}`, undo);
}
function removeProject(id) {
  const p = proj(id); if (!p) return;
  const n = st.data.tasks.filter(t => t.projectId === id).length;
  if (n && !confirm(`Delete ${p.name}? Its ${C.plural(n, 'task')} move to No project.`)) return;
  const undo = snapshotFor();
  st.data.tasks.forEach(t => { if (t.projectId === id) t.projectId = ''; });
  st.data.items = st.data.items.filter(i => i.projectId !== id);
  st.data.projects = st.data.projects.filter(x => x.id !== id);
  st.selProject = null;
  commit(`Deleted ${p.name}`, undo);
}

// ---------- events ----------
function go(view) { st.view = view; render(); $('#page').scrollTop = 0; }

document.addEventListener('click', e => {
  const el = e.target.closest('[data-act],[data-view],[data-f],[data-z]');
  if (!el) return;
  if (el.dataset.view && el.classList.contains('item')) return go(el.dataset.view);
  if (el.dataset.f) { st.projFilter = el.dataset.f; st.selProject = null; return render(); }
  if (el.dataset.z) { st.settings.timelineZoom = el.dataset.z; S.saveSettings(st.settings); return render(); }
  const id = el.dataset.id, t = id ? task(id) : null;
  switch (el.dataset.act) {
    case 'select': st.selTask = id; render(); break;
    case 'toggle': {
      const undo = snapshotFor();
      C.setStatus(t, t.status === 'Done' ? 'To do' : 'Done', st.today);
      commit(t.status === 'Done' ? `Done: ${t.title}` : `Reopened ${t.title}`, undo); break;
    }
    case 'advance': { const undo = snapshotFor(); C.setStatus(t, el.dataset.status, st.today); commit(t.status === 'Done' ? `Done: ${t.title}` : t.status === 'Doing' ? `Started ${t.title}` : `Reopened ${t.title}`, undo); break; }
    case 'tomorrow': { const undo = snapshotFor(); const nd = C.addDays(st.today, 1); if (t.start && t.start > nd) t.start = ''; t.due = nd; commit(`${t.title} moved to tomorrow`, undo); break; }
    case 'week': { const undo = snapshotFor(); const base = t.due && t.due > st.today ? t.due : st.today; if (t.start) t.start = C.addDays(t.start, C.diffDays(base, C.addDays(base, 7))); t.due = C.addDays(base, 7); commit(`${t.title} due ${C.fmtDay(t.due)}`, undo); break; }
    case 'delete': removeTask(id); break;
    case 'edit': openTask(id); break;
    case 'more': st.showAllLater = true; render(); break;
    case 'openproj': st.selProject = id; st.projFilter = proj(id)?.status === 'Active' ? 'active' : 'all'; go('projects'); break;
    case 'openloose': st.selProject = ''; st.projFilter = 'active'; go('projects'); break;
    case 'selproj': st.selProject = id; render(); break;
    case 'editproj': openProject(id); break;
    case 'projdone': case 'projactive': case 'projhold': {
      const p = proj(id), undo = snapshotFor();
      p.status = { projdone: 'Done', projactive: 'Active', projhold: 'On hold' }[el.dataset.act];
      commit(`${p.name}: ${p.status.toLowerCase()}`, undo); break;
    }
    case 'delproj': removeProject(id); break;
    case 'addday': openTask(null, { due: el.dataset.day }); break;
    case 'wbnew': createWorkbook(); break;
    case 'wbopen': openWorkbook(); break;
    case 'wbload': loadWorkbook(); break;
    case 'wbcopy': downloadCopy(); break;
    case 'wbunlink': unlink(); break;
    case 'syncnow': syncNow({ ask: true }); break;
    case 'reconnect': syncNow({ ask: true }); break;
    case 'undoload': { const d = S.restoreSnapshot(); if (d) { st.handle = null; S.setHandle(null); st.data = d; Object.assign(st.data.meta, { dirty: false, fileModified: null, lastSync: 0 }); primeStamps(); S.save(st.data); render(); toast('Put back your earlier data. Syncing is off.'); } break; }
    case 'toggleitem': { const it = item(id), undo = snapshotFor(); it.done = it.done ? '' : 'Yes'; it.doneOn = it.done ? st.today : ''; commit(it.done ? `Checked off ${it.title}` : `Unchecked ${it.title}`, undo); break; }
    case 'delitem': { const it = item(id), undo = snapshotFor(); st.data.items = st.data.items.filter(i => i !== it); commit(`Removed ${it.title} from the checklist`, undo); break; }
    case 'itemtask': {
      const it = item(id), undo = snapshotFor();
      const t = { id: C.uid(), projectId: it.projectId, title: it.title, status: 'To do', priority: 'Normal', start: '', due: '', doneOn: '', remind: '', reminded: '', notes: it.notes };
      st.data.tasks.push(t); st.data.items = st.data.items.filter(i => i !== it); st.selTask = t.id;
      commit(`${t.title} is now one of your tasks`, undo); break;
    }
    case 'snooze': openSnooze(id); break;
    case 'unsnooze': { const undo = snapshotFor(); t.snooze = ''; t.remind = ''; t.reminded = ''; commit(`${t.title} is back`, undo); break; }
    case 'copyref': { const p = proj(id); navigator.clipboard?.writeText(p.refUrl).then(() => toast('Path copied. Paste it into File Explorer.')); break; }
    case 'dgadd': case 'dgdismiss': {
      const g = st.data.digest.find(x => x.id === id), undo = snapshotFor();
      if (el.dataset.act === 'dgdismiss') { g.status = 'dismissed'; commit('Dismissed', undo); break; }
      const pm = g.project ? st.data.projects.find(p => p.name.toLowerCase() === g.project.toLowerCase()) : null;
      const nt = { id: C.uid(), projectId: pm ? pm.id : '', title: g.kind === 'meeting' ? `Prep: ${g.title}` : g.title, status: 'To do', priority: 'Normal',
        start: '', due: g.due || st.today, doneOn: '', remind: '', reminded: '', snooze: '', notes: [g.source, g.notes].filter(Boolean).join('\n') };
      if (g.kind === 'meeting' && /^\d{1,2}:\d{2}$/.test(g.time)) {
        const [h, m] = g.time.split(':').map(Number), d = new Date(); d.setHours(h, m - 30, 0, 0);
        if (d > new Date()) nt.remind = stampText(d);
      }
      if (pm && g.source === 'Suggested by Claude') st.data.items = st.data.items.filter(i => !(i.projectId === pm.id && i.title.toLowerCase() === g.title.toLowerCase() && !i.done));
      st.data.tasks.push(nt); g.status = 'added'; st.selTask = nt.id;
      commit(`Added ${nt.title} to your day`, undo); break;
    }
    case 'dgdismissall': {
      const undo = snapshotFor(), list = st.data.digest.filter(g => g.kind === 'action' && g.status === 'new' && g.date >= C.addDays(st.today, -3));
      list.forEach(g => { g.status = 'dismissed'; });
      commit(`Dismissed ${C.plural(list.length, 'suggestion')}. They stay on the project checklists.`, undo); break;
    }
    case 'clearremind': { const undo = snapshotFor(); t.remind = ''; t.reminded = ''; commit('Reminder cleared', undo); break; }
    case 'inboxlink': linkInbox(); break;
    case 'inboxunlink': st.inbox = null; S.setHandle(null, 'inbox'); render(); toast('Inbox disconnected. The file stays in OneDrive.'); break;
    case 'install': st.installPrompt?.prompt(); st.installPrompt?.userChoice.then(() => { st.installPrompt = null; render(); }); break;
  }
});

document.addEventListener('submit', e => {
  const n = e.target.closest('form[data-act="addnote"]');
  if (n) {
    e.preventDefault();
    const text = n.text.value.trim(); if (!text) return;
    const rec = n.dataset.kind === 'task' ? task(n.dataset.id) : proj(n.dataset.id);
    if (!rec) return;
    const undo = snapshotFor();
    addNote(rec, text);
    commit('Note added', undo);
    return;
  }
  const ai = e.target.closest('form[data-act="additem"]');
  if (ai) {
    e.preventDefault();
    const title = ai.item.value.trim(); if (!title) return;
    const undo = snapshotFor();
    st.data.items.push({ id: C.uid(), projectId: ai.dataset.id, title, done: '', doneOn: '', notes: '' });
    commit(`Added ${title} to the checklist`, undo);
    requestAnimationFrame(() => document.querySelector('form[data-act="additem"] input')?.focus());
    return;
  }
  const f = e.target.closest('form[data-act="quickadd"]');
  if (!f) return;
  e.preventDefault();
  const title = f.title.value.trim(); if (!title) return;
  const undo = snapshotFor();
  st.data.tasks.push({ id: C.uid(), projectId: f.dataset.id, title, status: 'To do', priority: 'Normal', start: '', due: f.due.value, doneOn: '', remind: '', reminded: '', notes: '' });
  commit(`Added ${title}`, undo);
  requestAnimationFrame(() => document.querySelector('form[data-act="quickadd"] input[name="title"]')?.focus());
});

document.addEventListener('change', e => {
  const k = e.target.dataset?.opt;
  if (k) { st.settings[k] = e.target.checked; S.saveSettings(st.settings); render(); }
  if (e.target.id === 'tlDone') { st.settings.showDone = e.target.checked; S.saveSettings(st.settings); render(); }
});

// Week drag and drop
document.addEventListener('dragstart', e => { const c = e.target.closest?.('.chip'); if (!c) return; e.dataTransfer.setData('text/plain', c.dataset.id); c.classList.add('dragging'); });
document.addEventListener('dragend', e => e.target.classList?.remove('dragging'));
document.addEventListener('dragover', e => { const d = e.target.closest?.('.day'); if (d) { e.preventDefault(); d.classList.add('over'); } });
document.addEventListener('dragleave', e => { const d = e.target.closest?.('.day'); if (d && !d.contains(e.relatedTarget)) d.classList.remove('over'); });
document.addEventListener('drop', e => {
  const d = e.target.closest?.('.day'); if (!d) return;
  e.preventDefault();
  const t = task(e.dataTransfer.getData('text/plain')); if (!t || t.due === d.dataset.day) return render();
  const undo = snapshotFor(), nd = d.dataset.day;
  if (t.start && t.due) t.start = C.addDays(t.start, C.diffDays(t.due, nd));
  else if (t.start && t.start > nd) t.start = '';
  t.due = nd;
  commit(`${t.title} due ${C.fmtDay(nd)}`, undo);
});

$('#wkPrev').onclick = () => { st.weekOf = C.addDays(st.weekOf, -7); render(); };
$('#wkNext').onclick = () => { st.weekOf = C.addDays(st.weekOf, 7); render(); };
$('#wkNow').onclick = () => { st.weekOf = C.weekStart(st.today); render(); };
$('#newProject').onclick = () => openProject(null);
$('#railNew').onclick = () => openTask(null);
$('#tlToday').onclick = jumpToday;

document.addEventListener('keydown', e => {
  if (dlg.open || e.ctrlKey || e.metaKey || e.altKey) return;
  if (/INPUT|TEXTAREA|SELECT/.test(document.activeElement?.tagName)) return;
  const k = e.key.toLowerCase();
  if (k === 'n') { e.preventDefault(); openTask(null); }
  else if (k === 't') go('today'); else if (k === 'w') go('week'); else if (k === 'p') go('projects'); else if (k === 'l') go('timeline');
});

// ---------- full screen ----------
const fsBtn = $('#fsBtn');
const fsLabel = () => { fsBtn.textContent = document.fullscreenElement ? 'Exit full screen' : 'Full screen'; };
fsBtn.onclick = () => document.fullscreenElement ? document.exitFullscreen() : document.documentElement.requestFullscreen().catch(() => toast('Full screen is blocked here. Press F11 instead.'));
document.addEventListener('fullscreenchange', fsLabel);
if (!document.documentElement.requestFullscreen) fsBtn.hidden = true;
if (STANDALONE) {
  try { window.moveTo(0, 0); window.resizeTo(screen.availWidth, screen.availHeight); } catch {}
  if (st.settings.autoFullscreen) {
    // Browsers only allow full screen after a user gesture, so take the first one.
    const once = () => {
      removeEventListener('pointerdown', once, true); removeEventListener('keydown', once, true);
      if (st.settings.autoFullscreen && !document.fullscreenElement) document.documentElement.requestFullscreen().catch(() => {});
    };
    addEventListener('pointerdown', once, true); addEventListener('keydown', once, true);
  }
}

// ---------- start ----------
addEventListener('beforeinstallprompt', e => { e.preventDefault(); st.installPrompt = e; if (st.view === 'workbook') render(); });
addEventListener('appinstalled', () => { st.installPrompt = null; toast('Installed. Open Projects from the Start menu.'); });
addEventListener('storage', e => { if (e.key && e.key.startsWith('pp.data')) { st.data = S.load(); primeStamps(); render(); } });
// Pick up changes from other devices when the window comes back, and retry pending saves.
addEventListener('focus', () => syncNow());
document.addEventListener('visibilitychange', () => { if (!document.hidden) syncNow(); });
setInterval(() => { if (!document.hidden) syncNow(); }, 30000);

setInterval(() => {
  const t = C.todayIso();
  if (t !== st.today) { st.today = t; render(); }
  else if (st.view === 'today' && !dlg.open) { const n = new Date(); $('#clockT').textContent = `${n.getHours()}:${C.pad(n.getMinutes())}`; }
}, 30000);

(async function init() {
  if (S.DEMO) { $('#demoBanner').hidden = false; document.title = 'Projects (demo)'; }
  const q = new URLSearchParams(location.search);
  if (['today', 'week', 'projects', 'timeline', 'workbook'].includes(q.get('view'))) st.view = q.get('view');
  initTimelineDrag();
  render();
  if (q.get('new') === 'task') openTask(null);
  primeStamps();
  st.persisted = await S.persist();
  st.handle = await S.getHandle();
  st.inbox = await S.getHandle('inbox');
  checkReminders();
  if (st.handle) { st.sync = 'ok'; await syncNow(); }
  render();
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(err => console.warn('service worker', err));
})();
