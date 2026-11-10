// Change requests ("ops") from the voice/chat inbox and the Claude skill. Pure: no DOM, no storage.
// One op per object. Tasks and projects are matched by id, exact title, or a unique part of the title.
//
//   {"op":"add_project","name":"","status":"Active","start":"","due":"","notes":""}
//   {"op":"add_task","title":"","project":"","status":"To do","priority":"Normal","start":"","due":"","remind":"2026-11-12 09:00","notes":""}
//   {"op":"update_task","match":"","set":{"title":"","project":"","status":"","priority":"","start":"","due":"","remind":""}}
//   {"op":"complete","match":""}
//   {"op":"note","match":"<task>","text":""}            add a dated note to a task
//   {"op":"note","project":"<project>","text":""}       add a dated note to a project
//   {"op":"remind","match":"","at":"2026-11-12 09:00"}  set ("" clears) a reminder
//   {"op":"update_project","match":"","set":{"name":"","status":"","start":"","due":"","ref_name":"Build sheet","ref_url":"https://..."}}
//   {"op":"delete_task","match":""}
//   {"op":"snooze","match":"","until":"2026-11-11 09:00"}   hide until then, then remind ("" un-snoozes)
//   {"op":"add_suggestion","date":"2026-11-11","kind":"action|meeting","title":"","time":"10:30","project":"","due":"","source":"","notes":""}
//   {"op":"add_item","project":"","item":"","done":false,"notes":""}   project checklist item (not a personal task)
//   {"op":"check_item","match":"","project":""}  /  {"op":"uncheck_item","match":"","project":""}
import { uid, STATUSES, PRIORITIES, PROJECT_STATUSES, isDate, pad } from './core.js';

const lc = s => String(s ?? '').trim().toLowerCase();
const oneOf = (v, list) => list.find(x => lc(x) === lc(v));

export function normRemind(v) {
  if (!v) return '';
  const m = String(v).trim().match(/^(\d{4}-\d{2}-\d{2})(?:[ T](\d{1,2}):(\d{2}))?/);
  return m ? `${m[1]} ${pad(m[2] ?? 9)}:${m[3] ?? '00'}` : '';
}
export function stampText(now = new Date()) {
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}`;
}
export function addNote(rec, text, now = new Date()) {
  const line = `${stampText(now)}  ${String(text).trim()}`;
  rec.notes = rec.notes ? `${rec.notes}\n${line}` : line;
}

function find(list, key, q) {
  if (!q) return { err: 'nothing to match' };
  const byId = list.find(r => r.id === q);
  if (byId) return { rec: byId };
  const exact = list.filter(r => lc(r[key]) === lc(q));
  if (exact.length === 1) return { rec: exact[0] };
  const part = list.filter(r => lc(r[key]).includes(lc(q)));
  if (part.length === 1) return { rec: part[0] };
  if (!part.length && !exact.length) return { err: `no match for "${q}"` };
  return { err: `"${q}" matches ${part.length || exact.length}: ${(part.length ? part : exact).slice(0, 4).map(r => r[key]).join('; ')}` };
}

function projectId(data, name, now, create = true) {
  if (!name) return '';
  const f = find(data.projects, 'name', name);
  if (f.rec) return f.rec.id;
  if (!create) return null;
  const p = { id: uid(), name: String(name).trim(), status: 'Active', start: '', due: '', notes: '', updated: now };
  data.projects.push(p);
  return p.id;
}

function setTask(data, t, set, today, now) {
  if (set.title) t.title = String(set.title).trim();
  if ('project' in set) { const id = projectId(data, set.project, now); t.projectId = id || ''; }
  if (set.priority && oneOf(set.priority, PRIORITIES)) t.priority = oneOf(set.priority, PRIORITIES);
  for (const k of ['start', 'due']) if (k in set) t[k] = isDate(set[k]) ? set[k] : '';
  if ('remind' in set) { t.remind = normRemind(set.remind); t.reminded = ''; }
  if ('snooze' in set) { t.snooze = normRemind(set.snooze); if (t.snooze) { t.remind = t.snooze; t.reminded = ''; } }
  if (set.status && oneOf(set.status, STATUSES)) {
    t.status = oneOf(set.status, STATUSES);
    t.doneOn = t.status === 'Done' ? (t.doneOn || today) : '';
  }
  if (set.notes) addNote(t, set.notes);
  t.updated = now;
}

// Returns { applied: [text], failed: [{op, why}] }. Mutates data.
export function applyOps(data, ops, today, now = Date.now()) {
  const applied = [], failed = [];
  for (const o of ops) {
    try {
      switch (o.op) {
        case 'add_project': {
          if (!o.name) throw new Error('project needs a name');
          if (find(data.projects, 'name', o.name).rec && lc(find(data.projects, 'name', o.name).rec.name) === lc(o.name)) { applied.push(`project exists: ${o.name}`); break; }
          data.projects.push({ id: uid(), name: String(o.name).trim(), status: oneOf(o.status, PROJECT_STATUSES) || 'Active',
            start: isDate(o.start) ? o.start : '', due: isDate(o.due) ? o.due : '', notes: '', refName: o.ref_name || '', refUrl: o.ref_url || '', updated: now });
          if (o.notes) addNote(data.projects[data.projects.length - 1], o.notes);
          applied.push(`added project ${o.name}`); break;
        }
        case 'add_task': {
          if (!o.title) throw new Error('task needs a title');
          const pid = o.project ? projectId(data, o.project, now, false) : '';
          if (pid !== null && data.tasks.some(x => x.status !== 'Done' && x.projectId === pid && lc(x.title) === lc(o.title))) { applied.push(`already on the list: ${o.title}`); break; }
          const t = { id: uid(), projectId: '', title: '', status: 'To do', priority: 'Normal', start: '', due: '', doneOn: '', remind: '', reminded: '', snooze: '', notes: '', updated: now };
          setTask(data, t, { ...o, project: o.project || '' }, today, now);
          if (!o.project) t.projectId = '';
          data.tasks.push(t);
          applied.push(`added task ${t.title}`); break;
        }
        case 'update_task': case 'complete': case 'remind': case 'delete_task': {
          const f = find(data.tasks, 'title', o.match);
          if (!f.rec) throw new Error(f.err);
          if (o.op === 'delete_task') { data.tasks = data.tasks.filter(t => t !== f.rec); applied.push(`deleted ${f.rec.title}`); break; }
          const set = o.op === 'complete' ? { status: 'Done' } : o.op === 'remind' ? { remind: o.at || '' } : (o.set || {});
          setTask(data, f.rec, set, today, now);
          applied.push(`${o.op === 'complete' ? 'done' : 'updated'}: ${f.rec.title}`); break;
        }
        case 'note': {
          if (!o.text) throw new Error('note needs text');
          const f = o.project ? find(data.projects, 'name', o.project) : find(data.tasks, 'title', o.match);
          if (!f.rec) throw new Error(f.err);
          addNote(f.rec, o.text); f.rec.updated = now;
          applied.push(`note on ${f.rec.title || f.rec.name}`); break;
        }
        case 'update_project': {
          const f = find(data.projects, 'name', o.match);
          if (!f.rec) throw new Error(f.err);
          const s = o.set || {}, p = f.rec;
          if (s.name) p.name = String(s.name).trim();
          if (s.status && oneOf(s.status, PROJECT_STATUSES)) p.status = oneOf(s.status, PROJECT_STATUSES);
          for (const k of ['start', 'due']) if (k in s) p[k] = isDate(s[k]) ? s[k] : '';
          if ('ref_name' in s) p.refName = String(s.ref_name || '');
          if ('ref_url' in s) p.refUrl = String(s.ref_url || '');
          if (s.notes) addNote(p, s.notes);
          p.updated = now;
          applied.push(`updated project ${p.name}`); break;
        }
        case 'snooze': {
          const f = find(data.tasks, 'title', o.match);
          if (!f.rec) throw new Error(f.err);
          const u = normRemind(o.until);
          f.rec.snooze = u; if (u) { f.rec.remind = u; f.rec.reminded = ''; }
          f.rec.updated = now;
          applied.push(u ? `snoozed ${f.rec.title} until ${u}` : `unsnoozed ${f.rec.title}`); break;
        }
        case 'add_suggestion': {
          if (!o.title) throw new Error('suggestion needs a title');
          data.digest = data.digest || [];
          const date = isDate(o.date) ? o.date : today, kind = o.kind === 'meeting' ? 'meeting' : 'action';
          if (data.digest.some(g => g.date === date && g.kind === kind && lc(g.title) === lc(o.title))) { applied.push(`already in the digest: ${o.title}`); break; }
          data.digest.push({ id: uid(), date, kind, time: o.time || '', title: String(o.title).trim(), project: o.project || '', due: isDate(o.due) ? o.due : '',
            source: o.source || '', notes: o.notes || '', status: 'new', updated: now });
          applied.push(`digest ${kind}: ${o.title}`); break;
        }
        case 'add_item': {
          if (!o.item || !o.project) throw new Error('checklist item needs item and project');
          data.items = data.items || [];
          const pid = projectId(data, o.project, now);
          if (data.items.some(i => i.projectId === pid && lc(i.title) === lc(o.item))) { applied.push(`already on the checklist: ${o.item}`); break; }
          const it = { id: uid(), projectId: pid, title: String(o.item).trim(), done: o.done ? 'Yes' : '', doneOn: o.done ? today : '', notes: '', updated: now };
          if (o.notes) addNote(it, o.notes);
          data.items.push(it);
          applied.push(`checklist ${o.done ? '(done) ' : ''}${it.title}`); break;
        }
        case 'check_item': case 'uncheck_item': {
          let pool = data.items || [];
          if (o.project) { const pf = find(data.projects, 'name', o.project); if (pf.rec) pool = pool.filter(i => i.projectId === pf.rec.id); }
          const f = find(pool, 'title', o.match);
          if (!f.rec) throw new Error(f.err);
          f.rec.done = o.op === 'check_item' ? 'Yes' : ''; f.rec.doneOn = f.rec.done ? today : ''; f.rec.updated = now;
          applied.push(`${f.rec.done ? 'checked' : 'unchecked'}: ${f.rec.title}`); break;
        }
        default: throw new Error(`unknown op "${o.op}"`);
      }
    } catch (e) { failed.push({ op: o, why: e.message }); }
  }
  return { applied, failed };
}

// Inbox file: one JSON op per line. Blank lines and lines starting with # are ignored.
export function parseInbox(text) {
  const ops = [], bad = [];
  for (const line of String(text || '').split(/\r?\n/)) {
    const s = line.trim();
    if (!s || s.startsWith('#')) continue;
    try { const v = JSON.parse(s); (Array.isArray(v) ? v : [v]).forEach(o => ops.push(o)); }
    catch { bad.push(s); }
  }
  return { ops, bad };
}

// Reminders that should fire now.
export function dueReminders(tasks, nowText) {
  return tasks.filter(t => t.remind && !t.reminded && t.status !== 'Done' && t.remind <= nowText);
}
