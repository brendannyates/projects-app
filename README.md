# Projects

Personal projects, to-dos and timeline as an installable desktop app (PWA), styled like the MHS ESK/TMS trackers.

- **Data lives on this PC** (browser storage). Nothing is sent anywhere.
- **The Excel workbook is the backup** and the way to edit in Excel: Workbook > Choose where to save (put it in OneDrive), then Save to workbook, or turn on auto-save. Load from workbook brings Excel edits back (one-step undo).
- Weeks run Sunday to Saturday.

## Views

Today (overdue, due today, in progress, next 7 days), This week (drag tasks between days), Projects (optional start/due per project, quick add), Timeline (Gantt; drag bars to move, drag the right edge to change the due date), Workbook.

Keys: `N` new task, `T` Today, `W` week, `P` projects, `L` timeline.

## Workbook

Sheets: Timeline (Gantt via conditional formatting, follows the Tasks sheet; change C1 to scroll), Tasks, Projects (open/done/progress formulas), Read me. Keep the ID column; new rows can leave it blank. Close the file in Excel before the app saves to it.

## Hosting on GitHub Pages

The repo holds code and invented demo data only. Never commit an exported workbook.

```bash
git init && git add . && git commit -m "Projects app"
gh repo create projects-app --public --source . --push
gh api -X POST repos/{owner}/projects-app/pages -f "source[branch]=main" -f "source[path]=/"
```

Then open `https://<owner>.github.io/projects-app/` in Edge and install it (address bar install icon, or ... > Apps > Install this site as an app). Try `?demo=1` to see invented sample data kept apart from your real data.

When you change a file, bump `VERSION` in `sw.js` so installed copies pick it up.

## Full screen

The app window opens maximized. Browsers only allow true full screen after a click or key press, so with "Go full screen when the app opens" on, it switches on your first click. F11 also works.

## Files

`js/core.js` pure date/state logic · `js/store.js` storage and demo data · `js/workbook.js` Excel export/import (ExcelJS, vendored in `vendor/`) · `js/app.js` UI · `sw.js` offline cache.
