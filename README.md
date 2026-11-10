# Projects

Personal projects, to-dos and timeline as an installable desktop app (PWA), styled like the MHS ESK/TMS trackers.

- **Your data lives in a workbook in your OneDrive folder.** Every PC with the app pointed at the same file sees the same projects. Nothing goes to GitHub or anywhere else.
- First PC: Workbook > Create a workbook in OneDrive. Other PCs: Workbook > Open my workbook, pick the same file.
- The app saves about 2 sec after each change, and checks for changes from other devices when it opens, when you switch back to it, and every 30 sec.
- Edits made on two devices before either syncs are merged by record (the newer edit wins for a record changed on both).
- Phone or tablet: open the workbook in the Excel or OneDrive app and edit the Tasks sheet. Mobile browsers can't run the sync.
- Close the workbook in Excel on a PC when done; the app can't save while Excel has it open and shows "Waiting for Excel to close".
- Each PC also keeps a working copy, so the app opens instantly and works offline.
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
