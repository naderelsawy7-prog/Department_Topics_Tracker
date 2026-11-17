# EPROM — Energy Efficiency and Sustainability dashboard

## Files

- `team-dashboard.html` — the offline dashboard. Open it directly in a browser; no installation or internet is required.
- `Department_Topics_Tracker.xlsx` — the authoritative workbook imported from the connected Drive file. It contains the `Topics` and `Legend (read me)` sheets.
- `project-pulse-data.json` — an optional local JSON backup generated from the Topics workbook.
- `team-dashboard-github-pages/` — GitHub Pages-ready team client; edit `online-config.js` with the Apps Script `/exec` URL.
- `apps-script/` — Google Apps Script API and setup files for shared Drive Excel persistence.

## Drive-based Topics Tracker workflow

The dashboard is now based on the Drive workbook named `Department_Topics_Tracker.xlsx`. Each row in the `Topics` sheet becomes a tracked project/topic with its next task, owner, helper/team, vendor, target client, status, comment, and source.

1. Open `team-dashboard.html`.
2. Click **Connect Excel** and choose `Department_Topics_Tracker.xlsx` from the team location.
3. The dashboard imports the `Topics` sheet and uses it as the starting data.
4. Add or update projects and tasks. Project owner uses the fixed department team dropdown: `Eng. Mohamed El Yamany`, `Eng. Ahmed El Sharnouby`, `Dr. Aya`, or `Eng. Nader`. Helper uses the same member choices and retains the original workbook's Helping team choices.
5. In **Projects in motion**, click anywhere on a project card—or use its pencil button—to open the project edit box. The box includes project ownership, status, target date, comment/notes, and editable project progress. Project progress, owner, status, and target date are linked to the project's tasks: saving a project update propagates those shared fields to its linked task rows, while editing a task recalculates the project card progress/status/target date.
6. In **All tasks**, edit a task normally. The linked project card updates immediately after saving. Moving a task to another project refreshes both the old and new project summaries.
7. After a project or task is saved, the dashboard automatically writes the updated `Topics` and `Legend (read me)` workbook back to the connected Excel file when the browser has write permission.
6. If direct file access is unavailable, click **Export Excel** and replace the shared `Department_Topics_Tracker.xlsx` manually.

Use **one editor at a time**. If two people edit simultaneously, the person who saves last can overwrite the other person's changes.

## Automatic saving

For direct automatic saving, use a Chromium-based desktop browser such as Chrome or Microsoft Edge. Click **Connect Excel** once, select the workbook, and grant write permission. The top-bar status shows whether the workbook is connected, ready, or needs attention.

The dashboard also saves the current state in the browser's local storage. **Sync Excel** forces an immediate write to the connected workbook.

Excel export refreshes only the `Topics` sheet rows and the dependent table/validation/conditional-format ranges from the embedded authorized workbook template. The `TopicsTable`, `Legend (read me)` sheet, styles, widths, freeze pane, date format, workbook metadata, filters, and original validation choices are retained; the four dashboard member labels are also allowed in `Helping team` so Helper selections remain valid in the exported workbook.

## Other exports and backups

- **JSON backup**: exports the complete dashboard data as JSON.
- **JSON import**: restores a JSON backup when needed.
- **Right-click Export Excel**: exports the task list as CSV.
- **Double-click Export Excel**: exports a portable HTML snapshot with the current data embedded inside it.
- **Ctrl/Cmd + Shift + E**: exports the task list as CSV.

## Department setup included

- Company: EPROM.
- Department: Energy Efficiency and Sustainability.
- EPROM department logo (the supplied section logo) embedded in the HTML.
- Color palette inspired by the logo: blue, green, orange, and red.
- Department goal: work in the field of energy and renewable energy.
- Topics and task rows are based on the imported `Department_Topics_Tracker.xlsx` workbook.

## What is included

- Topic/project cards with category, target client, owner, helper, progress, status, comments, and deadlines; clicking a card opens its edit box.
- Clickable team member profiles showing all owned projects, related tasks, delivery metrics, and editable profile details.
- Clickable metric and highlight cards for Active projects, Open tasks, Overdue, Completion rate, Milestones delivered, Energy initiatives, Team focus, and Next delivery.
- Clickable Delivery snapshot: the donut opens completion details and every status row opens its matching tasks.
- Clickable Team workload: each owner row opens the member profile or the owner's assigned tasks.
- Clickable Upcoming deadlines: each deadline opens the corresponding task editor. Projects with a Start date and Target date lasting more than 30 days enter the list 7 days before target; shorter projects enter it 3 days before target. If a task has no due date, its project's Target date is used. If no item is inside the reminder window, the next five dated open tasks are shown as a fallback. Overdue open tasks remain visible.
- Filters by project, owner, status, and priority.
- Add, edit, and delete projects and tasks.
- Task status: Not Started, In Progress, Review, Completed, Blocked.
- Local browser saving plus automatic Excel synchronization and JSON backup import/export.

## Team-only online deployment

The recommended shared architecture is:

```text
GitHub Pages frontend  →  Google Apps Script web API  →  Google Sheets mirror  →  Department_Topics_Tracker.xlsx in Google Drive
```

GitHub Pages is a static host and cannot securely persist multi-user edits by itself. The deployment package is in `team-dashboard-github-pages/`; the backend source and setup instructions are in `apps-script/`.

### 1. Configure the Apps Script backend

1. Create a Google Apps Script project and add `apps-script/Code.gs`, `apps-script/Config.gs`, and `apps-script/appsscript.json`.
2. In `Config.gs`, keep the workbook ID for `Department_Topics_Tracker.xlsx` or replace it with the team's Drive workbook ID. Set `ALLOWED_DOMAIN` to the team's Workspace domain and/or add the individual team addresses to `ALLOWED_EMAILS`.
3. Enable the Advanced Drive service (Google Drive API v2) if the manifest does not enable it automatically.
4. Run `setupDashboard()` once. The backend creates or uses a Google Sheets mirror, keeps the complete dashboard JSON in a hidden `Dashboard State` sheet, and retains the visible `Topics` and `Legend (read me)` workflow.
5. Deploy a **Web app**, execute as the backend owner, and limit access to the team's Google Workspace organization/domain. Copy the deployed URL ending in `/exec`.

The API rejects signed-out users, users outside the allowlist, and stale revision saves. Do not deploy the team API as an anonymous public service. If Excel conversion is blocked by an organization's Drive policy, make a native Google Sheets mirror manually and set `MIRROR_SHEET_ID` in `Config.gs`.

### 2. Configure and publish the GitHub Pages package

1. Edit `team-dashboard-github-pages/online-config.js` and paste the Apps Script `/exec` URL into `apiUrl`. Keep `autoLoad`, `autoSave`, and `requireOnline` set to `true` for the team deployment.
2. Upload the contents of `team-dashboard-github-pages/` to a GitHub repository. Enable **Settings → Pages** and publish the branch/folder containing `index.html`.
3. Open the Pages URL while signed in to an allowlisted Google account. The package starts with an empty data block and shows a team-access screen until the shared API load succeeds.

Standard GitHub Pages can still expose the empty UI shell to anyone who knows the URL. It does not contain the workbook records. If the page itself must be hidden from non-team users, use GitHub Enterprise private Pages or the organization's identity-aware proxy.

### 3. Shared online workflow

- **Load online** reads the shared state from the Drive-backed API. The deployment package loads it automatically.
- Project/task/member/profile edits are saved automatically after a short debounce and can be forced with **Save online**.
- If another person saves first, the stale save is rejected. Load the latest state, review the change, and save again rather than overwriting it.
- **Import Excel** still accepts `.xlsx` and `.csv` files. After import, the current state can be saved to the shared Drive workbook.
- **Export Excel** still downloads a current `.xlsx` file; JSON, CSV, and portable HTML exports remain available.
- **Connect Excel** is a separate local Chromium File System Access feature. It is not the shared online backend and should not be used as the team's multi-user storage.

The API URL and Drive file ID are identifiers, not secrets. Never place service-account keys, OAuth secrets, or passwords in the GitHub Pages package.
