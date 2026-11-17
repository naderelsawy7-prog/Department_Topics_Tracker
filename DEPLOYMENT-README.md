# EPROM Team Dashboard deployment package

The dashboard is ready as a package, but it has not been pushed to GitHub or deployed because no repository and no Apps Script deployment URL were provided.

## Package map

- `team-dashboard-github-pages/index.html` — GitHub Pages-ready, online-only dashboard frontend with empty initial data.
- `team-dashboard-github-pages/online-config.js` — edit only this file to set the Apps Script `/exec` URL and deployment options.
- `team-dashboard-github-pages/README.md` — GitHub Pages and online workflow instructions.
- `apps-script/Code.gs` — authenticated Apps Script API, Drive workbook mirror, revision checks, and Excel export-backend.
- `apps-script/Config.gs` — workbook and team allowlist configuration.
- `apps-script/appsscript.json` — Apps Script V8 manifest with Drive/Sheets scopes and Advanced Drive service.
- `apps-script/README.md` — Apps Script setup, deployment, access control, and troubleshooting.
- `TEAM-DASHBOARD-GUIDE.md` — the existing dashboard guide extended with the online architecture.

## Required administrator actions

1. Share `Department_Topics_Tracker.xlsx` with the team or place it in the team's private Shared drive. The current Drive file ID is `1w-Kvknl-WBwW2hN-f4X-qBiGjLHW2FVN`; the connected Drive account currently reports the file owner as `naderelsawy7@gmail.com`, so team sharing still needs to be confirmed by the Drive owner/admin.
2. Set the allowed Workspace domain or explicit team emails in `apps-script/Config.gs`.
3. Deploy the Apps Script web app for the team and copy its `/exec` URL.
4. Paste that URL into `team-dashboard-github-pages/online-config.js`.
5. Upload the contents of `team-dashboard-github-pages/` to GitHub and enable Pages.

The frontend keeps Excel import and download/export. Online saves update the Drive Excel source through the Apps Script Google Sheets mirror and use revision checks to prevent stale browsers from overwriting a teammate's update.
