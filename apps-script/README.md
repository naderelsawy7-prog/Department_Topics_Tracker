# EPROM dashboard Google Apps Script API

This project is the shared backend for the GitHub Pages dashboard. It reads and writes the team workbook in Google Drive and provides a small authenticated web API.

## What the backend does

- Authenticates the signed-in Google account with `Session.getActiveUser().getEmail()`.
- Authorizes only `ALLOWED_EMAILS` and/or `ALLOWED_DOMAIN` from `Config.gs`.
- Converts the Drive `.xlsx` workbook to a Google Sheets mirror on first setup when needed.
- Stores the complete dashboard JSON in a hidden `Dashboard State` sheet so member profiles, project colors, progress overrides, comments, and task fields survive Excel synchronization.
- Rebuilds the visible `Topics` sheet from the dashboard state.
- Exports the mirror to `.xlsx` after an accepted save and updates the original Excel file in place.
- Uses a monotonically increasing revision and Apps Script `LockService` to reject stale concurrent saves.
- Provides JSONP reads and form POST writes so the static GitHub Pages client does not need a CORS proxy.

## One-time setup

1. Put `Code.gs`, `Config.gs`, and `appsscript.json` into a new Google Apps Script project. In the Apps Script editor, use **Project Settings → Show "appsscript.json" manifest file in editor** before replacing the manifest.
2. In `Config.gs`, set:
   - `WORKBOOK_FILE_ID` to the shared Excel file ID. The supplied starter value is the file found as `Department_Topics_Tracker.xlsx`.
   - `ALLOWED_DOMAIN` to the team's Google Workspace domain, for example `example.com`, and/or list individual addresses in `ALLOWED_EMAILS`.
   - Leave `MIRROR_SHEET_ID` blank for first setup unless a native Google Sheets mirror already exists.
3. If the Advanced Drive service is not enabled automatically by the manifest, click **Services (+)**, add **Google Drive API**, version **v2**. In the linked Google Cloud project, enable the Google Drive API if Apps Script asks for it.
4. Make sure the Apps Script owner can edit the workbook. If the workbook is in a Shared drive, keep the mirror in that drive or give the owner the required Shared drive role.
5. Run `setupDashboard()` once from the editor and approve the Drive, Sheets, and user-email scopes. The first run creates a Google Sheets mirror beside the Excel workbook and stores its ID in Script Properties.
6. If the Drive API cannot convert the Excel file in your organization's policy, open the Excel file with Google Sheets and create a native Google Sheets copy in the same team location. Put that copy's ID into `MIRROR_SHEET_ID`, run `setupDashboard()`, and keep the original `.xlsx` as the exported source. The API will still update the original Excel file after saves.

## Deploy the web app

1. Use **Deploy → New deployment → Web app**.
2. Execute as the backend owner (the account that can edit the Shared drive workbook).
3. Set access to the team's Google Workspace organization/domain. Do not use anonymous public access for a team dashboard.
4. Deploy, complete any authorization/verification prompts, and copy the URL ending in `/exec`. Do not use the `/dev` test URL in GitHub Pages.
5. Paste the `/exec` URL into `team-dashboard-github-pages/online-config.js` as `apiUrl`.

The backend intentionally refuses requests when the caller email is blank or not allowlisted. If a team member gets `AUTH_REQUIRED`, verify that they are signed in to the same Workspace organization and that the web-app access setting permits that organization. If they get `FORBIDDEN`, add their address or correct the domain in `Config.gs`, rerun `setupDashboard()`, and create a new deployment version if required.

## Request contract

### Read shared state

`GET /exec?action=get&callback=someCallback`

Returns:

```json
{
  "ok": true,
  "state": {"projects": [], "tasks": [], "teamProfiles": {}},
  "revision": "7"
}
```

### Save shared state

The dashboard submits a form POST with a `payload` field containing JSON:

```json
{
  "action": "save",
  "requestId": "browser-id-timestamp-sequence",
  "clientId": "browser-id",
  "baseRevision": "7",
  "state": {"projects": [], "tasks": [], "teamProfiles": {}}
}
```

The result is placed in a short-lived Apps Script cache under `requestId`; the browser polls `GET /exec?action=status&requestId=...` using JSONP. A stale `baseRevision` returns `code: "CONFLICT"` and does not write the workbook.

## Workbook behavior

The original `.xlsx` file remains the Drive source. Apps Script uses a native Google Sheets mirror because Apps Script cannot edit an Office Open XML workbook directly with `SpreadsheetApp`. The mirror contains:

- `Topics` — visible project/task rows compatible with the existing tracker.
- `Legend (read me)` — retained when the source workbook is converted.
- `Dashboard State` — hidden JSON chunks containing the full dashboard state.

Every accepted online save writes the state sheet, refreshes `Topics`, exports the mirror as an `.xlsx`, and replaces the content of the original Drive file without changing its file ID. The dashboard keeps project/task links in the full state, so project owner/status/progress/target-date changes and task edits stay synchronized. Users should normally edit through the dashboard; direct Excel edits are supported as an import path through **Import Excel**. If someone replaces the workbook directly in Drive, the next API read refreshes a managed mirror from the newer source file and falls back to the visible `Topics` rows if no hidden state sheet is present.

## Important security notes

- `apiUrl`, workbook ID, and mirror ID are not credentials.
- Never put OAuth client secrets, service-account keys, or passwords in `Config.gs` committed to GitHub.
- Keep the workbook and mirror private. Team membership/Workspace sharing should be narrower than the API deployment scope.
- Use explicit `ALLOWED_EMAILS` when a whole Workspace domain is too broad.
- Standard GitHub Pages is a public static host. The deployment package contains an empty initial data block, but page visibility itself is not private. Use private Pages or an identity-aware proxy if the page URL must also be restricted.
