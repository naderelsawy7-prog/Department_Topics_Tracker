/**
 * Edit this file once, then run setupDashboard() from the Apps Script editor.
 * Do not put this file in the GitHub Pages package if you add private emails.
 */
var DASHBOARD_CONFIG = {
  // Current Drive workbook found for this dashboard. Replace when using another file.
  WORKBOOK_FILE_ID: '1w-Kvknl-WBwW2hN-f4X-qBiGjLHW2FVN',
  WORKBOOK_FILE_NAME: 'Department_Topics_Tracker.xlsx',

  // Use either explicit addresses, a Workspace domain, or both.
  // Keep this allowlist narrow: the GitHub Pages site is only the UI shell.
  ALLOWED_EMAILS: [],
  ALLOWED_DOMAIN: 'REPLACE_WITH_YOUR_WORKSPACE_DOMAIN',

  // Leave blank on first setup. setupDashboard() creates a Google Sheets mirror
  // beside the Excel file and stores its ID in Script Properties.
  MIRROR_SHEET_ID: '',
  AUTO_REFRESH_MIRROR: true
};
