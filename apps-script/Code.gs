/**
 * EPROM Energy Efficiency and Sustainability dashboard API.
 *
 * This web app keeps the full dashboard JSON in a hidden "Dashboard State"
 * sheet, mirrors the visible Topics sheet, and writes the Google Sheets mirror
 * back to the shared Excel workbook in Drive after every accepted save.
 *
 * The endpoint uses JSONP for reads and a regular HTML form POST for writes so
 * a GitHub Pages site can use it without a separate CORS proxy. The API still
 * requires a signed-in, allowlisted Google account; JSONP is not public data.
 */

var DASHBOARD_STATE_SHEET = 'Dashboard State';
var TOPICS_SHEET = 'Topics';
var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
var GOOGLE_SHEET_MIME = 'application/vnd.google-apps.spreadsheet';
var PROP = {
  MIRROR_ID: 'MIRROR_SHEET_ID',
  MIRROR_MANAGED: 'MIRROR_MANAGED',
  SOURCE_MODIFIED: 'SOURCE_MODIFIED_AT',
  REVISION: 'DASHBOARD_REVISION'
};

function doGet(e) {
  try {
    var email = requireAuthorizedUser_();
    var action = String((e && e.parameter && e.parameter.action) || 'get').toLowerCase();
    if (action === 'status') return output_(getSaveStatus_(e.parameter.requestId), e);
    if (action === 'health') return output_({ok: true, email: email, revision: getRevision_()}, e);
    if (action !== 'get') throw apiError_('BAD_REQUEST', 'Unknown GET action.');
    return output_(getStatePayload_(), e);
  } catch (error) {
    return output_(errorPayload_(error), e);
  }
}

function doPost(e) {
  var body = {};
  try {
    body = parsePostBody_(e);
    requireAuthorizedUser_();
    if (String(body.action || '').toLowerCase() !== 'save') throw apiError_('BAD_REQUEST', 'POST action must be save.');
    var result = saveState_(body);
    return output_(result, e);
  } catch (error) {
    var result = errorPayload_(error);
    if (body && body.requestId) cacheSaveStatus_(body.requestId, result);
    return output_(result, e);
  }
}

function setupDashboard() {
  var config = getConfig_();
  requireConfigured_(config);
  var mirror = ensureMirror_();
  Logger.log('Dashboard API is ready. Mirror sheet: ' + mirror.getId() + '; revision: ' + getRevision_());
}

function parsePostBody_(e) {
  var raw = '';
  if (e && e.parameter && e.parameter.payload) raw = e.parameter.payload;
  else if (e && e.postData && e.postData.contents) raw = e.postData.contents;
  if (!raw) return {};
  try { return JSON.parse(raw); } catch (error) { throw apiError_('BAD_REQUEST', 'The save payload is not valid JSON.'); }
}

function output_(payload, e) {
  var text = JSON.stringify(payload);
  var callback = e && e.parameter && e.parameter.callback;
  if (callback && /^[A-Za-z_$][A-Za-z0-9_$]*(?:\.[A-Za-z_$][A-Za-z0-9_$]*)*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + text + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function errorPayload_(error) {
  return {ok: false, code: String(error && error.code || 'SERVER_ERROR'), error: String(error && error.message || error || 'Unknown server error.')};
}

function apiError_(code, message) {
  var error = new Error(message);
  error.code = code;
  return error;
}

function getConfig_() {
  var stored = PropertiesService.getScriptProperties().getProperties();
  var defaults = (typeof DASHBOARD_CONFIG !== 'undefined') ? DASHBOARD_CONFIG : {};
  return {
    WORKBOOK_FILE_ID: String(stored.WORKBOOK_FILE_ID || defaults.WORKBOOK_FILE_ID || '').trim(),
    WORKBOOK_FILE_NAME: String(stored.WORKBOOK_FILE_NAME || defaults.WORKBOOK_FILE_NAME || 'Department_Topics_Tracker.xlsx').trim(),
    ALLOWED_EMAILS: String(stored.ALLOWED_EMAILS || (defaults.ALLOWED_EMAILS || []).join(',')).split(',').map(function(value) { return value.trim().toLowerCase(); }).filter(Boolean),
    ALLOWED_DOMAIN: String(stored.ALLOWED_DOMAIN || defaults.ALLOWED_DOMAIN || '').trim().toLowerCase().replace(/^@/, ''),
    MIRROR_SHEET_ID: String(stored.MIRROR_SHEET_ID || defaults.MIRROR_SHEET_ID || '').trim(),
    AUTO_REFRESH_MIRROR: String(stored.AUTO_REFRESH_MIRROR || defaults.AUTO_REFRESH_MIRROR || 'true').toLowerCase() !== 'false'
  };
}

function requireConfigured_(config) {
  if (!config.WORKBOOK_FILE_ID) throw apiError_('CONFIGURATION', 'WORKBOOK_FILE_ID is not configured.');
  if (/REPLACE_WITH|YOUR_DOMAIN|example\.com/i.test(config.ALLOWED_DOMAIN)) throw apiError_('CONFIGURATION', 'Replace the ALLOWED_DOMAIN placeholder with the team Workspace domain.');
  if (!config.ALLOWED_EMAILS.length && !config.ALLOWED_DOMAIN) throw apiError_('CONFIGURATION', 'Set ALLOWED_EMAILS or ALLOWED_DOMAIN before deploying the API.');
}

function requireAuthorizedUser_() {
  var config = getConfig_();
  requireConfigured_(config);
  var email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) throw apiError_('AUTH_REQUIRED', 'Google did not provide a signed-in team email. Deploy the web app for your Workspace users and sign in before loading the dashboard.');
  var emailAllowed = config.ALLOWED_EMAILS.indexOf(email) !== -1;
  var domainAllowed = config.ALLOWED_DOMAIN && (email === config.ALLOWED_DOMAIN || email.slice(-(config.ALLOWED_DOMAIN.length + 1)) === '@' + config.ALLOWED_DOMAIN);
  if (!emailAllowed && !domainAllowed) throw apiError_('FORBIDDEN', 'This Google account is not on the dashboard team allowlist.');
  return email;
}

function getStatePayload_() {
  var mirror = ensureMirror_();
  var state = readDashboardState_(mirror);
  if (!state) state = readTopicsState_(mirror);
  if (!state) state = {version: 1, lastUpdated: null, teamProfiles: {}, projects: [], tasks: []};
  return {ok: true, state: state, revision: getRevision_(), workbookFileId: getConfig_().WORKBOOK_FILE_ID, workbookFileName: getConfig_().WORKBOOK_FILE_NAME, updatedAt: PropertiesService.getScriptProperties().getProperty(PROP.SOURCE_MODIFIED) || null};
}

function getRevision_() {
  return String(PropertiesService.getScriptProperties().getProperty(PROP.REVISION) || '0');
}

function readDashboardState_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(DASHBOARD_STATE_SHEET);
  if (!sheet || sheet.getLastRow() < 1) return null;
  var values = sheet.getRange(1, 1, sheet.getLastRow(), Math.max(2, sheet.getLastColumn())).getDisplayValues();
  var chunks = [];
  values.forEach(function(row) {
    if (String(row[0]).toLowerCase() === 'statejson' || String(row[0]).toLowerCase().indexOf('statejsonpart') === 0) chunks.push(String(row[1] || ''));
  });
  if (!chunks.length) return null;
  try {
    var state = JSON.parse(chunks.join(''));
    return validState_(state) ? state : null;
  } catch (error) {
    return null;
  }
}

function readTopicsState_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(TOPICS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return null;
  var headers = values.shift().map(normalizeHeader_);
  var projects = [];
  var tasks = [];
  var projectByName = {};
  var projectCount = 0;
  var taskCount = 0;

  values.forEach(function(row) {
    var topic = cellByHeaders_(headers, row, ['topic', 'project', 'project name']);
    if (!topic) return;
    var projectKey = topic.toLowerCase();
    var project = projectByName[projectKey];
    if (!project) {
      projectCount += 1;
      project = {
        id: 'p' + projectCount,
        number: String(cellByHeaders_(headers, row, ['#', 'number']) || projectCount),
        name: topic,
        category: cellByHeaders_(headers, row, ['category']) || 'Project',
        client: cellByHeaders_(headers, row, ['target client company', 'target client', 'client company', 'client']) || '',
        targetClient: cellByHeaders_(headers, row, ['target client company', 'target client', 'client company', 'client']) || '',
        owner: cellByHeaders_(headers, row, ['owner']) || '',
        helper: cellByHeaders_(headers, row, ['helping team', 'helper']) || '',
        vendor: cellByHeaders_(headers, row, ['vendor supplier', 'vendor', 'supplier']) || '',
        source: cellByHeaders_(headers, row, ['source']) || '',
        sourceStatus: cellByHeaders_(headers, row, ['status', 'source status']) || 'Not started',
        status: 'Active',
        startDate: '',
        deadline: '',
        progress: null,
        notes: '',
        color: '#005fa8'
      };
      projectByName[projectKey] = project;
      projects.push(project);
    }
    var dueDate = parseDate_(cellByHeaders_(headers, row, ['task due date dd mm yyyy', 'due date', 'deadline', 'target date']));
    if (dueDate && !project.deadline) project.deadline = dueDate;
    var taskTitle = cellByHeaders_(headers, row, ['task', 'task title', 'title']);
    var status = cellByHeaders_(headers, row, ['status', 'source status']) || project.sourceStatus || 'Not started';
    project.sourceStatus = project.sourceStatus || status;
    if (taskTitle) {
      taskCount += 1;
      tasks.push({
        id: 't' + taskCount,
        title: taskTitle,
        projectId: project.id,
        owner: cellByHeaders_(headers, row, ['task owner', 'owner']) || project.owner || '',
        sourceStatus: status,
        status: dashboardStatus_(status),
        priority: 'Medium',
        dueDate: dueDate,
        progress: dashboardStatus_(status) === 'Completed' ? 100 : 0,
        notes: cellByHeaders_(headers, row, ['comment', 'notes', 'notes blockers']) || ''
      });
    }
  });
  return {version: 1, lastUpdated: new Date().toISOString(), teamProfiles: {}, projects: projects, tasks: tasks};
}

function normalizeHeader_(value) {
  return String(value || '').toLowerCase().replace(/[()\/_.-]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function cellByHeaders_(headers, row, names) {
  for (var i = 0; i < names.length; i += 1) {
    var index = headers.indexOf(normalizeHeader_(names[i]));
    if (index >= 0 && row[index] !== undefined && String(row[index]).trim() !== '') return String(row[index]).trim();
  }
  return '';
}

function parseDate_(value) {
  var text = String(value || '').trim();
  if (!text) return '';
  var match = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (match) return match[1] + '-' + ('0' + match[2]).slice(-2) + '-' + ('0' + match[3]).slice(-2);
  match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);
  if (match) return match[3] + '-' + ('0' + match[2]).slice(-2) + '-' + ('0' + match[1]).slice(-2);
  var parsed = new Date(text);
  if (isNaN(parsed.getTime())) return '';
  return Utilities.formatDate(parsed, Session.getScriptTimeZone() || 'Africa/Cairo', 'yyyy-MM-dd');
}

function dashboardStatus_(value) {
  var text = String(value || '').toLowerCase().replace(/[_-]+/g, ' ').trim();
  if (text.indexOf('complete') >= 0 || text === 'done' || text === 'closed') return 'Completed';
  if (text.indexOf('block') >= 0 || text.indexOf('hold') >= 0) return 'Blocked';
  if (text.indexOf('review') >= 0 || text.indexOf('approval') >= 0) return 'Review';
  if (text.indexOf('progress') >= 0 || text === 'active' || text === 'started') return 'In Progress';
  return 'Not Started';
}

function validState_(state) {
  return state && typeof state === 'object' && Array.isArray(state.projects) && Array.isArray(state.tasks);
}

function cleanState_(state) {
  if (!validState_(state)) throw apiError_('BAD_REQUEST', 'The dashboard state must contain projects and tasks arrays.');
  var copy = JSON.parse(JSON.stringify(state));
  if (copy.projects.length > 1000 || copy.tasks.length > 5000) throw apiError_('BAD_REQUEST', 'The dashboard state is larger than the configured workbook limit.');
  copy.version = 1;
  copy.lastUpdated = new Date().toISOString();
  if (!copy.teamProfiles || typeof copy.teamProfiles !== 'object') copy.teamProfiles = {};
  return copy;
}

function saveState_(body) {
  var config = getConfig_();
  var requestId = String(body.requestId || '').trim();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var currentRevision = getRevision_();
    var baseRevision = String(body.baseRevision || '');
    if (!baseRevision || baseRevision !== currentRevision) {
      var conflict = {ok: false, ready: true, code: 'CONFLICT', error: 'The shared dashboard changed after this browser loaded it.', revision: currentRevision};
      cacheSaveStatus_(requestId, conflict);
      return conflict;
    }
    var state = cleanState_(body.state);
    var mirror = ensureMirror_();
    writeDashboardState_(mirror, state);
    writeTopicsSheet_(mirror, state);
    SpreadsheetApp.flush();
    exportMirrorToWorkbook_(mirror, config);
    var nextRevision = String(Number(currentRevision) + 1);
    var props = PropertiesService.getScriptProperties();
    props.setProperty(PROP.REVISION, nextRevision);
    props.setProperty(PROP.SOURCE_MODIFIED, sourceModifiedDate_(config.WORKBOOK_FILE_ID));
    var result = {ok: true, ready: true, revision: nextRevision, updatedAt: new Date().toISOString(), requestId: requestId};
    cacheSaveStatus_(requestId, result);
    return result;
  } catch (error) {
    var failure = errorPayload_(error);
    failure.ready = true;
    cacheSaveStatus_(requestId, failure);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function cacheSaveStatus_(requestId, result) {
  if (!requestId) return;
  CacheService.getScriptCache().put('dashboard-save-' + requestId, JSON.stringify(result), 600);
}

function getSaveStatus_(requestId) {
  if (!requestId) return {ok: false, ready: true, code: 'BAD_REQUEST', error: 'requestId is required.'};
  var cached = CacheService.getScriptCache().get('dashboard-save-' + String(requestId));
  if (!cached) return {ok: true, ready: false, requestId: String(requestId)};
  try { return JSON.parse(cached); } catch (error) { return {ok: false, ready: true, code: 'SERVER_ERROR', error: 'The save status could not be read.'}; }
}

function writeDashboardState_(spreadsheet, state) {
  var sheet = spreadsheet.getSheetByName(DASHBOARD_STATE_SHEET) || spreadsheet.insertSheet(DASHBOARD_STATE_SHEET);
  var json = JSON.stringify(state);
  var values = [['key', 'value'], ['schema', 'eprom-dashboard-v1']];
  var chunkSize = 45000;
  for (var offset = 0, part = 1; offset < json.length; offset += chunkSize, part += 1) values.push(['stateJsonPart' + part, json.slice(offset, offset + chunkSize)]);
  sheet.clearContents();
  sheet.getRange(1, 1, values.length, 2).setValues(values);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 520);
  try { sheet.hideSheet(); } catch (error) { /* a workbook with one sheet cannot hide it */ }
}

function writeTopicsSheet_(spreadsheet, state) {
  var sheet = spreadsheet.getSheetByName(TOPICS_SHEET) || spreadsheet.insertSheet(TOPICS_SHEET);
  var rows = [['#', 'Topic', 'Category', 'Owner', 'Helping team', 'Vendor / Supplier', 'Target client / company', 'Task', 'Task due date (dd/mm/yyyy)', 'Status', 'Comment', 'Source']];
  var rowNumber = 1;
  (state.projects || []).forEach(function(project) {
    var tasks = (state.tasks || []).filter(function(task) { return String(task.projectId || '') === String(project.id || ''); });
    var rowsForProject = tasks.length ? tasks : [null];
    rowsForProject.forEach(function(task) {
      rows.push([
        project.number || rowNumber,
        project.name || 'Untitled project',
        project.category || 'Project',
        project.owner || (task && task.owner) || '',
        project.helper || '',
        project.vendor || '',
        project.targetClient || project.client || '',
        task ? (task.title || '') : '',
        workbookDate_(task && task.dueDate || project.deadline || ''),
        task ? (task.sourceStatus || task.status || 'Not started') : (project.sourceStatus || project.status || 'Active'),
        task ? (task.notes || '') : (project.notes || ''),
        project.source || ''
      ]);
      rowNumber += 1;
    });
  });
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#005fa8').setFontColor('#ffffff');
  if (rows.length > 1) sheet.getRange(2, 9, rows.length - 1, 1).setNumberFormat('@');
  sheet.autoResizeColumns(1, rows[0].length);
}

function workbookDate_(value) {
  var iso = parseDate_(value);
  if (!iso) return '';
  var parts = iso.split('-');
  return parts[2] + '/' + parts[1] + '/' + parts[0];
}

function ensureMirror_() {
  var config = getConfig_();
  var source = Drive.Files.get(config.WORKBOOK_FILE_ID, {supportsAllDrives: true});
  if (source.mimeType === GOOGLE_SHEET_MIME) {
    PropertiesService.getScriptProperties().setProperty(PROP.SOURCE_MODIFIED, String(source.modifiedDate || ''));
    return SpreadsheetApp.openById(source.id);
  }
  var props = PropertiesService.getScriptProperties();
  var mirrorId = String(props.getProperty(PROP.MIRROR_ID) || config.MIRROR_SHEET_ID || '').trim();
  var managed = String(props.getProperty(PROP.MIRROR_MANAGED) || '') === 'true';
  var storedModified = String(props.getProperty(PROP.SOURCE_MODIFIED) || '');
  var sourceModified = String(source.modifiedDate || '');
  if (mirrorId && managed && config.AUTO_REFRESH_MIRROR && storedModified && sourceModified && storedModified !== sourceModified) {
    try { Drive.Files.remove(mirrorId, {supportsAllDrives: true}); } catch (error) { /* the mirror may already be gone */ }
    mirrorId = '';
  }
  if (!mirrorId) {
    var converted = convertWorkbookToSheet_(source, config);
    mirrorId = converted.id;
    props.setProperty(PROP.MIRROR_ID, mirrorId);
    props.setProperty(PROP.MIRROR_MANAGED, 'true');
  }
  props.setProperty(PROP.SOURCE_MODIFIED, sourceModified);
  return SpreadsheetApp.openById(mirrorId);
}

function convertWorkbookToSheet_(source, config) {
  var parent = source.parents && source.parents.length ? [{id: source.parents[0].id}] : undefined;
  var resource = {title: (config.WORKBOOK_FILE_NAME || source.title || 'Department Topics Tracker') + ' — dashboard mirror', mimeType: GOOGLE_SHEET_MIME};
  if (parent) resource.parents = parent;
  try {
    var copied = Drive.Files.copy(resource, source.id, {supportsAllDrives: true});
    if (copied && copied.mimeType === GOOGLE_SHEET_MIME) return copied;
    if (copied && copied.id) { try { Drive.Files.remove(copied.id, {supportsAllDrives: true}); } catch (ignore) {} }
  } catch (copyError) {
    // Some Drive domains do not convert Office files through Files.copy().
  }
  var blob = DriveApp.getFileById(source.id).getBlob();
  var inserted = Drive.Files.insert(resource, blob, {convert: true, supportsAllDrives: true});
  if (!inserted || inserted.mimeType !== GOOGLE_SHEET_MIME) throw apiError_('WORKBOOK_CONVERSION', 'The Excel workbook could not be converted to a Google Sheets mirror. Create a native Sheets copy and set MIRROR_SHEET_ID.');
  return inserted;
}

function sourceModifiedDate_(fileId) {
  try { return String(Drive.Files.get(fileId, {supportsAllDrives: true}).modifiedDate || new Date().toISOString()); }
  catch (error) { return new Date().toISOString(); }
}

function exportMirrorToWorkbook_(mirror, config) {
  var source = Drive.Files.get(config.WORKBOOK_FILE_ID, {supportsAllDrives: true});
  if (source.mimeType === GOOGLE_SHEET_MIME || source.id === mirror.getId()) return;
  SpreadsheetApp.flush();
  var url = 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(mirror.getId()) + '/export?format=xlsx';
  var response = UrlFetchApp.fetch(url, {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions: true});
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) throw apiError_('WORKBOOK_EXPORT', 'Google Sheets could not export the shared workbook as Excel.');
  var blob = response.getBlob().setName(config.WORKBOOK_FILE_NAME || source.title || 'Department_Topics_Tracker.xlsx').setContentType(XLSX_MIME);
  Drive.Files.update({title: config.WORKBOOK_FILE_NAME || source.title, mimeType: XLSX_MIME}, source.id, blob, {supportsAllDrives: true});
}
