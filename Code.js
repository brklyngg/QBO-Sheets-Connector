/**
 * Spreadsheet entry points that bridge the lean auth-first build with the legacy feature set.
 */

const SCRIPT_VERSION = '1.1.2';
const SIDEBAR_TITLE = 'QBO Connection';
const SIDEBAR_WIDTH = 360;
const MENU_TITLE = `QBO Connection v${SCRIPT_VERSION}`;
const CONFIG_SHEET_NAME = '_QBO_Config';
const LOGS_SHEET_NAME = 'QBO_Connector_Logs';
const DEFAULT_MINOR_VERSION = (typeof QBO_DEFAULT_MINOR_VERSION !== 'undefined')
  ? QBO_DEFAULT_MINOR_VERSION
  : '75';

let globalJobStatus = null;

function onInstall(e) {
  onOpen(e);
  initializeConfigSheet();
  initializeLogsSheet();
  ensureLogFlushTrigger();
  cleanupOrphanedTriggers();
}

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  addMenus(ui);

  try {
    ensureLogFlushTrigger();
  } catch (error) {
    console.error('Error ensuring log flush trigger:', error);
  }

  try {
    cleanupOrphanedTriggers();
  } catch (error) {
    console.error('Error cleaning orphaned triggers:', error);
  }

  try {
    logAction('addon_open', {
      authMode: e ? e.authMode : 'UNKNOWN',
      version: SCRIPT_VERSION
    });
  } catch (error) {
    console.error('Error logging onOpen event:', error);
  }
}

function addMenus(ui) {
  try {
    const standardMenu = ui.createMenu(MENU_TITLE)
      .addItem('Open Connector', 'showSidebar')
      .addSeparator()
      .addItem('View Logs', 'showLogsSheet')
      .addItem('Refresh All Datasets', 'refreshAllDatasets')
      .addSeparator()
      .addItem('Settings', 'showSettings')
      .addItem('About', 'showAbout')
      .addSeparator()
      .addItem('Version Info', 'showVersionInfo');
    standardMenu.addToUi();
  } catch (error) {
    console.error('Error creating spreadsheet menu:', error);
  }

  try {
    const addonMenu = ui.createAddonMenu()
      .addItem('Open Connector', 'showSidebar')
      .addItem('View Logs', 'showLogsSheet')
      .addItem('Refresh All Datasets', 'refreshAllDatasets')
      .addSeparator()
      .addItem('Settings', 'showSettings')
      .addItem('About', 'showAbout')
      .addSeparator()
      .addItem('Version Info', 'showVersionInfo');
    addonMenu.addToUi();
  } catch (error) {
    console.error('Error creating add-on menu:', error);
  }
}

function showSidebar(initialTab) {
  const template = HtmlService.createTemplateFromFile('UI');
  template.initialStateJson = JSON.stringify(loadSidebarData(initialTab));

  const htmlOutput = template.evaluate()
    .setTitle(SIDEBAR_TITLE)
    .setWidth(SIDEBAR_WIDTH);

  SpreadsheetApp.getUi().showSidebar(htmlOutput);

  logAction('show_sidebar', {
    initialTab: initialTab || 'connection'
  });
}

function loadSidebarData(initialTab) {
  return {
    status: getConnectionStatus(),
    redirectUri: getOAuthRedirectUri(),
    datasets: getDatasets(),
    reports: getAvailableReports(),
    entities: getAvailableEntities(),
    scheduleStatus: getScheduleStatus(),
    logStats: getLogStats(),
    initialTab: initialTab || 'connection'
  };
}

function saveConnectionSettings(settings) {
  if (!settings) {
    throw new Error('No settings payload supplied.');
  }

  const clientId = settings.clientId || '';
  const clientSecret = settings.clientSecret || '';
  const environment = settings.environment || 'sandbox';

  setOAuthCredentials(clientId, clientSecret, environment);
  return getConnectionStatus();
}

function startAuthorization() {
  return {
    authorizationUrl: getAuthorizationUrl()
  };
}

function disconnectFromQuickBooks() {
  disconnect();
  return getConnectionStatus();
}

function refreshConnectionStatus() {
  return getConnectionStatus();
}

function runConnectionTest() {
  return testQuickBooksConnection();
}

function getOAuthRedirectUri() {
  return 'https://script.google.com/macros/d/' + ScriptApp.getScriptId() + '/usercallback';
}

function showLogsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logsSheet = ss.getSheetByName(LOGS_SHEET_NAME);

    if (!logsSheet) {
      logsSheet = initializeLogsSheet();
    }

    ss.setActiveSheet(logsSheet);
    logAction('show_logs');
  } catch (error) {
    showError('Failed to show logs', error);
  }
}

function refreshAllDatasets() {
  try {
    const datasets = getDatasets();
    const enabledDatasets = datasets.filter(d => d.schedule && d.schedule.enabled);

    if (enabledDatasets.length === 0) {
      showToast('No scheduled datasets found', 'QBO Connection');
      return;
    }

    const jobId = Utilities.getUuid();
    const jobs = enabledDatasets.map(dataset => ({
      datasetId: dataset.id,
      status: 'queued',
      startTime: null,
      endTime: null,
      error: null
    }));

    const userProperties = getUserProps();
    userProperties.setProperty('refreshJob_' + jobId, JSON.stringify({
      id: jobId,
      jobs: jobs,
      startTime: new Date().toISOString(),
      status: 'running'
    }));

    processRefreshJob(jobId);
    showToast(`Refreshing ${enabledDatasets.length} datasets...`, 'QBO Connection');

    logAction('refresh_all_datasets', {
      count: enabledDatasets.length,
      jobId: jobId
    });
  } catch (error) {
    showError('Failed to refresh datasets', error);
  }
}

function showSettings() {
  showSidebar('settings');
  logAction('show_settings');
}

function showAbout() {
  const ui = SpreadsheetApp.getUi();

  const message = `QuickBooks Online Connector\n\nVersion: ${SCRIPT_VERSION}\n\nThis add-on imports data from QuickBooks Online into Google Sheets.\n\nFeatures:\n• OAuth 2.0 authentication\n• Standard reports and custom queries\n• Dataset scheduling with triggers\n• Comprehensive logging`;

  ui.alert('About', message, ui.ButtonSet.OK);
  logAction('show_about');
}

function showVersionInfo() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Running version ${SCRIPT_VERSION}`, 'QuickBooks Online Connector', 5);
    logAction('show_version', {
      version: SCRIPT_VERSION
    });
  } catch (error) {
    console.error('Error showing version info:', error);
  }
}

function initializeConfigSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

    if (!configSheet) {
      configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
      configSheet.hideSheet();

      const headers = ['Key', 'Value', 'Updated', 'Type'];
      configSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      configSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4285F4')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');

      configSheet.setColumnWidth(1, 200);
      configSheet.setColumnWidth(2, 400);
      configSheet.setColumnWidth(3, 150);
      configSheet.setColumnWidth(4, 100);

      const defaults = [
        ['version', SCRIPT_VERSION, new Date().toISOString(), 'system'],
        ['minorVersion', DEFAULT_MINOR_VERSION, new Date().toISOString(), 'config'],
        ['maxCellsWarning', '2000000', new Date().toISOString(), 'config'],
        ['maxCellsHard', '8000000', new Date().toISOString(), 'config'],
        ['logBatchSize', String(LOG_BATCH_SIZE), new Date().toISOString(), 'config'],
        ['logRotationRows', String(LOG_ROTATION_ROWS), new Date().toISOString(), 'config']
      ];

      configSheet.getRange(2, 1, defaults.length, 4).setValues(defaults);
    }

    return configSheet;
  } catch (error) {
    console.error('Error initializing config sheet:', error);
    throw error;
  }
}

function initializeLogsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logsSheet = ss.getSheetByName(LOGS_SHEET_NAME);

    if (!logsSheet) {
      logsSheet = ss.insertSheet(LOGS_SHEET_NAME);

      const headers = [
        'ts_iso', 'user_email', 'schedule_owner_email', 'action', 'dataset_type',
        'dataset_id', 'report_name', 'params_json', 'start_date', 'end_date',
        'method', 'query_select', 'query_from', 'query_where', 'query_orderby',
        'query_startposition', 'query_maxresults', 'realmId', 'minorversion',
        'target_sheet', 'target_named_range', 'rows', 'cols', 'elapsed_ms',
        'status', 'http_status', 'error_code', 'error_message', 'intuit_tid',
        'retry_after', 'schedule_id', 'schedule_freq', 'request_id', 'job_id',
        'dataset_version', 'qbo_endpoint', 'response_bytes', 'transport'
      ];

      logsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      logsSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#34A853')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');

      logsSheet.setFrozenRows(1);
      logsSheet.setColumnWidth(1, 180);
      logsSheet.setColumnWidth(2, 200);
      logsSheet.setColumnWidth(4, 150);
      logsSheet.setColumnWidth(28, 300);
    }

    return logsSheet;
  } catch (error) {
    console.error('Error initializing logs sheet:', error);
    throw error;
  }
}

function showError(message, error) {
  console.error(message, error);

  const ui = SpreadsheetApp.getUi();
  const details = error ? `\n\nDetails: ${error.toString()}` : '';
  ui.alert('Error', message + details, ui.ButtonSet.OK);

  logAction('error', {
    message: message,
    error: error ? error.toString() : null,
    stack: error && error.stack ? error.stack : null
  });
}

function showToast(message, title) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title || SIDEBAR_TITLE, 5);
}

function getConfig(key, defaultValue) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

    if (!configSheet) {
      return defaultValue;
    }

    const data = configSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        return data[i][1];
      }
    }

    return defaultValue;
  } catch (error) {
    console.error('Error getting config:', error);
    return defaultValue;
  }
}

function setConfig(key, value, type = 'config') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

    if (!configSheet) {
      configSheet = initializeConfigSheet();
    }

    const data = configSheet.getDataRange().getValues();
    let found = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        configSheet.getRange(i + 1, 2, 1, 3).setValues([[value, new Date().toISOString(), type]]);
        found = true;
        break;
      }
    }

    if (!found) {
      const newRow = configSheet.getLastRow() + 1;
      configSheet.getRange(newRow, 1, 1, 4).setValues([[key, value, new Date().toISOString(), type]]);
    }
  } catch (error) {
    console.error('Error setting config:', error);
    throw error;
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function processRefreshJob(jobId) {
  const userProperties = getUserProps();
  const jobDataRaw = userProperties.getProperty('refreshJob_' + jobId);
  if (!jobDataRaw) {
    console.error('Job not found:', jobId);
    return;
  }

  const jobData = JSON.parse(jobDataRaw);

  for (let i = 0; i < jobData.jobs.length; i++) {
    const job = jobData.jobs[i];

    if (job.status === 'queued') {
      try {
        job.status = 'running';
        job.startTime = new Date().toISOString();

        const dataset = getDatasetById(job.datasetId);
        if (dataset) {
          runDataset(dataset.id);
          job.status = 'completed';
        } else {
          job.status = 'error';
          job.error = 'Dataset not found';
        }
      } catch (error) {
        job.status = 'error';
        job.error = error.toString();
      } finally {
        job.endTime = new Date().toISOString();
      }
    }
  }

  jobData.status = jobData.jobs.every(j => j.status === 'completed') ? 'completed'
    : jobData.jobs.some(j => j.status === 'error') ? 'error'
    : 'running';
  jobData.endTime = new Date().toISOString();

  userProperties.setProperty('refreshJob_' + jobId, JSON.stringify(jobData));
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function getScriptId() {
  return ScriptApp.getScriptId();
}
