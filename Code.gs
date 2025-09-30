/**
 * QuickBooks Online Connector for Google Sheets
 * Main entry point and add-on lifecycle functions
 */

// Constants
const ADDON_TITLE = 'QuickBooks Online Connector';
const CONFIG_SHEET_NAME = '_QBO_Config';
const LOGS_SHEET_NAME = 'QBO_Connector_Logs';
const DEFAULT_MINOR_VERSION = '75';
const SCRIPT_VERSION = '1.0.3';

// Global state for UI updates
let globalJobStatus = null;

/**
 * Add-on install event - runs once when add-on is installed
 */
function onInstall(e) {
  onOpen(e);
  
  // Initialize config sheet
  initializeConfigSheet();
  
  // Initialize logs sheet
  initializeLogsSheet();

  // Ensure background maintenance is configured
  ensureLogFlushTrigger();
  cleanupOrphanedTriggers();
}

/**
 * Add-on open event - adds menu items
 */
function onOpen(e) {
  try {
    const ui = SpreadsheetApp.getUi();
    
    const versionLabel = `Version ${SCRIPT_VERSION}`;

    ui.createAddonMenu()
      .addItem('Open Connector', 'showSidebar')
      .addSeparator()
      .addItem('View Logs', 'showLogsSheet')
      .addItem('Refresh All Datasets', 'refreshAllDatasets')
      .addSeparator()
      .addItem('Settings', 'showSettings')
      .addItem('About', 'showAbout')
      .addSeparator()
      .addItem(versionLabel, 'showVersionInfo')
      .addToUi();
      
    ensureLogFlushTrigger();
    cleanupOrphanedTriggers();

    // Log add-on open
    logAction('addon_open', {
      authMode: e ? e.authMode : 'UNKNOWN',
      version: SCRIPT_VERSION
    });
  } catch (error) {
    console.error('Error in onOpen:', error);
  }
}

/**
 * Shows the main sidebar UI
 */
function showSidebar(initialTab) {
  try {
    const template = HtmlService.createTemplateFromFile('UI');
    template.isConnected = isConnected();
    template.userEmail = Session.getActiveUser().getEmail();
    template.initialTab = initialTab || 'datasets';
    
    const html = template.evaluate()
      .setTitle(ADDON_TITLE)
      .setWidth(350);
      
    SpreadsheetApp.getUi().showSidebar(html);
    
    logAction('show_sidebar', {
      isConnected: template.isConnected,
      initialTab: template.initialTab
    });
  } catch (error) {
    showError('Failed to open sidebar', error);
  }
}

/**
 * Shows the logs sheet or creates it if it doesn't exist
 */
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

/**
 * Refreshes all enabled datasets
 */
function refreshAllDatasets() {
  try {
    const datasets = getDatasets();
    const enabledDatasets = datasets.filter(d => d.schedule && d.schedule.enabled);
    
    if (enabledDatasets.length === 0) {
      showToast('No scheduled datasets found', 'Info');
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
    
    // Store job status
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('refreshJob_' + jobId, JSON.stringify({
      id: jobId,
      jobs: jobs,
      startTime: new Date().toISOString(),
      status: 'running'
    }));
    
    // Process datasets asynchronously
    processRefreshJob(jobId);
    
    showToast(`Refreshing ${enabledDatasets.length} datasets...`, 'Info');
    
    logAction('refresh_all_datasets', {
      count: enabledDatasets.length,
      jobId: jobId
    });
  } catch (error) {
    showError('Failed to refresh datasets', error);
  }
}

/**
 * Shows settings dialog
 */
function showSettings() {
  showSidebar('settings');
  logAction('show_settings');
}

/**
 * Shows about dialog
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  
  const message = `QuickBooks Online Connector for Google Sheets
  
Version: ${SCRIPT_VERSION}
  
This add-on allows you to import data from QuickBooks Online into Google Sheets.
  
Features:
• OAuth 2.0 authentication
• Standard reports (P&L, Balance Sheet, etc.)
• Custom QBO queries
• Scheduled refreshes
• Comprehensive logging
  
For support, visit: https://github.com/brklyngg/QBO-Sheets-Connector`;
  
  ui.alert('About', message, ui.ButtonSet.OK);
  
  logAction('show_about');
}

/**
 * Shows the current deployed version in a toast
 */
function showVersionInfo() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const title = 'QuickBooks Online Connector';
    ss.toast(`Running version ${SCRIPT_VERSION}`, title, 5);

    logAction('show_version', {
      version: SCRIPT_VERSION
    });
  } catch (error) {
    console.error('Error showing version info:', error);
  }
}

/**
 * Initializes the hidden config sheet
 */
function initializeConfigSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    
    if (!configSheet) {
      configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
      configSheet.hideSheet();
      
      // Set headers
      const headers = ['Key', 'Value', 'Updated', 'Type'];
      configSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format headers
      configSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4285F4')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
      
      // Set column widths
      configSheet.setColumnWidth(1, 200);
      configSheet.setColumnWidth(2, 400);
      configSheet.setColumnWidth(3, 150);
      configSheet.setColumnWidth(4, 100);
      
      // Initialize default values
      const defaults = [
        ['version', SCRIPT_VERSION, new Date().toISOString(), 'system'],
        ['minorVersion', DEFAULT_MINOR_VERSION, new Date().toISOString(), 'config'],
        ['maxCellsWarning', '2000000', new Date().toISOString(), 'config'],
        ['maxCellsHard', '8000000', new Date().toISOString(), 'config'],
        ['logBatchSize', '100', new Date().toISOString(), 'config'],
        ['logRotationRows', '50000', new Date().toISOString(), 'config']
      ];
      
      if (defaults.length > 0) {
        configSheet.getRange(2, 1, defaults.length, 4).setValues(defaults);
      }
    }
    
    return configSheet;
  } catch (error) {
    console.error('Error initializing config sheet:', error);
    throw error;
  }
}

/**
 * Initializes the logs sheet
 */
function initializeLogsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logsSheet = ss.getSheetByName(LOGS_SHEET_NAME);
    
    if (!logsSheet) {
      logsSheet = ss.insertSheet(LOGS_SHEET_NAME);
      
      // Set headers
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
      
      // Format headers
      logsSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#34A853')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
      
      // Freeze header row
      logsSheet.setFrozenRows(1);
      
      // Set column widths
      logsSheet.setColumnWidth(1, 180); // ts_iso
      logsSheet.setColumnWidth(2, 200); // user_email
      logsSheet.setColumnWidth(4, 150); // action
      logsSheet.setColumnWidth(28, 300); // error_message
    }
    
    return logsSheet;
  } catch (error) {
    console.error('Error initializing logs sheet:', error);
    throw error;
  }
}

/**
 * Shows an error message to the user
 */
function showError(message, error) {
  console.error(message, error);
  
  const ui = SpreadsheetApp.getUi();
  const details = error ? `\n\nDetails: ${error.toString()}` : '';
  
  ui.alert('Error', message + details, ui.ButtonSet.OK);
  
  // Log the error
  logAction('error', {
    message: message,
    error: error ? error.toString() : null,
    stack: error && error.stack ? error.stack : null
  });
}

/**
 * Shows a toast notification
 */
function showToast(message, title) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title || ADDON_TITLE, 5);
}

/**
 * Gets configuration value
 */
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

/**
 * Sets configuration value
 */
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
        configSheet.getRange(i + 1, 2, 1, 3).setValues([[
          value,
          new Date().toISOString(),
          type
        ]]);
        found = true;
        break;
      }
    }
    
    if (!found) {
      const newRow = configSheet.getLastRow() + 1;
      configSheet.getRange(newRow, 1, 1, 4).setValues([[
        key,
        value,
        new Date().toISOString(),
        type
      ]]);
    }
  } catch (error) {
    console.error('Error setting config:', error);
    throw error;
  }
}

/**
 * Utility function to include HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Processes a refresh job asynchronously
 */
function processRefreshJob(jobId) {
  const userProperties = PropertiesService.getUserProperties();
  const jobData = JSON.parse(userProperties.getProperty('refreshJob_' + jobId));
  
  if (!jobData) {
    console.error('Job not found:', jobId);
    return;
  }
  
  // Process each dataset
  for (let i = 0; i < jobData.jobs.length; i++) {
    const job = jobData.jobs[i];
    
    if (job.status === 'queued') {
      try {
        job.status = 'running';
        job.startTime = new Date().toISOString();
        
        // Run the dataset
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
  
  // Update job status
  jobData.status = jobData.jobs.every(j => j.status === 'completed') ? 'completed' : 
                   jobData.jobs.some(j => j.status === 'error') ? 'error' : 'running';
  jobData.endTime = new Date().toISOString();
  
  userProperties.setProperty('refreshJob_' + jobId, JSON.stringify(jobData));
}

/**
 * Gets the current script URL for OAuth callback
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Gets the script ID
 */
function getScriptId() {
  return ScriptApp.getScriptId();
}
