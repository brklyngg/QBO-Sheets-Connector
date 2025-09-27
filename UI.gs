/**
 * UI Server Functions
 * Handles all server-side functions called from the HTML UI
 */

/**
 * Server functions exposed to UI via google.script.run
 * These are already defined in other files but listed here for reference:
 * 
 * From Auth.gs:
 * - getAuthorizationUrl()
 * - isConnected()
 * - getConnectionDetails()
 * - disconnect()
 * - setOAuthCredentials(clientId, clientSecret)
 * - getOAuthCredentials()
 * 
 * From Datasets.gs:
 * - getDatasets()
 * - getDatasetById(datasetId)
 * - createDataset(dataset)
 * - updateDataset(datasetId, updates)
 * - deleteDataset(datasetId)
 * - runDataset(datasetId)
 * - getJobStatus(jobId)
 * 
 * From Scheduler.gs:
 * - testScheduledRun(datasetId)
 * - getScheduleStatus()
 * 
 * From Logging.gs:
 * - getRecentLogs(limit, filters)
 * - exportLogs(filters)
 * - getLogStats()
 * 
 * From Code.gs:
 * - showToast(message, title)
 * - showLogsSheet()
 * - showSettings()
 * - showAbout()
 */

/**
 * Additional UI-specific server functions
 */

/**
 * Gets available report types for the UI dropdown
 */
function getReportTypes() {
  return getAvailableReports();
}

/**
 * Gets available entity types for query builder
 */
function getEntityTypes() {
  return getAvailableEntities();
}

/**
 * Validates a query string
 */
function validateQuery(query) {
  try {
    const result = parseQuery(query);
    if (result.valid) {
      return {
        valid: true,
        parsed: result
      };
    } else {
      return {
        valid: false,
        error: result.error
      };
    }
  } catch (error) {
    return {
      valid: false,
      error: error.toString()
    };
  }
}

/**
 * Gets sheets in the current spreadsheet
 */
function getAvailableSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    return sheets.map(sheet => ({
      id: sheet.getSheetId(),
      name: sheet.getName(),
      index: sheet.getIndex()
    }));
  } catch (error) {
    console.error('Error getting sheets:', error);
    return [];
  }
}

/**
 * Creates a new sheet with the given name
 */
function createSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet(sheetName);
    
    return {
      success: true,
      sheet: {
        id: sheet.getSheetId(),
        name: sheet.getName(),
        index: sheet.getIndex()
      }
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Gets named ranges in the spreadsheet
 */
function getNamedRanges() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const namedRanges = ss.getNamedRanges();
    
    return namedRanges.map(namedRange => ({
      name: namedRange.getName(),
      range: namedRange.getRange().getA1Notation(),
      sheet: namedRange.getRange().getSheet().getName()
    }));
  } catch (error) {
    console.error('Error getting named ranges:', error);
    return [];
  }
}

/**
 * Tests QuickBooks connection with company info
 */
function testConnection() {
  try {
    if (!isConnected()) {
      return {
        success: false,
        error: 'Not connected to QuickBooks'
      };
    }
    
    const companyInfo = fetchCompanyInfo();
    
    return {
      success: true,
      companyInfo: {
        companyName: companyInfo.CompanyName,
        legalName: companyInfo.LegalName,
        country: companyInfo.Country,
        fiscalYearStartMonth: companyInfo.FiscalYearStartMonth
      }
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Gets dataset run history
 */
function getDatasetHistory(datasetId, limit = 10) {
  try {
    const logs = getRecentLogs(limit, { dataset_id: datasetId });
    
    return logs.filter(log => 
      log.action === 'run_dataset' || 
      log.action === 'scheduled_run_complete'
    );
  } catch (error) {
    console.error('Error getting dataset history:', error);
    return [];
  }
}

/**
 * Clears dataset cache/data
 */
function clearDatasetData(datasetId) {
  try {
    const dataset = getDatasetById(datasetId);
    if (!dataset) {
      throw new Error('Dataset not found');
    }
    
    // Clear the target sheet if it exists
    if (dataset.target && dataset.target.sheetId) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = ss.getSheets();
      const sheet = sheets.find(s => s.getSheetId() === parseInt(dataset.target.sheetId));
      
      if (sheet) {
        sheet.clear();
        logAction('clear_dataset_data', {
          dataset_id: datasetId,
          dataset_name: dataset.name,
          sheet_name: sheet.getName()
        });
      }
    }
    
    return {
      success: true,
      message: 'Dataset data cleared successfully'
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Duplicates a dataset
 */
function duplicateDataset(datasetId) {
  try {
    const original = getDatasetById(datasetId);
    if (!original) {
      throw new Error('Dataset not found');
    }
    
    // Create a copy with a new name
    const duplicate = {
      ...original,
      id: undefined, // Let createDataset generate a new ID
      name: original.name + ' (Copy)',
      schedule: {
        ...original.schedule,
        enabled: false // Disable schedule for duplicate
      }
    };
    
    const created = createDataset(duplicate);
    
    logAction('duplicate_dataset', {
      original_id: datasetId,
      new_id: created.id,
      original_name: original.name,
      new_name: created.name
    });
    
    return created;
  } catch (error) {
    console.error('Error duplicating dataset:', error);
    throw error;
  }
}

/**
 * Gets system information for debugging
 */
function getSystemInfo() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const triggers = ScriptApp.getProjectTriggers();
    const properties = PropertiesService.getUserProperties();
    
    return {
      spreadsheet: {
        name: ss.getName(),
        id: ss.getId(),
        sheets: ss.getSheets().length,
        locale: ss.getSpreadsheetLocale(),
        timezone: ss.getSpreadsheetTimeZone()
      },
      script: {
        id: ScriptApp.getScriptId(),
        timezone: Session.getScriptTimeZone(),
        triggers: triggers.length,
        scheduledTriggers: triggers.filter(t => t.getHandlerFunction() === 'scheduledDatasetRun').length
      },
      user: {
        email: Session.getActiveUser().getEmail(),
        timezone: Session.getTimeZone()
      },
      quotas: {
        // These are approximate limits
        maxExecutionTime: '6 minutes',
        maxTriggers: '20 per user',
        maxPropertiesSize: '500KB total',
        maxCellsPerSpreadsheet: '10 million'
      },
      addon: {
        version: SCRIPT_VERSION,
        minorVersion: getConfig('minorVersion', DEFAULT_MINOR_VERSION)
      }
    };
  } catch (error) {
    console.error('Error getting system info:', error);
    return {
      error: error.toString()
    };
  }
}

/**
 * Exports dataset configuration
 */
function exportDatasetConfig(datasetId) {
  try {
    const dataset = getDatasetById(datasetId);
    if (!dataset) {
      throw new Error('Dataset not found');
    }
    
    // Remove internal fields
    const exportData = {
      ...dataset,
      id: undefined,
      created: undefined,
      updated: undefined,
      version: undefined,
      lastWrite: undefined
    };
    
    const json = JSON.stringify(exportData, null, 2);
    const blob = Utilities.newBlob(json, 'application/json', `${dataset.name}_config.json`);
    
    return {
      success: true,
      data: Utilities.base64Encode(blob.getBytes()),
      filename: `${dataset.name}_config.json`,
      mimeType: 'application/json'
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Imports dataset configuration
 */
function importDatasetConfig(configJson, newName) {
  try {
    const config = JSON.parse(configJson);
    config.target = normalizeDatasetTarget(config.target || {}, config.name);
    config.pagination = normalizeDatasetPagination(config.pagination);

    // Validate the configuration
    const validation = validateDataset(config);
    if (!validation.valid) {
      throw new Error(validation.error);
    }
    
    // Create new dataset with imported config
    const dataset = {
      ...config,
      name: newName || config.name + ' (Imported)',
      schedule: {
        ...config.schedule,
        enabled: false // Disable schedule for imported dataset
      }
    };
    
    const created = createDataset(dataset);
    
    logAction('import_dataset_config', {
      dataset_id: created.id,
      dataset_name: created.name,
      source_name: config.name
    });
    
    return {
      success: true,
      dataset: created
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Gets usage statistics
 */
function getUsageStats() {
  try {
    const datasets = getDatasets();
    const logs = getLogStats();
    const scheduleStatus = getScheduleStatus();
    
    return {
      datasets: {
        total: datasets.length,
        standard: datasets.filter(d => d.type === 'standard').length,
        query: datasets.filter(d => d.type === 'query').length,
        scheduled: datasets.filter(d => d.schedule && d.schedule.enabled).length
      },
      logs: logs,
      schedules: scheduleStatus,
      lastRun: datasets
        .filter(d => d.lastWrite)
        .sort((a, b) => new Date(b.lastWrite.wroteAt) - new Date(a.lastWrite.wroteAt))[0]?.lastWrite?.wroteAt
    };
  } catch (error) {
    console.error('Error getting usage stats:', error);
    return {
      error: error.toString()
    };
  }
}

/**
 * Performs maintenance tasks
 */
function performMaintenance() {
  try {
    const results = {
      orphanedTriggers: cleanupOrphanedTriggers(),
      oldLogs: clearOldLogs(30),
      logFlush: false
    };
    
    // Force flush any pending logs
    try {
      forceFlushLogs();
      results.logFlush = true;
    } catch (e) {
      console.error('Error flushing logs:', e);
    }
    
    logAction('perform_maintenance', results);
    
    return {
      success: true,
      results: results
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Gets persisted add-on settings
 */
function getAddonSettings() {
  try {
    const props = PropertiesService.getUserProperties();
    return {
      success: true,
      useSandbox: props.getProperty('QBO_USE_SANDBOX') === 'true',
      enableNotifications: props.getProperty('SCHEDULE_ERROR_NOTIFICATIONS') === 'true'
    };
  } catch (error) {
    console.error('Error getting add-on settings:', error);
    return {
      success: false,
      error: error.toString(),
      useSandbox: false,
      enableNotifications: false
    };
  }
}

/**
 * Saves add-on settings
 */
function saveAddonSettings(settings) {
  try {
    if (!settings || typeof settings !== 'object') {
      throw new Error('Invalid settings payload');
    }

    const props = PropertiesService.getUserProperties();

    if (settings.useSandbox !== undefined) {
      props.setProperty('QBO_USE_SANDBOX', settings.useSandbox ? 'true' : 'false');
    }

    if (settings.enableNotifications !== undefined) {
      props.setProperty('SCHEDULE_ERROR_NOTIFICATIONS', settings.enableNotifications ? 'true' : 'false');
    }

    logAction('save_addon_settings', {
      useSandbox: settings.useSandbox === true,
      enableNotifications: settings.enableNotifications === true
    });

    return {
      success: true
    };
  } catch (error) {
    console.error('Error saving add-on settings:', error);
    logAction('save_addon_settings_error', {
      error: error.toString()
    });

    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Gets help content for a specific topic
 */
function getHelpContent(topic) {
  const helpContent = {
    'oauth': {
      title: 'Setting up OAuth',
      content: `
1. Go to https://developer.intuit.com
2. Create a new app or use existing
3. Add redirect URI: ${getScriptUrl()}
4. Copy Client ID and Client Secret
5. Paste them in Settings tab
6. Click Connect to authorize
      `
    },
    'datasets': {
      title: 'Creating Datasets',
      content: `
Standard Reports:
- Choose from 5 pre-built reports
- Set date ranges or use presets
- Data updates automatically

Custom Queries:
- Write SQL-like queries
- Example: SELECT * FROM Customer
- Max 1000 results per run
      `
    },
    'schedules': {
      title: 'Scheduling Refreshes',
      content: `
- Enable schedule on any dataset
- Choose frequency: hourly, daily, weekly, monthly
- Set specific time for non-hourly
- Runs automatically in background
- Check logs for run history
      `
    },
    'troubleshooting': {
      title: 'Troubleshooting',
      content: `
Common issues:
- 401 Error: Reconnect to QuickBooks
- 429 Error: Rate limit, wait and retry
- No data: Check date ranges and filters
- Schedule not running: Check timezone settings
      `
    }
  };
  
  return helpContent[topic] || {
    title: 'Help',
    content: 'Topic not found. Please contact support.'
  };
}
