/**
 * Comprehensive Logging System
 * Handles all logging operations with batching and rotation
 */

// Logging configuration
const LOG_BATCH_SIZE = 100;
const LOG_ROTATION_ROWS = 50000;
const LOG_BATCH_KEY = 'log_batch';
const LOG_BATCH_FLUSH_INTERVAL = 10000; // 10 seconds

// Log entry queue
let logQueue = [];
let lastFlushTime = new Date().getTime();

/**
 * Main logging function
 */
function logAction(action, data = {}) {
  try {
    const logEntry = {
      ts_iso: new Date().toISOString(),
      user_email: Session.getActiveUser().getEmail() || 'unknown',
      schedule_owner_email: data.schedule_owner_email || null,
      action: action,
      dataset_type: data.dataset_type || null,
      dataset_id: data.dataset_id || null,
      report_name: data.report_name || null,
      params_json: data.params_json || null,
      start_date: data.start_date || null,
      end_date: data.end_date || null,
      method: data.method || null,
      query_select: data.query_select || null,
      query_from: data.query_from || null,
      query_where: data.query_where || null,
      query_orderby: data.query_orderby || null,
      query_startposition: data.query_startposition || null,
      query_maxresults: data.query_maxresults || null,
      realmId: data.realmId || PropertiesService.getUserProperties().getProperty('QBO_REALM_ID'),
      minorversion: data.minorversion || getConfig('minorVersion', DEFAULT_MINOR_VERSION),
      target_sheet: data.target_sheet || null,
      target_named_range: data.target_named_range || null,
      rows: data.rows !== undefined ? data.rows : null,
      cols: data.cols !== undefined ? data.cols : null,
      elapsed_ms: data.elapsed_ms !== undefined ? data.elapsed_ms : null,
      status: data.status || 'info',
      http_status: data.http_status !== undefined ? data.http_status : null,
      error_code: data.error_code || null,
      error_message: data.error_message || null,
      intuit_tid: data.intuit_tid || null,
      retry_after: data.retry_after || null,
      schedule_id: data.schedule_id || null,
      schedule_freq: data.schedule_freq || null,
      request_id: data.request_id || Utilities.getUuid(),
      job_id: data.job_id || null,
      dataset_version: data.dataset_version || null,
      qbo_endpoint: data.qbo_endpoint || null,
      response_bytes: data.response_bytes !== undefined ? data.response_bytes : null,
      transport: data.transport || null
    };
    
    // Add to queue
    addToLogQueue(logEntry);
    
    // Check if we should flush
    const currentTime = new Date().getTime();
    const shouldFlush = logQueue.length >= LOG_BATCH_SIZE || 
                       (currentTime - lastFlushTime) > LOG_BATCH_FLUSH_INTERVAL;
    
    if (shouldFlush) {
      flushLogQueue();
    }
  } catch (error) {
    console.error('Error in logAction:', error);
  }
}

/**
 * Adds a log entry to the queue
 */
function addToLogQueue(logEntry) {
  logQueue.push(logEntry);
  
  // Also store in properties as backup
  try {
    const userProperties = PropertiesService.getUserProperties();
    const existingBatch = userProperties.getProperty(LOG_BATCH_KEY);
    const batch = existingBatch ? JSON.parse(existingBatch) : [];
    batch.push(logEntry);
    
    // Keep only last N entries in properties
    if (batch.length > LOG_BATCH_SIZE * 2) {
      batch.splice(0, batch.length - LOG_BATCH_SIZE);
    }
    
    userProperties.setProperty(LOG_BATCH_KEY, JSON.stringify(batch));
  } catch (error) {
    console.error('Error storing log batch:', error);
  }
}

/**
 * Flushes the log queue to the sheet
 */
function flushLogQueue() {
  try {
    if (logQueue.length === 0) {
      return;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logsSheet = ss.getSheetByName(LOGS_SHEET_NAME);
    
    if (!logsSheet) {
      logsSheet = initializeLogsSheet();
    }
    
    // Check for rotation
    const currentRows = logsSheet.getLastRow();
    if (currentRows > LOG_ROTATION_ROWS) {
      rotateLogsSheet(logsSheet);
      logsSheet = ss.getSheetByName(LOGS_SHEET_NAME);
    }
    
    // Convert log entries to rows
    const rows = logQueue.map(entry => [
      entry.ts_iso,
      entry.user_email,
      entry.schedule_owner_email,
      entry.action,
      entry.dataset_type,
      entry.dataset_id,
      entry.report_name,
      entry.params_json,
      entry.start_date,
      entry.end_date,
      entry.method,
      entry.query_select,
      entry.query_from,
      entry.query_where,
      entry.query_orderby,
      entry.query_startposition,
      entry.query_maxresults,
      entry.realmId,
      entry.minorversion,
      entry.target_sheet,
      entry.target_named_range,
      entry.rows,
      entry.cols,
      entry.elapsed_ms,
      entry.status,
      entry.http_status,
      entry.error_code,
      entry.error_message,
      entry.intuit_tid,
      entry.retry_after,
      entry.schedule_id,
      entry.schedule_freq,
      entry.request_id,
      entry.job_id,
      entry.dataset_version,
      entry.qbo_endpoint,
      entry.response_bytes,
      entry.transport
    ]);
    
    // Append to sheet
    if (rows.length > 0) {
      const startRow = logsSheet.getLastRow() + 1;
      const range = logsSheet.getRange(startRow, 1, rows.length, rows[0].length);
      range.setValues(rows);
    }
    
    // Clear queue
    logQueue = [];
    lastFlushTime = new Date().getTime();
    
    // Clear properties backup
    PropertiesService.getUserProperties().deleteProperty(LOG_BATCH_KEY);
  } catch (error) {
    console.error('Error flushing log queue:', error);
  }
}

/**
 * Forces immediate flush of log queue
 */
function forceFlushLogs() {
  flushLogQueue();
}

/**
 * Rotates the logs sheet when it gets too large
 */
function rotateLogsSheet(logsSheet) {
  try {
    const ss = logsSheet.getParent();
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
    const archiveName = `${LOGS_SHEET_NAME}_Archive_${timestamp}`;
    
    // Rename current sheet to archive
    logsSheet.setName(archiveName);
    
    // Create new logs sheet
    initializeLogsSheet();
    
    // Hide the archive sheet
    ss.getSheetByName(archiveName).hideSheet();
    
    console.log('Rotated logs sheet to:', archiveName);
  } catch (error) {
    console.error('Error rotating logs sheet:', error);
  }
}

/**
 * Gets recent logs for display
 */
function getRecentLogs(limit = 100, filters = {}) {
  try {
    // Flush any pending logs first
    flushLogQueue();
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logsSheet = ss.getSheetByName(LOGS_SHEET_NAME);
    
    if (!logsSheet) {
      return [];
    }
    
    const lastRow = logsSheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }
    
    // Get headers
    const headers = logsSheet.getRange(1, 1, 1, logsSheet.getLastColumn()).getValues()[0];
    
    // Calculate range to fetch
    const startRow = Math.max(2, lastRow - limit + 1);
    const numRows = lastRow - startRow + 1;
    
    if (numRows <= 0) {
      return [];
    }
    
    // Fetch data
    const data = logsSheet.getRange(startRow, 1, numRows, headers.length).getValues();
    
    // Convert to objects and apply filters
    const logs = [];
    for (let i = data.length - 1; i >= 0; i--) {
      const logEntry = {};
      headers.forEach((header, index) => {
        logEntry[header] = data[i][index];
      });
      
      // Apply filters
      if (filters.action && logEntry.action !== filters.action) continue;
      if (filters.status && logEntry.status !== filters.status) continue;
      if (filters.dataset_id && logEntry.dataset_id !== filters.dataset_id) continue;
      if (filters.startDate && new Date(logEntry.ts_iso) < new Date(filters.startDate)) continue;
      if (filters.endDate && new Date(logEntry.ts_iso) > new Date(filters.endDate)) continue;
      
      logs.push(logEntry);
      
      if (logs.length >= limit) {
        break;
      }
    }
    
    return logs;
  } catch (error) {
    console.error('Error getting recent logs:', error);
    return [];
  }
}

/**
 * Exports logs to CSV
 */
function exportLogs(filters = {}) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logsSheet = ss.getSheetByName(LOGS_SHEET_NAME);
    
    if (!logsSheet) {
      throw new Error('No logs found');
    }
    
    // Get all data
    const dataRange = logsSheet.getDataRange();
    const data = dataRange.getValues();
    
    if (data.length <= 1) {
      throw new Error('No log entries found');
    }
    
    // Convert to CSV
    const csv = data.map(row => {
      return row.map(cell => {
        // Escape quotes and wrap in quotes if contains comma or newline
        const value = cell ? cell.toString() : '';
        if (value.includes(',') || value.includes('\n') || value.includes('"')) {
          return '"' + value.replace(/"/g, '""') + '"';
        }
        return value;
      }).join(',');
    }).join('\n');
    
    // Create blob
    const blob = Utilities.newBlob(csv, 'text/csv', 'qbo_connector_logs.csv');
    
    return {
      success: true,
      data: Utilities.base64Encode(blob.getBytes()),
      filename: 'qbo_connector_logs.csv',
      mimeType: 'text/csv'
    };
  } catch (error) {
    console.error('Error exporting logs:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Clears old logs
 */
function clearOldLogs(daysToKeep = 30) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logsSheet = ss.getSheetByName(LOGS_SHEET_NAME);
    
    if (!logsSheet || logsSheet.getLastRow() <= 1) {
      return 0;
    }
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
    
    // Get timestamp column (first column)
    const timestamps = logsSheet.getRange(2, 1, logsSheet.getLastRow() - 1, 1).getValues();
    
    // Find first row to keep
    let firstRowToKeep = -1;
    for (let i = 0; i < timestamps.length; i++) {
      const timestamp = new Date(timestamps[i][0]);
      if (timestamp >= cutoffDate) {
        firstRowToKeep = i + 2; // +2 because of header row and 0-based index
        break;
      }
    }
    
    if (firstRowToKeep > 2) {
      // Delete old rows
      logsSheet.deleteRows(2, firstRowToKeep - 2);
      
      const deletedCount = firstRowToKeep - 2;
      logAction('clear_old_logs', {
        days_to_keep: daysToKeep,
        rows_deleted: deletedCount
      });
      
      return deletedCount;
    }
    
    return 0;
  } catch (error) {
    console.error('Error clearing old logs:', error);
    return 0;
  }
}

/**
 * Gets log statistics
 */
function getLogStats() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logsSheet = ss.getSheetByName(LOGS_SHEET_NAME);
    
    if (!logsSheet || logsSheet.getLastRow() <= 1) {
      return {
        totalLogs: 0,
        oldestLog: null,
        newestLog: null,
        byAction: {},
        byStatus: {},
        errors: 0
      };
    }
    
    const data = logsSheet.getDataRange().getValues();
    const headers = data[0];
    const actionIndex = headers.indexOf('action');
    const statusIndex = headers.indexOf('status');
    const timestampIndex = headers.indexOf('ts_iso');
    
    const stats = {
      totalLogs: data.length - 1,
      oldestLog: data[1][timestampIndex],
      newestLog: data[data.length - 1][timestampIndex],
      byAction: {},
      byStatus: {},
      errors: 0
    };
    
    // Count by action and status
    for (let i = 1; i < data.length; i++) {
      const action = data[i][actionIndex];
      const status = data[i][statusIndex];
      
      stats.byAction[action] = (stats.byAction[action] || 0) + 1;
      stats.byStatus[status] = (stats.byStatus[status] || 0) + 1;
      
      if (status === 'error') {
        stats.errors++;
      }
    }
    
    return stats;
  } catch (error) {
    console.error('Error getting log stats:', error);
    return {
      error: error.toString()
    };
  }
}

/**
 * Creates a trigger to periodically flush logs
 */
function setupLogFlushTrigger() {
  try {
    // Remove existing flush triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'flushLogQueue') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create new trigger to flush every minute
    ScriptApp.newTrigger('flushLogQueue')
      .timeBased()
      .everyMinutes(1)
      .create();
      
    console.log('Log flush trigger created');
  } catch (error) {
    console.error('Error setting up log flush trigger:', error);
  }
}

function ensureLogFlushTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const hasTrigger = triggers.some(trigger => trigger.getHandlerFunction() === 'flushLogQueue');
    if (!hasTrigger) {
      setupLogFlushTrigger();
      logAction('log_flush_trigger_created', { reason: 'ensure' });
    }
  } catch (error) {
    console.error('Error ensuring log flush trigger:', error);
    logAction('log_flush_trigger_error', { error: error.toString() });
  }
}

/**
 * Gets a formatted error message for display
 */
function formatErrorForDisplay(error) {
  if (!error) return 'Unknown error';
  
  if (typeof error === 'string') {
    return error;
  }
  
  if (error.message) {
    return error.message;
  }
  
  return error.toString();
}

/**
 * Creates a detailed error log entry
 */
function logError(action, error, context = {}) {
  const errorData = {
    ...context,
    status: 'error',
    error_message: formatErrorForDisplay(error),
    error_stack: error.stack || null,
    error_type: error.name || 'Error'
  };
  
  logAction(action, errorData);
}

/**
 * Logs API request details
 */
function logApiRequest(method, endpoint, params = {}) {
  logAction('api_request', {
    method: method,
    qbo_endpoint: endpoint,
    params_json: JSON.stringify(params)
  });
}

/**
 * Logs API response details
 */
function logApiResponse(method, endpoint, response, elapsedMs) {
  const responseCode = response.getResponseCode();
  const headers = response.getHeaders();
  
  logAction('api_response', {
    method: method,
    qbo_endpoint: endpoint,
    http_status: responseCode,
    elapsed_ms: elapsedMs,
    intuit_tid: headers['intuit_tid'] || null,
    retry_after: headers['Retry-After'] || null,
    response_bytes: response.getBlob().getBytes().length,
    status: responseCode >= 200 && responseCode < 300 ? 'success' : 'error'
  });
}
