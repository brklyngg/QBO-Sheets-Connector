/**
 * Comprehensive logging utilities for the QuickBooks connector.
 * Ported from the legacy build and adapted to the new JS module layout.
 */

const LOG_BATCH_SIZE = 100;
const LOG_ROTATION_ROWS = 50000;
const LOG_BATCH_KEY = 'log_batch';
const LOG_BATCH_FLUSH_INTERVAL = 10000; // 10 seconds

let logQueue = [];
let lastFlushTime = new Date().getTime();

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
      realmId: data.realmId || getUserProps().getProperty(PROPERTY_REALM_ID),
      minorversion: data.minorversion || getConfig('minorVersion', getDefaultMinorVersion()),
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

    addToLogQueue(logEntry);

    const currentTime = new Date().getTime();
    const shouldFlush = logQueue.length >= LOG_BATCH_SIZE || (currentTime - lastFlushTime) > LOG_BATCH_FLUSH_INTERVAL;

    if (shouldFlush) {
      flushLogQueue();
    }
  } catch (error) {
    console.error('Error in logAction:', error);
  }
}

function addToLogQueue(logEntry) {
  logQueue.push(logEntry);

  try {
    const userProperties = getUserProps();
    const existingBatch = userProperties.getProperty(LOG_BATCH_KEY);
    const batch = existingBatch ? JSON.parse(existingBatch) : [];
    batch.push(logEntry);

    if (batch.length > LOG_BATCH_SIZE * 2) {
      batch.splice(0, batch.length - LOG_BATCH_SIZE);
    }

    userProperties.setProperty(LOG_BATCH_KEY, JSON.stringify(batch));
  } catch (error) {
    console.error('Error storing log batch:', error);
  }
}

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

    const currentRows = logsSheet.getLastRow();
    if (currentRows > LOG_ROTATION_ROWS) {
      rotateLogsSheet(logsSheet);
      logsSheet = ss.getSheetByName(LOGS_SHEET_NAME);
    }

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

    if (rows.length > 0) {
      const startRow = logsSheet.getLastRow() + 1;
      logsSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    }

    logQueue = [];
    lastFlushTime = new Date().getTime();
    getUserProps().deleteProperty(LOG_BATCH_KEY);
  } catch (error) {
    console.error('Error flushing log queue:', error);
  }
}

function forceFlushLogs() {
  flushLogQueue();
}

function rotateLogsSheet(logsSheet) {
  try {
    const ss = logsSheet.getParent();
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
    const archiveName = `${LOGS_SHEET_NAME}_Archive_${timestamp}`;

    logsSheet.copyTo(ss).setName(archiveName);
    logsSheet.clearContents();

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
  } catch (error) {
    console.error('Error rotating logs sheet:', error);
  }
}

function getRecentLogs(limit = 100, filters = {}) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LOGS_SHEET_NAME);
    if (!sheet) {
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const entries = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const entry = {};
      headers.forEach((header, index) => {
        entry[header] = row[index];
      });

      const matchesFilters = Object.keys(filters || {}).every(key => {
        if (!filters[key]) {
          return true;
        }
        return entry[key] === filters[key];
      });

      if (matchesFilters) {
        entries.push(entry);
      }
    }

    return entries.slice(-limit).reverse();
  } catch (error) {
    console.error('Error getting recent logs:', error);
    return [];
  }
}

function exportLogs(filters = {}) {
  try {
    const logs = getRecentLogs(1000, filters);
    const csvRows = [];

    if (logs.length > 0) {
      const headers = Object.keys(logs[0]);
      csvRows.push(headers.join(','));

      logs.forEach(log => {
        const row = headers.map(header => {
          const value = log[header];
          if (value === null || value === undefined) {
            return '';
          }
          const stringValue = String(value).replace(/"/g, '""');
          return `"${stringValue}"`;
        });
        csvRows.push(row.join(','));
      });
    }

    const csvContent = csvRows.join('\n');
    const blob = Utilities.newBlob(csvContent, 'text/csv', 'qbo_connector_logs.csv');

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

function clearOldLogs(daysToKeep = 30) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LOGS_SHEET_NAME);
    if (!sheet) {
      return 0;
    }

    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - daysToKeep);

    const data = sheet.getDataRange().getValues();
    const rowsToKeep = [data[0]];
    let removed = 0;

    for (let i = 1; i < data.length; i++) {
      const timestamp = new Date(data[i][0]);
      if (timestamp >= cutoff) {
        rowsToKeep.push(data[i]);
      } else {
        removed++;
      }
    }

    if (removed > 0) {
      sheet.clearContents();
      sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
      sheet.getRange(1, 1, 1, rowsToKeep[0].length)
        .setBackground('#34A853')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    return removed;
  } catch (error) {
    console.error('Error clearing old logs:', error);
    return 0;
  }
}

function getLogStats() {
  try {
    const logs = getRecentLogs(500);
    const totals = logs.reduce((acc, log) => {
      acc.total++;
      acc[log.status] = (acc[log.status] || 0) + 1;
      return acc;
    }, { total: 0 });

    return totals;
  } catch (error) {
    console.error('Error getting log stats:', error);
    return {
      total: 0,
      error: error.toString()
    };
  }
}

function setupLogFlushTrigger() {
  ScriptApp.newTrigger('forceFlushLogs')
    .timeBased()
    .everyMinutes(5)
    .create();
}

function ensureLogFlushTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const existing = triggers.some(trigger => trigger.getHandlerFunction() === 'forceFlushLogs');
  if (!existing) {
    setupLogFlushTrigger();
  }
}

function formatErrorForDisplay(error) {
  if (!error) {
    return 'Unknown error';
  }
  if (typeof error === 'string') {
    return error;
  }
  if (error.message) {
    return error.message;
  }
  return JSON.stringify(error);
}

function logError(action, error, context = {}) {
  logAction(action, Object.assign({}, context, {
    status: 'error',
    error_message: formatErrorForDisplay(error)
  }));
}

function logApiRequest(method, endpoint, params = {}) {
  logAction('api_request', {
    method: method,
    qbo_endpoint: endpoint,
    params_json: params ? JSON.stringify(params) : null
  });
}

function logApiResponse(method, endpoint, response, elapsedMs) {
  logAction('api_response', {
    method: method,
    qbo_endpoint: endpoint,
    http_status: response ? response.getResponseCode() : null,
    elapsed_ms: elapsedMs,
    response_bytes: response ? response.getContentText().length : null
  });
}

function getDefaultMinorVersion() {
  if (typeof QBO_DEFAULT_MINOR_VERSION !== 'undefined') {
    return QBO_DEFAULT_MINOR_VERSION;
  }
  return '75';
}
