/**
 * Dataset Management
 * Handles CRUD operations for datasets and job execution
 */

// Dataset storage key prefix
const DATASET_PREFIX = 'dataset_';
const DATASETS_INDEX_KEY = 'datasets_index';

/**
 * Dataset model structure
 * {
 *   id: string,
 *   type: 'standard' | 'query',
 *   name: string,
 *   params: object,
 *   target: {sheetId: string, anchorA1: string, allowResize: boolean},
 *   pagination: {startPosition: number, maxResults: number},
 *   schedule: {enabled: boolean, freq: string, timeOfDay: string},
 *   lastWrite: {rows: number, cols: number, wroteAt: string, sheetId: string, rangeA1: string, schemaHash: string}
 * }
 */

/**
 * Gets all datasets
 */
function getDatasets() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const indexData = userProperties.getProperty(DATASETS_INDEX_KEY);
    
    if (!indexData) {
      return [];
    }
    
    const datasetIds = JSON.parse(indexData);
    const datasets = [];
    
    datasetIds.forEach(id => {
      const datasetData = userProperties.getProperty(DATASET_PREFIX + id);
      if (datasetData) {
        try {
          datasets.push(JSON.parse(datasetData));
        } catch (e) {
          console.error('Error parsing dataset:', id, e);
        }
      }
    });
    
    return datasets;
  } catch (error) {
    console.error('Error getting datasets:', error);
    return [];
  }
}

/**
 * Gets a dataset by ID
 */
function getDatasetById(datasetId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const datasetData = userProperties.getProperty(DATASET_PREFIX + datasetId);
    
    if (!datasetData) {
      return null;
    }
    
    return JSON.parse(datasetData);
  } catch (error) {
    console.error('Error getting dataset:', error);
    return null;
  }
}

/**
 * Creates a new dataset
 */
function createDataset(dataset) {
  try {
    dataset.target = normalizeDatasetTarget(dataset.target || {}, dataset.name);

    // Validate dataset
    const validation = validateDataset(dataset);
    if (!validation.valid) {
      throw new Error(validation.error);
    }
    
    // Generate ID if not provided
    if (!dataset.id) {
      dataset.id = Utilities.getUuid();
    }
    
    // Set defaults
    dataset.created = new Date().toISOString();
    dataset.updated = new Date().toISOString();
    dataset.version = 1;

    if (!dataset.pagination) {
      dataset.pagination = {
        startPosition: 1,
        maxResults: 1000
      };
    }
    
    if (!dataset.schedule) {
      dataset.schedule = {
        enabled: false,
        freq: 'daily',
        timeOfDay: '09:00'
      };
    }
    
    // Save dataset
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty(DATASET_PREFIX + dataset.id, JSON.stringify(dataset));
    
    // Update index
    updateDatasetsIndex(dataset.id, 'add');
    
    logAction('create_dataset', {
      dataset_id: dataset.id,
      dataset_type: dataset.type,
      name: dataset.name
    });
    
    return dataset;
  } catch (error) {
    console.error('Error creating dataset:', error);
    throw error;
  }
}

/**
 * Updates an existing dataset
 */
function updateDataset(datasetId, updates) {
  try {
    const existing = getDatasetById(datasetId);
    if (!existing) {
      throw new Error('Dataset not found');
    }
    
    // Merge updates
    const updated = Object.assign({}, existing, updates);
    updated.updated = new Date().toISOString();
    updated.version = (existing.version || 1) + 1;

    updated.target = normalizeDatasetTarget(updated.target || existing.target || {}, updated.name);

    // Validate updated dataset
    const validation = validateDataset(updated);
    if (!validation.valid) {
      throw new Error(validation.error);
    }
    
    // Save updated dataset
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty(DATASET_PREFIX + datasetId, JSON.stringify(updated));
    
    logAction('update_dataset', {
      dataset_id: datasetId,
      dataset_type: updated.type,
      name: updated.name,
      version: updated.version
    });
    
    // Update schedule if needed
    if (updated.schedule && updated.schedule.enabled !== existing.schedule.enabled) {
      if (updated.schedule.enabled) {
        createScheduleTrigger(datasetId, updated.schedule);
      } else {
        removeScheduleTrigger(datasetId);
      }
    }
    
    return updated;
  } catch (error) {
    console.error('Error updating dataset:', error);
    throw error;
  }
}

/**
 * Deletes a dataset
 */
function deleteDataset(datasetId) {
  try {
    const dataset = getDatasetById(datasetId);
    if (!dataset) {
      throw new Error('Dataset not found');
    }
    
    // Remove schedule trigger if exists
    if (dataset.schedule && dataset.schedule.enabled) {
      removeScheduleTrigger(datasetId);
    }
    
    // Delete dataset
    const userProperties = PropertiesService.getUserProperties();
    userProperties.deleteProperty(DATASET_PREFIX + datasetId);
    
    // Update index
    updateDatasetsIndex(datasetId, 'remove');
    
    logAction('delete_dataset', {
      dataset_id: datasetId,
      dataset_type: dataset.type,
      name: dataset.name
    });
    
    return true;
  } catch (error) {
    console.error('Error deleting dataset:', error);
    throw error;
  }
}

/**
 * Runs a dataset (async with job tracking)
 */
function runDataset(datasetId) {
  try {
    const dataset = getDatasetById(datasetId);
    if (!dataset) {
      throw new Error('Dataset not found');
    }
    
    // Create job
    const jobId = Utilities.getUuid();
    const job = {
      id: jobId,
      datasetId: datasetId,
      datasetName: dataset.name,
      status: 'running',
      startTime: new Date().toISOString(),
      progress: 0,
      message: 'Initializing...'
    };
    
    // Store job
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('job_' + jobId, JSON.stringify(job));
    
    // Set global job status for UI updates
    globalJobStatus = job;
    
    try {
      // Update job progress
      updateJobProgress(jobId, 10, 'Connecting to QuickBooks...');
      
      // Execute based on dataset type
      let result;
      if (dataset.type === 'standard') {
        result = runStandardReportDataset(dataset, jobId);
      } else if (dataset.type === 'query') {
        result = runQueryDataset(dataset, jobId);
      } else {
        throw new Error(`Unknown dataset type: ${dataset.type}`);
      }
      
      // Update job as completed
      job.status = 'completed';
      job.endTime = new Date().toISOString();
      job.progress = 100;
      job.message = 'Dataset run completed successfully';
      job.result = result;
      
      userProperties.setProperty('job_' + jobId, JSON.stringify(job));
      globalJobStatus = job;
      
      // Update dataset metadata
      const updates = {
        lastWrite: result.lastWrite
      };

      if (result.targetUpdates) {
        updates.target = Object.assign({}, dataset.target || {}, result.targetUpdates);
      }

      updateDataset(datasetId, updates);
      
      return job;
    } catch (error) {
      // Update job as failed
      job.status = 'failed';
      job.endTime = new Date().toISOString();
      job.error = error.toString();
      job.message = 'Dataset run failed: ' + error.toString();
      
      userProperties.setProperty('job_' + jobId, JSON.stringify(job));
      globalJobStatus = job;
      
      throw error;
    }
  } catch (error) {
    console.error('Error running dataset:', error);
    throw error;
  }
}

/**
 * Runs a standard report dataset
 */
function runStandardReportDataset(dataset, jobId) {
  try {
    updateJobProgress(jobId, 20, 'Fetching report from QuickBooks...');
    
    // Run the report
    const reportResult = runStandardReport(dataset.params.reportType, dataset.params);
    
    if (!reportResult.success) {
      throw new Error(reportResult.error || 'Failed to fetch report');
    }
    
    updateJobProgress(jobId, 50, 'Processing report data...');
    
    // Convert to sheet data
    const sheetData = convertReportToSheetData(reportResult.data);
    
    updateJobProgress(jobId, 70, 'Writing to spreadsheet...');
    
    // Write to sheet
    const writeResult = writeDataToSheet(dataset, sheetData.data);
    
    updateJobProgress(jobId, 90, 'Finalizing...');

    // Log success
    logAction('run_dataset', {
      dataset_id: dataset.id,
      dataset_type: dataset.type,
      report_name: dataset.params.reportType,
      rows: sheetData.rows,
      cols: sheetData.cols,
      status: 'success',
      job_id: jobId,
      intuit_tid: reportResult.intuitTid
    });
    
    return {
      success: true,
      lastWrite: {
        rows: sheetData.rows,
        cols: sheetData.cols,
        wroteAt: new Date().toISOString(),
        sheetId: writeResult.sheetId,
        sheetName: writeResult.sheetName,
        rangeA1: writeResult.rangeA1,
        schemaHash: generateSchemaHash(sheetData.data[0] || [])
      },
      targetUpdates: writeResult.targetUpdates
    };
  } catch (error) {
    logAction('run_dataset', {
      dataset_id: dataset.id,
      dataset_type: dataset.type,
      report_name: dataset.params.reportType,
      status: 'error',
      error_message: error.toString(),
      job_id: jobId
    });
    
    throw error;
  }
}

/**
 * Runs a query dataset
 */
function runQueryDataset(dataset, jobId) {
  try {
    updateJobProgress(jobId, 20, 'Executing query...');
    
    // Run the query
    const queryResult = runCustomQuery(
      dataset.params.query,
      dataset.pagination.startPosition,
      dataset.pagination.maxResults
    );
    
    if (!queryResult.success) {
      throw new Error(queryResult.error || 'Failed to execute query');
    }
    
    updateJobProgress(jobId, 50, 'Processing query results...');
    
    // Parse entity type from query
    const parsedQuery = parseQuery(dataset.params.query);
    const entityType = parsedQuery.from;
    
    // Convert to sheet data
    const sheetData = convertEntitiesToSheetData(queryResult.data, entityType);
    
    updateJobProgress(jobId, 70, 'Writing to spreadsheet...');
    
    // Write to sheet
    const writeResult = writeDataToSheet(dataset, sheetData.data);
    
    updateJobProgress(jobId, 90, 'Finalizing...');
    
    // Log success
    logAction('run_dataset', {
      dataset_id: dataset.id,
      dataset_type: dataset.type,
      query_from: entityType,
      rows: sheetData.rows,
      cols: sheetData.cols,
      status: 'success',
      job_id: jobId,
      intuit_tid: queryResult.intuitTid,
      total_count: queryResult.totalCount
    });
    
    return {
      success: true,
      lastWrite: {
        rows: sheetData.rows,
        cols: sheetData.cols,
        wroteAt: new Date().toISOString(),
        sheetId: writeResult.sheetId,
        sheetName: writeResult.sheetName,
        rangeA1: writeResult.rangeA1,
        schemaHash: generateSchemaHash(sheetData.data[0] || [])
      },
      targetUpdates: writeResult.targetUpdates
    };
  } catch (error) {
    logAction('run_dataset', {
      dataset_id: dataset.id,
      dataset_type: dataset.type,
      status: 'error',
      error_message: error.toString(),
      job_id: jobId
    });
    
    throw error;
  }
}

/**
 * Writes data to sheet based on dataset target configuration
 */
function writeDataToSheet(dataset, data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const originalTarget = dataset.target || {};
    const target = normalizeDatasetTarget(originalTarget, dataset.name);
    const targetUpdates = {};

    let sheet = null;

    const sheetId = target.sheetId ? parseInt(target.sheetId, 10) : null;
    if (sheetId) {
      sheet = ss.getSheets().find(s => s.getSheetId() === sheetId) || null;
    }

    if (!sheet && target.sheetName) {
      sheet = ss.getSheetByName(target.sheetName) || null;
    }

    if (!sheet) {
      const newSheetName = ensureUniqueSheetName(ss, target.sheetName || dataset.name || 'QBO Data');
      sheet = ss.insertSheet(newSheetName);
      targetUpdates.sheetName = sheet.getName();
      targetUpdates.sheetId = sheet.getSheetId().toString();
    } else {
      const currentSheetId = sheet.getSheetId().toString();
      if (!target.sheetId || target.sheetId !== currentSheetId) {
        targetUpdates.sheetId = currentSheetId;
      }
      if (!target.sheetName || target.sheetName !== sheet.getName()) {
        targetUpdates.sheetName = sheet.getName();
      }
    }

    const originalAnchor = originalTarget.anchorA1 ? originalTarget.anchorA1.toString().trim().toUpperCase() : '';
    const anchorA1 = sanitizeAnchorCell(target.anchorA1);
    if (originalAnchor !== anchorA1) {
      targetUpdates.anchorA1 = anchorA1;
    }

    if (typeof originalTarget.allowResize !== 'boolean') {
      targetUpdates.allowResize = target.allowResize;
    } else if (originalTarget.allowResize !== target.allowResize) {
      targetUpdates.allowResize = target.allowResize;
    }

    if ((originalTarget.namedRange || '') !== target.namedRange) {
      targetUpdates.namedRange = target.namedRange;
    }

    const anchorCell = sheet.getRange(anchorA1);
    const startRow = anchorCell.getRow();
    const startCol = anchorCell.getColumn();

    // Clear previous write range if available
    if (dataset.lastWrite && dataset.lastWrite.sheetId && dataset.lastWrite.rangeA1) {
      const lastSheetId = dataset.lastWrite.sheetId.toString();
      if (lastSheetId === sheet.getSheetId().toString()) {
        try {
          sheet.getRange(dataset.lastWrite.rangeA1).clearContent();
        } catch (clearError) {
          console.warn('Failed to clear previous range', clearError);
        }
      }
    }

    // Check data size limits
    const totalCells = data.length * (data[0] ? data[0].length : 0);
    const maxCellsWarning = parseInt(getConfig('maxCellsWarning', '2000000'));
    const maxCellsHard = parseInt(getConfig('maxCellsHard', '8000000'));
    
    if (totalCells > maxCellsHard) {
      throw new Error(`Data exceeds hard limit of ${maxCellsHard} cells (${totalCells} cells)`);
    }
    
    if (totalCells > maxCellsWarning) {
      console.warn(`Data exceeds warning limit of ${maxCellsWarning} cells (${totalCells} cells)`);
    }
    
    // Write data
    if (data.length > 0 && data[0].length > 0) {
      const numRows = data.length;
      const numCols = data[0].length;
      const range = sheet.getRange(startRow, startCol, numRows, numCols);
      range.setValues(data);

      // Format header row
      if (startRow === 1 || anchorA1 === 'A1') {
        const headerRange = sheet.getRange(startRow, startCol, 1, numCols);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#f0f0f0');
      }

      // Auto-resize columns if allowed
      if (target.allowResize) {
        for (let i = 0; i < numCols; i++) {
          sheet.autoResizeColumn(startCol + i);
        }
      }

      // Create or update named range
      const namedRange = target.namedRange;
      if (namedRange) {
        updateNamedRange(sheet, namedRange, range);
      }

      return {
        sheetId: sheet.getSheetId().toString(),
        sheetName: sheet.getName(),
        rangeA1: range.getA1Notation(),
        rows: numRows,
        cols: numCols,
        targetUpdates: Object.keys(targetUpdates).length ? targetUpdates : null
      };
    } else {
      // No data to write
      return {
        sheetId: sheet.getSheetId().toString(),
        sheetName: sheet.getName(),
        rangeA1: anchorA1,
        rows: 0,
        cols: 0,
        targetUpdates: Object.keys(targetUpdates).length ? targetUpdates : null
      };
    }
  } catch (error) {
    console.error('Error writing data to sheet:', error);
    throw error;
  }
}

/**
 * Updates datasets index
 */
function updateDatasetsIndex(datasetId, action) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    let indexData = userProperties.getProperty(DATASETS_INDEX_KEY);
    let datasetIds = indexData ? JSON.parse(indexData) : [];
    
    if (action === 'add') {
      if (!datasetIds.includes(datasetId)) {
        datasetIds.push(datasetId);
      }
    } else if (action === 'remove') {
      datasetIds = datasetIds.filter(id => id !== datasetId);
    }
    
    userProperties.setProperty(DATASETS_INDEX_KEY, JSON.stringify(datasetIds));
  } catch (error) {
    console.error('Error updating datasets index:', error);
    throw error;
  }
}

/**
 * Validates a dataset object
 */
function validateDataset(dataset) {
  if (!dataset) {
    return { valid: false, error: 'Dataset is required' };
  }
  
  if (!dataset.type || !['standard', 'query'].includes(dataset.type)) {
    return { valid: false, error: 'Invalid dataset type. Must be "standard" or "query"' };
  }
  
  if (!dataset.name || dataset.name.trim() === '') {
    return { valid: false, error: 'Dataset name is required' };
  }
  
  if (!dataset.params) {
    return { valid: false, error: 'Dataset parameters are required' };
  }
  
  if (dataset.type === 'standard') {
    if (!dataset.params.reportType) {
      return { valid: false, error: 'Report type is required for standard datasets' };
    }
    
    if (!SUPPORTED_REPORTS[dataset.params.reportType]) {
      return { valid: false, error: `Unsupported report type: ${dataset.params.reportType}` };
    }
  } else if (dataset.type === 'query') {
    if (!dataset.params.query) {
      return { valid: false, error: 'Query is required for query datasets' };
    }
    
    const parsedQuery = parseQuery(dataset.params.query);
    if (!parsedQuery.valid) {
      return { valid: false, error: `Invalid query: ${parsedQuery.error}` };
    }
  }
  
  if (!dataset.target) {
    return { valid: false, error: 'Target configuration is required' };
  }

  if (!dataset.target.sheetName || dataset.target.sheetName.toString().trim() === '') {
    return { valid: false, error: 'Target sheet name is required' };
  }

  if (!dataset.target.anchorA1 || !/^[A-Z]+[1-9][0-9]*$/.test(dataset.target.anchorA1)) {
    return { valid: false, error: 'Invalid anchor cell. Use A1 notation.' };
  }
  
  return { valid: true };
}

/**
 * Updates job progress
 */
function updateJobProgress(jobId, progress, message) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const jobData = userProperties.getProperty('job_' + jobId);
    
    if (jobData) {
      const job = JSON.parse(jobData);
      job.progress = progress;
      job.message = message;
      job.lastUpdate = new Date().toISOString();
      
      userProperties.setProperty('job_' + jobId, JSON.stringify(job));
      globalJobStatus = job;
    }
  } catch (error) {
    console.error('Error updating job progress:', error);
  }
}

/**
 * Gets job status
 */
function getJobStatus(jobId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const jobData = userProperties.getProperty('job_' + jobId);
    
    if (!jobData) {
      return null;
    }
    
    return JSON.parse(jobData);
  } catch (error) {
    console.error('Error getting job status:', error);
    return null;
  }
}

/**
 * Generates a hash of column headers for schema comparison
 */
function generateSchemaHash(headers) {
  const schemaString = headers.join('|');
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, schemaString)
    .map(byte => (byte & 0xFF).toString(16).padStart(2, '0'))
    .join('');
}

function normalizeDatasetTarget(target, datasetName) {
  const normalized = Object.assign({}, target || {});
  normalized.sheetId = normalized.sheetId ? normalized.sheetId.toString() : '';
  normalized.sheetName = normalized.sheetName || datasetName || 'QBO Data';
  normalized.anchorA1 = sanitizeAnchorCell(normalized.anchorA1);
  normalized.allowResize = normalized.allowResize !== false;
  normalized.namedRange = normalized.namedRange || '';
  return normalized;
}

function sanitizeAnchorCell(anchor) {
  if (!anchor) {
    return 'A1';
  }
  const cleaned = anchor.toString().trim().toUpperCase();
  const anchorRegex = /^[A-Z]+[1-9][0-9]*$/;
  return anchorRegex.test(cleaned) ? cleaned : 'A1';
}

function ensureUniqueSheetName(ss, desiredName) {
  const baseName = desiredName && desiredName.trim() ? desiredName.trim() : 'QBO Data';
  let name = baseName;
  let counter = 1;
  while (ss.getSheetByName(name)) {
    name = `${baseName} (${counter++})`;
  }
  return name;
}

/**
 * Updates or creates a named range
 */
function updateNamedRange(sheet, rangeName, range) {
  try {
    const ss = sheet.getParent();
    const namedRanges = ss.getNamedRanges();
    
    // Remove existing named range with same name
    namedRanges.forEach(namedRange => {
      if (namedRange.getName() === rangeName) {
        namedRange.remove();
      }
    });
    
    // Create new named range
    ss.setNamedRange(rangeName, range);
  } catch (error) {
    console.error('Error updating named range:', error);
  }
}
