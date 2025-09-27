/**
 * QuickBooks Online API Integration
 * Handles all QBO API calls and data processing
 */

// QBO API constants
const QBO_BASE_URL = 'https://quickbooks.api.intuit.com';
const QBO_SANDBOX_URL = 'https://sandbox-quickbooks.api.intuit.com';
const QBO_API_VERSION = 'v3';

// Supported reports
const SUPPORTED_REPORTS = {
  'profitAndLoss': 'ProfitAndLoss',
  'balanceSheet': 'BalanceSheet',
  'trialBalance': 'TrialBalance',
  'transactionListByDate': 'TransactionList',
  'salesByCustomer': 'CustomerSales'
};

// Entity types that support queries
const QUERYABLE_ENTITIES = [
  'Account', 'Bill', 'BillPayment', 'Budget', 'Class', 'CreditMemo',
  'Customer', 'Department', 'Deposit', 'Employee', 'Estimate', 'Invoice',
  'Item', 'JournalEntry', 'Payment', 'PaymentMethod', 'Purchase',
  'PurchaseOrder', 'RefundReceipt', 'SalesReceipt', 'TaxCode', 'TaxRate',
  'Term', 'TimeActivity', 'Transfer', 'Vendor', 'VendorCredit'
];

/**
 * Gets the base URL for QBO API calls
 */
function getQBOBaseUrl() {
  const isSandbox = PropertiesService.getUserProperties().getProperty('QBO_USE_SANDBOX') === 'true';
  return isSandbox ? QBO_SANDBOX_URL : QBO_BASE_URL;
}

/**
 * Fetches company info from QuickBooks
 */
function fetchCompanyInfo() {
  try {
    const realmId = PropertiesService.getUserProperties().getProperty('QBO_REALM_ID');
    if (!realmId) {
      throw new Error('No QuickBooks company connected');
    }
    
    const url = `${getQBOBaseUrl()}/${QBO_API_VERSION}/company/${realmId}/companyinfo/${realmId}`;
    const minorVersion = getConfig('minorVersion', DEFAULT_MINOR_VERSION);
    
    const startTime = new Date().getTime();
    const response = makeAuthenticatedRequest(url, {
      method: 'GET',
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      }
    });
    const elapsedMs = new Date().getTime() - startTime;
    
    const responseCode = response.getResponseCode();
    const intuitTid = response.getHeaders()['intuit_tid'] || '';
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      const companyInfo = data.CompanyInfo;
      
      // Store company name
      PropertiesService.getUserProperties().setProperty('QBO_COMPANY_NAME', companyInfo.CompanyName);
      
      logAction('fetch_company_info', {
        status: 'success',
        realmId: realmId,
        companyName: companyInfo.CompanyName,
        http_status: responseCode,
        intuit_tid: intuitTid,
        elapsed_ms: elapsedMs
      });
      
      return companyInfo;
    } else {
      const errorText = response.getContentText();
      logAction('fetch_company_info', {
        status: 'error',
        realmId: realmId,
        http_status: responseCode,
        error_message: errorText,
        intuit_tid: intuitTid,
        elapsed_ms: elapsedMs
      });
      
      throw new Error(`Failed to fetch company info: ${responseCode} - ${errorText}`);
    }
  } catch (error) {
    console.error('Error fetching company info:', error);
    throw error;
  }
}

/**
 * Runs a standard report
 */
function runStandardReport(reportType, params = {}) {
  try {
    const realmId = PropertiesService.getUserProperties().getProperty('QBO_REALM_ID');
    if (!realmId) {
      throw new Error('No QuickBooks company connected');
    }
    
    const reportName = SUPPORTED_REPORTS[reportType];
    if (!reportName) {
      throw new Error(`Unsupported report type: ${reportType}`);
    }
    
    // Build query parameters
    const queryParams = [];
    const minorVersion = getConfig('minorVersion', DEFAULT_MINOR_VERSION);
    queryParams.push(`minorversion=${minorVersion}`);
    
    // Add date parameters
    if (params.start_date) queryParams.push(`start_date=${params.start_date}`);
    if (params.end_date) queryParams.push(`end_date=${params.end_date}`);
    if (params.date_macro) queryParams.push(`date_macro=${params.date_macro}`);
    
    // Add other parameters
    if (params.accounting_method) queryParams.push(`accounting_method=${params.accounting_method}`);
    if (params.summarize_column_by) queryParams.push(`summarize_column_by=${params.summarize_column_by}`);
    if (params.columns) queryParams.push(`columns=${params.columns}`);
    
    const url = `${getQBOBaseUrl()}/${QBO_API_VERSION}/company/${realmId}/reports/${reportName}?${queryParams.join('&')}`;
    
    const startTime = new Date().getTime();
    const response = makeAuthenticatedRequest(url, {
      method: 'GET',
      headers: {
        'Accept': 'application/json'
      }
    });
    const elapsedMs = new Date().getTime() - startTime;
    
    const responseCode = response.getResponseCode();
    const intuitTid = response.getHeaders()['intuit_tid'] || '';
    const responseBytes = response.getBlob().getBytes().length;
    
    const logEntry = {
      action: 'run_standard_report',
      dataset_type: 'standard',
      report_name: reportType,
      params_json: JSON.stringify(params),
      start_date: params.start_date || null,
      end_date: params.end_date || null,
      realmId: realmId,
      minorversion: minorVersion,
      elapsed_ms: elapsedMs,
      http_status: responseCode,
      intuit_tid: intuitTid,
      qbo_endpoint: url,
      response_bytes: responseBytes,
      method: 'GET'
    };
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      
      logAction('run_standard_report', {
        ...logEntry,
        status: 'success',
        rows: data.Rows ? data.Rows.length : 0
      });
      
      return {
        success: true,
        data: data,
        intuitTid: intuitTid
      };
    } else if (responseCode === 429) {
      // Rate limit error
      const retryAfter = response.getHeaders()['Retry-After'] || '60';
      const errorData = parseQBOError(response);
      
      logAction('run_standard_report', {
        ...logEntry,
        status: 'error',
        error_code: '429',
        error_message: 'Rate limit exceeded',
        retry_after: retryAfter
      });
      
      return {
        success: false,
        error: 'Rate limit exceeded',
        errorCode: 429,
        retryAfter: parseInt(retryAfter),
        intuitTid: intuitTid
      };
    } else {
      const errorData = parseQBOError(response);
      
      logAction('run_standard_report', {
        ...logEntry,
        status: 'error',
        error_code: errorData.code,
        error_message: errorData.message
      });
      
      return {
        success: false,
        error: errorData.message,
        errorCode: responseCode,
        errorDetail: errorData.detail,
        intuitTid: intuitTid
      };
    }
  } catch (error) {
    console.error('Error running standard report:', error);
    
    logAction('run_standard_report', {
      action: 'run_standard_report',
      dataset_type: 'standard',
      report_name: reportType,
      status: 'error',
      error_message: error.toString()
    });
    
    throw error;
  }
}

/**
 * Runs a custom query
 */
function runCustomQuery(query, startPosition = 1, maxResults = 1000) {
  try {
    const realmId = PropertiesService.getUserProperties().getProperty('QBO_REALM_ID');
    if (!realmId) {
      throw new Error('No QuickBooks company connected');
    }
    
    // Parse and validate query
    const parsedQuery = parseQuery(query);
    if (!parsedQuery.valid) {
      throw new Error(`Invalid query: ${parsedQuery.error}`);
    }
    
    // Build final query with pagination
    let finalQuery = query.trim();
    if (!finalQuery.toLowerCase().includes('startposition')) {
      finalQuery += ` STARTPOSITION ${startPosition}`;
    }
    if (!finalQuery.toLowerCase().includes('maxresults')) {
      finalQuery += ` MAXRESULTS ${maxResults}`;
    }
    
    const minorVersion = getConfig('minorVersion', DEFAULT_MINOR_VERSION);
    
    // Determine transport method
    const usePost = finalQuery.length > 2000;
    const endpoint = `${getQBOBaseUrl()}/${QBO_API_VERSION}/company/${realmId}/query`;
    
    const startTime = new Date().getTime();
    let response;
    
    if (usePost) {
      // Use POST for long queries
      response = makeAuthenticatedRequest(endpoint + `?minorversion=${minorVersion}`, {
        method: 'POST',
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/text'
        },
        payload: finalQuery
      });
    } else {
      // Use GET for short queries
      const encodedQuery = encodeURIComponent(finalQuery);
      response = makeAuthenticatedRequest(`${endpoint}?query=${encodedQuery}&minorversion=${minorVersion}`, {
        method: 'GET',
        headers: {
          'Accept': 'application/json'
        }
      });
    }
    
    const elapsedMs = new Date().getTime() - startTime;
    const responseCode = response.getResponseCode();
    const intuitTid = response.getHeaders()['intuit_tid'] || '';
    const responseBytes = response.getBlob().getBytes().length;
    
    const logEntry = {
      action: 'run_custom_query',
      dataset_type: 'query',
      method: usePost ? 'POST' : 'GET',
      query_select: parsedQuery.select,
      query_from: parsedQuery.from,
      query_where: parsedQuery.where,
      query_orderby: parsedQuery.orderBy,
      query_startposition: startPosition.toString(),
      query_maxresults: maxResults.toString(),
      realmId: realmId,
      minorversion: minorVersion,
      elapsed_ms: elapsedMs,
      http_status: responseCode,
      intuit_tid: intuitTid,
      qbo_endpoint: endpoint,
      response_bytes: responseBytes,
      transport: usePost ? 'POST' : 'GET'
    };
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      const queryResponse = data.QueryResponse || {};
      const entities = queryResponse[parsedQuery.from] || [];
      
      logAction('run_custom_query', {
        ...logEntry,
        status: 'success',
        rows: entities.length
      });
      
      return {
        success: true,
        data: entities,
        totalCount: queryResponse.totalCount || entities.length,
        startPosition: queryResponse.startPosition || startPosition,
        maxResults: queryResponse.maxResults || maxResults,
        intuitTid: intuitTid
      };
    } else if (responseCode === 429) {
      // Rate limit error
      const retryAfter = response.getHeaders()['Retry-After'] || '60';
      
      logAction('run_custom_query', {
        ...logEntry,
        status: 'error',
        error_code: '429',
        error_message: 'Rate limit exceeded',
        retry_after: retryAfter
      });
      
      return {
        success: false,
        error: 'Rate limit exceeded',
        errorCode: 429,
        retryAfter: parseInt(retryAfter),
        intuitTid: intuitTid
      };
    } else {
      const errorData = parseQBOError(response);
      
      logAction('run_custom_query', {
        ...logEntry,
        status: 'error',
        error_code: errorData.code,
        error_message: errorData.message
      });
      
      return {
        success: false,
        error: errorData.message,
        errorCode: responseCode,
        errorDetail: errorData.detail,
        intuitTid: intuitTid
      };
    }
  } catch (error) {
    console.error('Error running custom query:', error);
    
    logAction('run_custom_query', {
      action: 'run_custom_query',
      dataset_type: 'query',
      status: 'error',
      error_message: error.toString()
    });
    
    throw error;
  }
}

/**
 * Parses a QBO query string
 */
function parseQuery(query) {
  try {
    const normalizedQuery = query.trim().replace(/\s+/g, ' ');
    
    // Extract SELECT
    const selectMatch = normalizedQuery.match(/^SELECT\s+(.*?)\s+FROM/i);
    if (!selectMatch) {
      return { valid: false, error: 'Missing SELECT clause' };
    }
    const select = selectMatch[1].trim();
    
    // Extract FROM
    const fromMatch = normalizedQuery.match(/FROM\s+(\w+)/i);
    if (!fromMatch) {
      return { valid: false, error: 'Missing FROM clause' };
    }
    const from = fromMatch[1];
    
    // Validate entity
    if (!QUERYABLE_ENTITIES.includes(from)) {
      return { valid: false, error: `Invalid entity: ${from}. Supported entities: ${QUERYABLE_ENTITIES.join(', ')}` };
    }
    
    // Extract WHERE (optional)
    const whereMatch = normalizedQuery.match(/WHERE\s+(.*?)(?:\s+ORDER\s+BY|\s+STARTPOSITION|\s+MAXRESULTS|$)/i);
    const where = whereMatch ? whereMatch[1].trim() : null;
    
    // Extract ORDER BY (optional)
    const orderByMatch = normalizedQuery.match(/ORDER\s+BY\s+(.*?)(?:\s+STARTPOSITION|\s+MAXRESULTS|$)/i);
    const orderBy = orderByMatch ? orderByMatch[1].trim() : null;
    
    // Extract STARTPOSITION (optional)
    const startPosMatch = normalizedQuery.match(/STARTPOSITION\s+(\d+)/i);
    const startPosition = startPosMatch ? parseInt(startPosMatch[1]) : null;
    
    // Extract MAXRESULTS (optional)
    const maxResultsMatch = normalizedQuery.match(/MAXRESULTS\s+(\d+)/i);
    const maxResults = maxResultsMatch ? parseInt(maxResultsMatch[1]) : null;
    
    return {
      valid: true,
      select: select,
      from: from,
      where: where,
      orderBy: orderBy,
      startPosition: startPosition,
      maxResults: maxResults
    };
  } catch (error) {
    return { valid: false, error: error.toString() };
  }
}

/**
 * Parses QBO error response
 */
function parseQBOError(response) {
  try {
    const contentText = response.getContentText();
    const contentType = response.getHeaders()['Content-Type'] || '';
    
    if (contentType.includes('application/json')) {
      const errorData = JSON.parse(contentText);
      
      if (errorData.Fault) {
        const fault = errorData.Fault;
        const error = fault.Error && fault.Error[0] ? fault.Error[0] : {};
        
        return {
          code: error.code || 'UNKNOWN',
          message: error.Message || 'Unknown error',
          detail: error.Detail || null,
          type: fault.type || null
        };
      }
    }
    
    // Fallback for non-JSON responses
    return {
      code: response.getResponseCode().toString(),
      message: contentText || 'Unknown error',
      detail: null,
      type: null
    };
  } catch (error) {
    return {
      code: 'PARSE_ERROR',
      message: 'Failed to parse error response',
      detail: error.toString(),
      type: null
    };
  }
}

/**
 * Converts report data to sheet format
 */
function convertReportToSheetData(reportData) {
  try {
    const rows = [];
    const report = reportData;
    
    // Add report header
    if (report.Header) {
      rows.push([report.Header.ReportName || 'QuickBooks Report']);
      
      // Add date range if available
      if (report.Header.StartPeriod && report.Header.EndPeriod) {
        rows.push([`Period: ${report.Header.StartPeriod} to ${report.Header.EndPeriod}`]);
      } else if (report.Header.DateMacro) {
        rows.push([`Period: ${report.Header.DateMacro}`]);
      }
      
      rows.push([]); // Empty row
    }
    
    // Process columns
    if (report.Columns && report.Columns.Column) {
      const columns = report.Columns.Column;
      const headerRow = columns.map(col => col.ColTitle || '');
      rows.push(headerRow);
    }
    
    // Process rows
    if (report.Rows && report.Rows.Row) {
      processReportRows(report.Rows.Row, rows);
    }
    
    return {
      data: rows,
      rows: rows.length,
      cols: rows.length > 0 ? rows[0].length : 0
    };
  } catch (error) {
    console.error('Error converting report to sheet data:', error);
    throw error;
  }
}

/**
 * Recursively processes report rows
 */
function processReportRows(reportRows, outputRows, level = 0) {
  if (!Array.isArray(reportRows)) {
    reportRows = [reportRows];
  }
  
  reportRows.forEach(row => {
    if (row.ColData) {
      const rowData = [];
      
      // Add indentation for nested rows
      if (level > 0 && row.ColData[0]) {
        row.ColData[0].value = '  '.repeat(level) + (row.ColData[0].value || '');
      }
      
      // Extract column values
      row.ColData.forEach(col => {
        rowData.push(col.value || '');
      });
      
      outputRows.push(rowData);
    }
    
    // Process sub-rows
    if (row.Rows && row.Rows.Row) {
      processReportRows(row.Rows.Row, outputRows, level + 1);
    }
    
    // Process summary row
    if (row.Summary && row.Summary.ColData) {
      const summaryData = [];
      row.Summary.ColData.forEach(col => {
        summaryData.push(col.value || '');
      });
      outputRows.push(summaryData);
    }
  });
}

/**
 * Converts entity data to sheet format
 */
function convertEntitiesToSheetData(entities, entityType) {
  try {
    if (!entities || entities.length === 0) {
      return {
        data: [['No data found']],
        rows: 1,
        cols: 1
      };
    }
    
    // Get all unique keys from all entities
    const allKeys = new Set();
    entities.forEach(entity => {
      Object.keys(entity).forEach(key => allKeys.add(key));
    });
    
    // Sort keys for consistent column order
    const keys = Array.from(allKeys).sort();
    
    // Create header row
    const rows = [keys];
    
    // Add data rows
    entities.forEach(entity => {
      const row = keys.map(key => {
        const value = entity[key];
        
        // Handle nested objects
        if (value && typeof value === 'object') {
          // Special handling for common nested fields
          if (key === 'MetaData') {
            return value.LastUpdatedTime || '';
          } else if (key === 'CurrencyRef') {
            return value.value || '';
          } else if (key === 'CustomerRef' || key === 'VendorRef' || key === 'AccountRef') {
            return value.name || value.value || '';
          } else {
            return JSON.stringify(value);
          }
        }
        
        return value !== null && value !== undefined ? value.toString() : '';
      });
      
      rows.push(row);
    });
    
    return {
      data: rows,
      rows: rows.length,
      cols: keys.length
    };
  } catch (error) {
    console.error('Error converting entities to sheet data:', error);
    throw error;
  }
}

/**
 * Gets available report types for UI
 */
function getAvailableReports() {
  return Object.keys(SUPPORTED_REPORTS).map(key => ({
    id: key,
    name: SUPPORTED_REPORTS[key],
    displayName: formatReportName(key)
  }));
}

/**
 * Gets available entity types for queries
 */
function getAvailableEntities() {
  return QUERYABLE_ENTITIES.map(entity => ({
    id: entity,
    name: entity,
    displayName: entity
  }));
}

/**
 * Formats report name for display
 */
function formatReportName(reportType) {
  const names = {
    'profitAndLoss': 'Profit and Loss',
    'balanceSheet': 'Balance Sheet',
    'trialBalance': 'Trial Balance',
    'transactionListByDate': 'Transaction List by Date',
    'salesByCustomer': 'Sales by Customer'
  };
  
  return names[reportType] || reportType;
}

/**
 * Validates date format (YYYY-MM-DD)
 */
function isValidDate(dateString) {
  const regex = /^\d{4}-\d{2}-\d{2}$/;
  if (!regex.test(dateString)) return false;
  
  const date = new Date(dateString);
  return date instanceof Date && !isNaN(date);
}
