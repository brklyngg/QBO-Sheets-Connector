/**
 * QuickBooks Online API integration shared by the auth-first build and the legacy feature set.
 * Provides report/query execution helpers that work with the modern OAuth helpers in Auth.js.
 */

const QBO_PRODUCTION_BASE = 'https://quickbooks.api.intuit.com';
const QBO_SANDBOX_BASE = 'https://sandbox-quickbooks.api.intuit.com';
const QBO_API_VERSION = 'v3';
const QBO_DEFAULT_MINOR_VERSION = '75';

const QBO_SUPPORTED_REPORTS = {
  profitAndLoss: 'ProfitAndLoss',
  balanceSheet: 'BalanceSheet',
  trialBalance: 'TrialBalance',
  transactionListByDate: 'TransactionList',
  salesByCustomer: 'CustomerSales'
};

const QBO_QUERYABLE_ENTITIES = [
  'Account', 'Bill', 'BillPayment', 'Budget', 'Class', 'CreditMemo',
  'Customer', 'Department', 'Deposit', 'Employee', 'Estimate', 'Invoice',
  'Item', 'JournalEntry', 'Payment', 'PaymentMethod', 'Purchase',
  'PurchaseOrder', 'RefundReceipt', 'SalesReceipt', 'TaxCode', 'TaxRate',
  'Term', 'TimeActivity', 'Transfer', 'Vendor', 'VendorCredit'
];

const QBO_MAX_PAGE_SIZE = 1000;

function getQboBaseUrl() {
  const { environment } = getStoredOAuthCredentials();
  return environment === 'production' ? QBO_PRODUCTION_BASE : QBO_SANDBOX_BASE;
}

function getQboMinorVersion() {
  return typeof getConfig === 'function' ? getConfig('minorVersion', QBO_DEFAULT_MINOR_VERSION) : QBO_DEFAULT_MINOR_VERSION;
}

function requireConnectedRealm() {
  const status = getConnectionStatus();
  if (!status.isConnected) {
    throw new Error('Connect to QuickBooks before running API requests.');
  }
  if (!status.realmId) {
    throw new Error('No QuickBooks company is associated with this connection.');
  }
  return status.realmId;
}

function makeAuthenticatedJsonRequest(url, options) {
  const service = getOAuthService();
  if (!service.hasAccess()) {
    throw new Error('OAuth access token unavailable. Reconnect to QuickBooks.');
  }

  const accessToken = service.getAccessToken();
  const headers = Object.assign({
    Authorization: 'Bearer ' + accessToken,
    Accept: 'application/json'
  }, options && options.headers ? options.headers : {});

  const fetchOptions = Object.assign({}, options, {
    muteHttpExceptions: true,
    headers: headers
  });

  return UrlFetchApp.fetch(url, fetchOptions);
}

function getHeaderCaseInsensitive(headers, key) {
  if (!headers) {
    return null;
  }
  const target = key.toLowerCase();
  for (const header in headers) {
    if (!Object.prototype.hasOwnProperty.call(headers, header)) {
      continue;
    }
    if (header && header.toLowerCase() === target) {
      return headers[header];
    }
  }
  return null;
}

function fetchCompanyInfo() {
  try {
    const realmId = requireConnectedRealm();
    const minorVersion = getQboMinorVersion();
    const url = `${getQboBaseUrl()}/${QBO_API_VERSION}/company/${realmId}/companyinfo/${realmId}?minorversion=${minorVersion}`;

    const start = Date.now();
    const response = makeAuthenticatedJsonRequest(url, { method: 'GET' });
    const elapsedMs = Date.now() - start;
    const statusCode = response.getResponseCode();
    const headers = response.getHeaders();
    const bodyText = response.getContentText();
    const intuitTid = getHeaderCaseInsensitive(headers, 'intuit_tid') || '';
    const retryAfter = getHeaderCaseInsensitive(headers, 'Retry-After') || null;

    if (statusCode >= 200 && statusCode < 300) {
      const payload = JSON.parse(bodyText);
      const companyInfo = payload && payload.CompanyInfo ? payload.CompanyInfo : {};
      setCompanyMetadata(realmId, companyInfo.CompanyName || '');
      logAction('fetch_company_info', {
        status: 'success',
        realmId: realmId,
        companyName: companyInfo.CompanyName || '',
        http_status: statusCode,
        intuit_tid: intuitTid,
        elapsed_ms: elapsedMs,
        retry_after: retryAfter
      });
      return companyInfo;
    }

    const fault = parseQboError(response);
    logAction('fetch_company_info', {
      status: 'error',
      realmId: realmId,
      http_status: statusCode,
      error_message: fault.message,
      error_code: fault.code,
      intuit_tid: intuitTid,
      elapsed_ms: elapsedMs,
      retry_after: retryAfter
    });

    throw new Error(fault.message || `Failed to fetch company info (${statusCode}).`);
  } catch (error) {
    console.error('Error fetching company info:', error);
    throw error;
  }
}

function runStandardReport(reportType, params = {}) {
  try {
    const realmId = requireConnectedRealm();
    const reportName = QBO_SUPPORTED_REPORTS[reportType];
    if (!reportName) {
      throw new Error(`Unsupported report type: ${reportType}`);
    }

    const queryParams = [];
    const minorVersion = getQboMinorVersion();
    queryParams.push(`minorversion=${minorVersion}`);

    if (params.start_date) queryParams.push(`start_date=${params.start_date}`);
    if (params.end_date) queryParams.push(`end_date=${params.end_date}`);
    if (params.date_macro) queryParams.push(`date_macro=${params.date_macro}`);
    if (params.accounting_method) queryParams.push(`accounting_method=${params.accounting_method}`);
    if (params.summarize_column_by) queryParams.push(`summarize_column_by=${params.summarize_column_by}`);
    if (params.columns) queryParams.push(`columns=${params.columns}`);

    const url = `${getQboBaseUrl()}/${QBO_API_VERSION}/company/${realmId}/reports/${reportName}?${queryParams.join('&')}`;

    const start = Date.now();
    const response = makeAuthenticatedJsonRequest(url, { method: 'GET' });
    const elapsedMs = Date.now() - start;
    const headers = response.getHeaders();
    const statusCode = response.getResponseCode();
    const intuitTid = getHeaderCaseInsensitive(headers, 'intuit_tid') || '';
    const retryAfter = getHeaderCaseInsensitive(headers, 'Retry-After') || null;
    const responseBytes = response.getContentText().length;

    const logPayload = {
      action: 'run_standard_report',
      dataset_type: 'standard',
      report_name: reportType,
      params_json: JSON.stringify(params),
      start_date: params.start_date || null,
      end_date: params.end_date || null,
      realmId: realmId,
      minorversion: minorVersion,
      elapsed_ms: elapsedMs,
      http_status: statusCode,
      intuit_tid: intuitTid,
      qbo_endpoint: url,
      response_bytes: responseBytes,
      method: 'GET',
      retry_after: retryAfter
    };

    if (statusCode === 200) {
      const data = JSON.parse(response.getContentText());
      logAction('run_standard_report', Object.assign(logPayload, {
        status: 'success',
        rows: data.Rows ? data.Rows.length : 0
      }));

      return {
        success: true,
        data: data,
        intuitTid: intuitTid
      };
    }

    if (statusCode === 429) {
      logAction('run_standard_report', Object.assign(logPayload, {
        status: 'error',
        error_code: '429',
        error_message: 'Rate limit exceeded'
      }));

      return {
        success: false,
        error: 'Rate limit exceeded',
        errorCode: 429,
        retryAfter: parseInt(retryAfter || '60', 10),
        intuitTid: intuitTid
      };
    }

    const fault = parseQboError(response);
    logAction('run_standard_report', Object.assign(logPayload, {
      status: 'error',
      error_code: fault.code,
      error_message: fault.message
    }));

    return {
      success: false,
      error: fault.message,
      errorCode: statusCode,
      errorDetail: fault.detail,
      intuitTid: intuitTid
    };
  } catch (error) {
    console.error('Error running standard report:', error);
    logAction('run_standard_report', {
      status: 'error',
      error_message: error.toString(),
      report_name: reportType,
      dataset_type: 'standard'
    });
    throw error;
  }
}

function runCustomQuery(query, startPosition = 1, maxResults = QBO_MAX_PAGE_SIZE, options = {}) {
  try {
    const realmId = requireConnectedRealm();
    const parsedQuery = parseQuery(query);
    if (!parsedQuery.valid) {
      throw new Error(`Invalid query: ${parsedQuery.error}`);
    }

    const fetchAll = options.fetchAll === true;
    const maxPages = options.maxPages || 50;
    const pageSize = Math.max(1, Math.min(QBO_MAX_PAGE_SIZE, parseInt(maxResults, 10) || QBO_MAX_PAGE_SIZE));
    let currentStart = Math.max(1, parseInt(startPosition, 10) || 1);

    const baseQuery = stripQueryPaginationClauses(query.trim());
    const minorVersion = getQboMinorVersion();
    const endpoint = `${getQboBaseUrl()}/${QBO_API_VERSION}/company/${realmId}/query`;

    const aggregatedEntities = [];
    let totalCount = 0;
    let fetchedPages = 0;
    let hasMore = false;
    let nextStartPosition = currentStart;
    let intuitTid = '';

    while (true) {
      fetchedPages++;
      const paginatedQuery = appendQueryPaginationClauses(baseQuery, currentStart, pageSize);
      const usePost = paginatedQuery.length > 2000;

      const requestStart = Date.now();
      let response;

      if (usePost) {
        response = makeAuthenticatedJsonRequest(`${endpoint}?minorversion=${minorVersion}`, {
          method: 'POST',
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/text'
          },
          payload: paginatedQuery
        });
      } else {
        const encodedQuery = encodeURIComponent(paginatedQuery);
        response = makeAuthenticatedJsonRequest(`${endpoint}?query=${encodedQuery}&minorversion=${minorVersion}`, {
          method: 'GET',
          headers: {
            'Accept': 'application/json'
          }
        });
      }

      const elapsedMs = Date.now() - requestStart;
      const statusCode = response.getResponseCode();
      const headers = response.getHeaders();
      const retryAfter = getHeaderCaseInsensitive(headers, 'Retry-After') || null;
      const responseBytes = response.getContentText().length;

      const logPayload = {
        action: 'run_custom_query',
        dataset_type: 'query',
        method: usePost ? 'POST' : 'GET',
        query_select: parsedQuery.select,
        query_from: parsedQuery.from,
        query_where: parsedQuery.where,
        query_orderby: parsedQuery.orderBy,
        query_startposition: currentStart.toString(),
        query_maxresults: pageSize.toString(),
        realmId: realmId,
        minorversion: minorVersion,
        elapsed_ms: elapsedMs,
        http_status: statusCode,
        intuit_tid: getHeaderCaseInsensitive(headers, 'intuit_tid') || '',
        qbo_endpoint: endpoint,
        response_bytes: responseBytes,
        transport: usePost ? 'POST' : 'GET',
        retry_after: retryAfter,
        page_index: fetchedPages
      };

      if (statusCode === 200) {
        const data = JSON.parse(response.getContentText());
        const queryResponse = data.QueryResponse || {};
        const entities = queryResponse[parsedQuery.from] || [];
        aggregatedEntities.push(...entities);
        totalCount = queryResponse.totalCount != null ? queryResponse.totalCount : (totalCount || aggregatedEntities.length);
        const returnedCount = entities.length;
        const startFromResponse = queryResponse.startPosition || currentStart;
        nextStartPosition = startFromResponse + returnedCount;
        intuitTid = logPayload.intuit_tid || intuitTid;

        const moreByTotal = queryResponse.totalCount != null ? (startFromResponse - 1 + returnedCount) < queryResponse.totalCount : false;
        const moreByBatch = returnedCount === pageSize;
        const progressed = nextStartPosition > currentStart;
        hasMore = (moreByTotal || moreByBatch) && progressed;

        logAction('run_custom_query', Object.assign(logPayload, {
          status: 'success',
          rows: returnedCount,
          cols: entities.length ? Object.keys(entities[0]).length : 0
        }));

        if (!fetchAll || !hasMore || fetchedPages >= maxPages) {
          break;
        }

        currentStart = nextStartPosition;
        Utilities.sleep(200);
        continue;
      }

      if (statusCode === 429) {
        logAction('run_custom_query', Object.assign(logPayload, {
          status: 'error',
          error_code: '429',
          error_message: 'Rate limit exceeded'
        }));

        return {
          success: false,
          error: 'Rate limit exceeded',
          errorCode: 429,
          retryAfter: parseInt(retryAfter || '60', 10),
          intuitTid: logPayload.intuit_tid
        };
      }

      const fault = parseQboError(response);
      logAction('run_custom_query', Object.assign(logPayload, {
        status: 'error',
        error_code: fault.code,
        error_message: fault.message
      }));

      return {
        success: false,
        error: fault.message,
        errorCode: statusCode,
        errorDetail: fault.detail,
        intuitTid: logPayload.intuit_tid
      };
    }

    return {
      success: true,
      data: aggregatedEntities,
      totalCount: totalCount,
      hasMore: hasMore,
      nextStartPosition: hasMore ? nextStartPosition : null,
      intuitTid: intuitTid
    };
  } catch (error) {
    console.error('Error running custom query:', error);
    logAction('run_custom_query', {
      status: 'error',
      error_message: error.toString(),
      dataset_type: 'query'
    });
    throw error;
  }
}

function parseQuery(query) {
  try {
    const normalized = query.trim().replace(/\s+/g, ' ');
    const selectMatch = normalized.match(/^select\s+(.+?)\s+from\s+/i);
    const fromMatch = normalized.match(/from\s+(\w+)/i);
    const whereMatch = normalized.match(/where\s+(.+?)(?:\s+order\s+by|\s+startposition|\s+maxresults|$)/i);
    const orderByMatch = normalized.match(/order\s+by\s+(.+?)(?:\s+startposition|\s+maxresults|$)/i);
    const startPositionMatch = normalized.match(/startposition\s+(\d+)/i);
    const maxResultsMatch = normalized.match(/maxresults\s+(\d+)/i);

    if (!selectMatch || !fromMatch) {
      return {
        valid: false,
        error: 'Query must include SELECT and FROM clauses'
      };
    }

    return {
      valid: true,
      select: selectMatch[1],
      from: fromMatch[1],
      where: whereMatch ? whereMatch[1] : null,
      orderBy: orderByMatch ? orderByMatch[1] : null,
      startPosition: startPositionMatch ? parseInt(startPositionMatch[1], 10) : null,
      maxResults: maxResultsMatch ? parseInt(maxResultsMatch[1], 10) : null
    };
  } catch (error) {
    return {
      valid: false,
      error: error.message
    };
  }
}

function parseQboError(response) {
  let message = 'QuickBooks request failed.';
  let code = null;
  let detail = null;

  try {
    const payload = JSON.parse(response.getContentText());
    if (payload && payload.Fault && payload.Fault.Error && payload.Fault.Error.length) {
      const faultError = payload.Fault.Error[0];
      message = faultError.Message || message;
      code = faultError.Code || null;
      detail = faultError.Detail || null;
    }
  } catch (error) {
    message = `${message} (${response.getResponseCode()})`;
  }

  return { message, code, detail };
}

function convertReportToSheetData(reportData) {
  const headers = [];
  const rows = [];

  if (!reportData) {
    return {
      data: [],
      rows: 0,
      cols: 0
    };
  }

  if (reportData.Columns && reportData.Columns.Column) {
    reportData.Columns.Column.forEach(column => {
      headers.push(column.ColTitle || column.ColType || 'Column');
    });
    rows.push(headers);
  }

  if (reportData.Rows && reportData.Rows.Row) {
    processReportRows(reportData.Rows.Row, rows, 0);
  }

  return {
    data: rows,
    rows: rows.length,
    cols: rows.length ? rows[0].length : 0
  };
}

function processReportRows(reportRows, outputRows, level = 0) {
  if (!Array.isArray(reportRows)) {
    return;
  }

  reportRows.forEach(row => {
    if (row.type === 'Section') {
      const header = row.Header && row.Header.ColData ? row.Header.ColData.map(col => col.value || '') : [];
      if (header.length) {
        outputRows.push(header);
      }
      processReportRows(row.Rows ? row.Rows.Row : [], outputRows, level + 1);
      if (row.Summary && row.Summary.ColData) {
        outputRows.push(row.Summary.ColData.map(col => col.value || ''));
      }
    } else if (row.type === 'Data' && row.ColData) {
      const rowData = row.ColData.map(col => col.value || '');
      if (level > 0) {
        rowData.unshift(new Array(level).fill('').join(''));
      }
      outputRows.push(rowData);
    } else if (row.type === 'Summary' && row.Summary && row.Summary.ColData) {
      outputRows.push(row.Summary.ColData.map(col => col.value || ''));
    }
  });
}

function convertEntitiesToSheetData(entities, entityType) {
  if (!Array.isArray(entities) || entities.length === 0) {
    return {
      data: [],
      rows: 0,
      cols: 0
    };
  }

  const headers = Object.keys(entities[0]);
  const rows = [headers];

  entities.forEach(entity => {
    const row = headers.map(header => {
      const value = entity[header];
      if (value === null || value === undefined) {
        return '';
      }
      if (typeof value === 'object') {
        return JSON.stringify(value);
      }
      return value;
    });
    rows.push(row);
  });

  return {
    data: rows,
    rows: rows.length,
    cols: headers.length,
    entityType: entityType
  };
}

function getAvailableReports() {
  return Object.keys(QBO_SUPPORTED_REPORTS).map(key => ({
    id: key,
    apiName: QBO_SUPPORTED_REPORTS[key],
    label: formatReportName(key)
  }));
}

function stripQueryPaginationClauses(queryText) {
  return queryText
    .replace(/\s+startposition\s+\d+/i, '')
    .replace(/\s+maxresults\s+\d+/i, '')
    .trim();
}

function appendQueryPaginationClauses(baseQuery, startPosition, maxResults) {
  return `${baseQuery} STARTPOSITION ${startPosition} MAXRESULTS ${maxResults}`;
}

function getAvailableEntities() {
  return QBO_QUERYABLE_ENTITIES.slice();
}

function formatReportName(reportType) {
  return reportType
    .replace(/([A-Z])/g, ' $1')
    .replace(/^(.)/, (_, first) => first.toUpperCase())
    .trim();
}

function isValidDate(dateString) {
  if (!dateString) {
    return false;
  }
  const date = new Date(dateString);
  return !isNaN(date.getTime());
}

function testQuickBooksConnection() {
  try {
    const companyInfo = fetchCompanyInfo();
    return {
      success: true,
      companyName: companyInfo.CompanyName || 'Unknown Company',
      legalName: companyInfo.LegalName || '',
      realmId: companyInfo.Id || null,
      country: companyInfo.Country || '',
      fiscalYearStartMonth: companyInfo.FiscalYearStartMonth || ''
    };
  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}
