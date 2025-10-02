const QBO_ENV = {
  sandbox: 'https://sandbox-quickbooks.api.intuit.com',
  production: 'https://quickbooks.api.intuit.com'
};

function qboBaseUrl_() {
  const cfg = getConfig_();
  if (!cfg.realmId) {
    throw new Error('Missing QBO realmId. Authorize first.');
  }
  const host = QBO_ENV[cfg.env] || QBO_ENV.sandbox;
  return host + '/v3/company/' + encodeURIComponent(cfg.realmId);
}

function qboUrl_(path, params) {
  const minor = (getConfig_().minor || '75').trim();
  const qp = Object.assign({}, params || {}, { minorversion: minor });
  const qs = Object.keys(qp)
    .filter(key => qp[key] != null && qp[key] !== '')
    .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(String(qp[key])))
    .join('&');
  const normalizedPath = path.startsWith('/') ? path : '/' + path;
  return qboBaseUrl_() + normalizedPath + (qs ? '?' + qs : '');
}

function qboFetch_(url, options, attempt) {
  const svc = getQboService_();
  if (!svc.hasAccess()) {
    throw new Error('Not authorized.');
  }

  const token = svc.getAccessToken();
  const method = options && options.method ? options.method : 'GET';
  const headers = Object.assign({
    Authorization: 'Bearer ' + token,
    Accept: 'application/json'
  }, options && options.headers ? options.headers : {});

  const fetchParams = {
    method: method,
    headers: headers,
    muteHttpExceptions: true
  };

  if (options && options.payload != null) {
    fetchParams.payload = options.payload;
  }

  if (options && options.contentType) {
    fetchParams.contentType = options.contentType;
  }

  const attemptNumber = attempt || 1;
  const maxAttempts = 6;
  const response = UrlFetchApp.fetch(url, fetchParams);
  const status = response.getResponseCode();
  const bodyText = response.getContentText() || '';
  const responseHeaders = response.getHeaders();

  if (status >= 200 && status < 300) {
    return bodyText ? JSON.parse(bodyText) : {};
  }

  if (status === 401 && attemptNumber === 1) {
    try {
      svc.refresh();
    } catch (err) {
      console.error('Token refresh failed', err);
    }
    return qboFetch_(url, options, attemptNumber + 1);
  }

  if ((status === 429 || (status >= 500 && status <= 504)) && attemptNumber < maxAttempts) {
    const retryAfter = Number(responseHeaders['Retry-After'] || 0);
    const baseDelay = Math.min(30000, Math.pow(2, attemptNumber) * 250);
    const jitter = Math.floor(Math.random() * 250);
    const waitMs = retryAfter > 0 ? retryAfter * 1000 : baseDelay + jitter;
    Utilities.sleep(waitMs);
    return qboFetch_(url, options, attemptNumber + 1);
  }

  let message = 'QBO HTTP ' + status;
  try {
    const parsed = JSON.parse(bodyText);
    if (parsed && parsed.Fault && parsed.Fault.Error && parsed.Fault.Error.length) {
      const err = parsed.Fault.Error[0];
      if (err.Message) {
        message += ': ' + err.Message;
      }
      if (err.Detail) {
        message += ' | ' + err.Detail;
      }
    } else if (parsed && parsed.error) {
      message += ': ' + parsed.error;
    }
  } catch (parseErr) {
    message += ': ' + bodyText;
  }

  if (status === 401) {
    try {
      svc.reset();
    } catch (resetErr) {
      console.error('Failed to reset OAuth service after 401', resetErr);
    }
    throw new Error(message + ' | Authorization expired. Re-authorize.');
  }

  throw new Error(message);
}

function qboReport_(reportName, params) {
  if (!reportName) {
    throw new Error('Missing reportName.');
  }
  const url = qboUrl_('/reports/' + encodeURIComponent(reportName), params || {});
  return qboFetch_(url, { method: 'GET' });
}

function qboReportToTable_(reportJson) {
  const report = reportJson && reportJson.Report;
  if (!report) {
    return [['No data']];
  }

  const headers = (report.Columns && report.Columns.Column ? report.Columns.Column : [])
    .map(col => col.ColTitle || col.ColType || '');

  const rows = [];

  function pushRow(values) {
    const arr = (values || []).map(value => value || '');
    if (!arr.some(value => value !== '')) {
      return;
    }
    rows.push(arr);
  }

  function walk(node) {
    if (!node) {
      return;
    }
    (node.Row || []).forEach(row => {
      if (row.type === 'Section') {
        const header = (row.Header && row.Header.ColData ? row.Header.ColData : [])
          .map(cell => cell.value || '');
        pushRow(header);
        walk(row.Rows);
        if (row.Summary && row.Summary.ColData) {
          const summary = row.Summary.ColData.map(cell => cell.value || '');
          pushRow(summary);
        }
      } else if (row.type === 'Data') {
        const dataRow = (row.ColData || []).map(cell => cell.value || '');
        pushRow(dataRow);
      }
    });
  }

  walk(report.Rows);

  if (!headers.length && rows.length === 0) {
    return [['No data']];
  }

  let width = headers.length;
  rows.forEach(row => {
    if (row.length > width) {
      width = row.length;
    }
  });
  width = Math.max(1, width);

  const headerRow = headers.slice();
  while (headerRow.length < width) {
    headerRow.push('');
  }

  if (rows.length === 0) {
    return headerRow.some(value => value !== '') ? [headerRow] : [['No data']];
  }

  const dataRows = rows.map(row => {
    const copy = row.slice();
    while (copy.length < width) {
      copy.push('');
    }
    return copy;
  });

  return [headerRow].concat(dataRows);
}

function normalizeTable_(table) {
  if (!Array.isArray(table) || table.length === 0) {
    return { rows: 0, cols: 0, values: [] };
  }
  let cols = 0;
  table.forEach(row => {
    if (Array.isArray(row) && row.length > cols) {
      cols = row.length;
    }
  });
  const width = Math.max(1, cols);
  const values = table.map(row => {
    const safe = Array.isArray(row) ? row.slice(0, width) : [];
    while (safe.length < width) {
      safe.push('');
    }
    return safe;
  });
  return { rows: values.length, cols: width, values: values };
}

function writeTableToSheet_(sheetName, table) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sheet.clearContents();
  const normalized = normalizeTable_(table);
  if (normalized.rows === 0 || normalized.cols === 0) {
    return;
  }
  sheet.getRange(1, 1, normalized.rows, normalized.cols).setValues(normalized.values);
  for (let col = 1; col <= normalized.cols; col++) {
    sheet.autoResizeColumn(col);
  }
  try {
    sheet.setFrozenRows(1);
  } catch (err) {
    console.error('Failed to freeze header row', err);
  }
}

function logReport_(reportName, params, status, dims, elapsed) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('QBO_Log') || ss.insertSheet('QBO_Log');
    sheet.appendRow([
      new Date(),
      'REPORT',
      reportName,
      JSON.stringify(params || {}),
      status,
      dims || '',
      elapsed || ''
    ]);
  } catch (err) {
    console.error('Failed to log report run', err);
  }
}

const QBO_API = {
  fetchReport: function(reportName, params) {
    const started = Date.now();
    try {
      const json = qboReport_(reportName, params);
      const table = qboReportToTable_(json);
      const normalized = normalizeTable_(table);
      const dims = normalized.rows + 'x' + normalized.cols;
      logReport_(reportName, params, 'OK(fetch)', dims, Date.now() - started);
      return normalized.values;
    } catch (error) {
      logReport_(reportName, params, 'ERR(fetch): ' + error.message);
      throw error;
    }
  },
  reportToSheet: function(reportName, params, sheetName) {
    const started = Date.now();
    const values = this.fetchReport(reportName, params);
    const normalized = normalizeTable_(values);
    const name = sheetName || ('QBO_' + reportName);
    writeTableToSheet_(name, values);
    const dims = normalized.rows + 'x' + normalized.cols;
    logReport_(reportName, params, 'OK(write)', dims, Date.now() - started);
    return {
      rows: normalized.rows,
      cols: normalized.cols,
      sheet: name
    };
  }
};
