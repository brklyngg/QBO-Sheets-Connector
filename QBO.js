/**
 * Minimal QuickBooks API helper utilities for connection testing.
 */

const QBO_PRODUCTION_BASE = 'https://quickbooks.api.intuit.com';
const QBO_SANDBOX_BASE = 'https://sandbox-quickbooks.api.intuit.com';
const QBO_API_VERSION = 'v3';
const QBO_DEFAULT_MINOR_VERSION = '75';

/**
 * Returns the REST API host based on the stored environment.
 */
function getQboBaseUrl() {
  const { environment } = getStoredOAuthCredentials();
  return environment === 'production' ? QBO_PRODUCTION_BASE : QBO_SANDBOX_BASE;
}

/**
 * Ensures we have a connected realm before making requests.
 */
function requireConnectedRealm() {
  const status = getConnectionStatus();
  if (!status.isConnected) {
    throw new Error('Connect to QuickBooks before running API checks.');
  }
  if (!status.realmId) {
    throw new Error('No QuickBooks company is associated with this connection.');
  }
  return status.realmId;
}

/**
 * Sends an authenticated request to the QuickBooks Online API.
 */
function makeAuthenticatedJsonRequest(url, options) {
  const service = getOAuthService();
  if (!service.hasAccess()) {
    throw new Error('OAuth access token unavailable. Reconnect to QuickBooks.');
  }

  const accessToken = service.getAccessToken();
  const headers = Object.assign({
    'Authorization': 'Bearer ' + accessToken,
    'Accept': 'application/json'
  }, options && options.headers ? options.headers : {});

  const fetchOptions = Object.assign({}, options, {
    muteHttpExceptions: true,
    headers: headers
  });

  return UrlFetchApp.fetch(url, fetchOptions);
}

/**
 * Retrieves basic company details to confirm the connection is working.
 */
function fetchCompanyInfo() {
  const realmId = requireConnectedRealm();
  const url = `${getQboBaseUrl()}/${QBO_API_VERSION}/company/${realmId}/companyinfo/${realmId}?minorversion=${QBO_DEFAULT_MINOR_VERSION}`;

  const response = makeAuthenticatedJsonRequest(url, { method: 'GET' });
  const statusCode = response.getResponseCode();
  const body = response.getContentText();

  if (statusCode >= 200 && statusCode < 300) {
    const payload = JSON.parse(body);
    const companyInfo = payload && payload.CompanyInfo ? payload.CompanyInfo : {};
    setCompanyMetadata(realmId, companyInfo.CompanyName || '');
    return companyInfo;
  }

  let errorMessage = `QuickBooks API request failed (${statusCode}).`;
  try {
    const parsed = JSON.parse(body);
    if (parsed && parsed.Fault && parsed.Fault.Error && parsed.Fault.Error.length) {
      errorMessage = parsed.Fault.Error[0].Message || errorMessage;
    }
  } catch (ignored) {
    // Leave default error message.
  }

  throw new Error(errorMessage);
}

/**
 * Simple helper for a connection check button in the UI.
 */
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
