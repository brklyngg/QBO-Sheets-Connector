/**
 * Minimal QuickBooks Online OAuth helper functions.
 * Focused solely on establishing and clearing the connection.
 */

const OAUTH_SERVICE_NAME = 'QuickBooksConnectionTester';
const OAUTH_CALLBACK_FUNCTION = 'authCallback';
const OAUTH_SCOPE = 'com.intuit.quickbooks.accounting';
const INTUIT_AUTH_URL = 'https://appcenter.intuit.com/connect/oauth2';
const INTUIT_TOKEN_URL = 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer';

const PROPERTY_CLIENT_ID = 'QBO_CLIENT_ID';
const PROPERTY_CLIENT_SECRET = 'QBO_CLIENT_SECRET';
const PROPERTY_ENVIRONMENT = 'QBO_ENVIRONMENT';
const PROPERTY_REALM_ID = 'QBO_REALM_ID';
const PROPERTY_COMPANY_NAME = 'QBO_COMPANY_NAME';
const PROPERTY_LAST_CONNECTED = 'QBO_LAST_CONNECTED_AT';

/**
 * Convenience accessor for user-level properties.
 */
function getUserProps() {
  return PropertiesService.getUserProperties();
}

/**
 * Reads the stored OAuth client credentials.
 */
function getStoredOAuthCredentials() {
  const props = getUserProps();
  return {
    clientId: (props.getProperty(PROPERTY_CLIENT_ID) || '').trim(),
    clientSecret: (props.getProperty(PROPERTY_CLIENT_SECRET) || '').trim(),
    environment: (props.getProperty(PROPERTY_ENVIRONMENT) || 'sandbox').trim() || 'sandbox'
  };
}

/**
 * Persists OAuth credentials entered by the user.
 */
function setOAuthCredentials(clientId, clientSecret, environment) {
  const props = getUserProps();
  props.setProperty(PROPERTY_CLIENT_ID, String(clientId || '').trim());
  props.setProperty(PROPERTY_CLIENT_SECRET, String(clientSecret || '').trim());
  props.setProperty(PROPERTY_ENVIRONMENT, environment === 'production' ? 'production' : 'sandbox');
  resetOAuthState();
}

/**
 * Clears company-specific connection state.
 */
function clearConnectionMetadata() {
  const props = getUserProps();
  props.deleteProperty(PROPERTY_REALM_ID);
  props.deleteProperty(PROPERTY_COMPANY_NAME);
  props.deleteProperty(PROPERTY_LAST_CONNECTED);
}

/**
 * Returns the Apps Script OAuth2 service configured for Intuit.
 */
function getOAuthService() {
  const { clientId, clientSecret } = getStoredOAuthCredentials();

  return OAuth2.createService(OAUTH_SERVICE_NAME)
    .setAuthorizationBaseUrl(INTUIT_AUTH_URL)
    .setTokenUrl(INTUIT_TOKEN_URL)
    .setClientId(clientId || '')
    .setClientSecret(clientSecret || '')
    .setCallbackFunction(OAUTH_CALLBACK_FUNCTION)
    .setPropertyStore(getUserProps())
    .setScope(OAUTH_SCOPE)
    .setParam('access_type', 'offline')
    .setParam('prompt', 'consent');
}

/**
 * Ensures the user has supplied credentials before attempting OAuth.
 */
function requireOAuthCredentials() {
  const creds = getStoredOAuthCredentials();
  if (!creds.clientId || !creds.clientSecret) {
    const error = new Error('Enter your QuickBooks Client ID and Client Secret before connecting.');
    error.name = 'MissingCredentialsError';
    throw error;
  }
  return creds;
}

/**
 * Generates an authorization URL that the UI opens in a popup.
 */
function getAuthorizationUrl() {
  requireOAuthCredentials();
  const service = getOAuthService();
  return service.getAuthorizationUrl();
}

/**
 * Handles the Intuit OAuth callback.
 */
function authCallback(request) {
  const service = getOAuthService();
  let htmlContent;

  try {
    const authorized = service.handleCallback(request);
    if (authorized) {
      const realmId = (request.parameter && request.parameter.realmId) || '';
      const props = getUserProps();

      if (realmId) {
        props.setProperty(PROPERTY_REALM_ID, realmId);
      }
      props.setProperty(PROPERTY_LAST_CONNECTED, new Date().toISOString());

      htmlContent = '<p>QuickBooks connection successful. You can close this tab.</p>';
    } else {
      const errorDetail = service.getLastError() || 'Authorization was denied.';
      htmlContent = '<p>QuickBooks connection failed.</p><p>' + sanitizeHtml(errorDetail) + '</p>';
    }
  } catch (error) {
    htmlContent = '<p>Unexpected error during QuickBooks authorization.</p><p>' + sanitizeHtml(error.message) + '</p>';
  }

  return HtmlService.createHtmlOutput(htmlContent)
    .setTitle('QuickBooks Connection');
}

/**
 * Removes any stored tokens and connection metadata.
 */
function disconnect() {
  const service = getOAuthService();
  try {
    service.reset();
  } finally {
    clearConnectionMetadata();
  }
}

/**
 * Clears OAuth state without clearing saved credentials.
 */
function resetOAuthState() {
  clearConnectionMetadata();
  const service = getOAuthService();
  service.reset();
}

/**
 * Returns a lightweight snapshot of the connection.
 */
function getConnectionStatus() {
  const props = getUserProps();
  const { clientId, clientSecret, environment } = getStoredOAuthCredentials();
  const service = getOAuthService();
  const isConnected = clientId && clientSecret ? service.hasAccess() : false;

  return {
    hasCredentials: Boolean(clientId && clientSecret),
    clientId: clientId ? maskCredential(clientId) : '',
    clientSecret: clientSecret ? maskCredential(clientSecret) : '',
    environment: environment,
    isConnected: isConnected,
    realmId: props.getProperty(PROPERTY_REALM_ID) || null,
    companyName: props.getProperty(PROPERTY_COMPANY_NAME) || null,
    lastConnectedAt: props.getProperty(PROPERTY_LAST_CONNECTED) || null
  };
}

/**
 * Stores company metadata once we learn it from the API.
 */
function setCompanyMetadata(realmId, companyName) {
  const props = getUserProps();
  if (realmId) {
    props.setProperty(PROPERTY_REALM_ID, realmId);
  }
  if (companyName) {
    props.setProperty(PROPERTY_COMPANY_NAME, companyName);
  }
}

/**
 * Masks credentials so we can safely echo them back into the UI.
 */
function maskCredential(value) {
  const trimmed = String(value || '').trim();
  if (trimmed.length <= 4) {
    return '••••';
  }
  const visibleTail = trimmed.slice(-4);
  return '••••••••' + visibleTail;
}

/**
 * Simple sanitiser for callback messaging.
 */
function sanitizeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function isConnected() {
  return getConnectionStatus().isConnected;
}

function getConnectionDetails() {
  return getConnectionStatus();
}

function getOAuthCredentials() {
  const { clientId, clientSecret } = getStoredOAuthCredentials();
  return {
    clientId: clientId,
    hasClientSecret: Boolean(clientSecret)
  };
}
