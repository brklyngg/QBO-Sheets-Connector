/**
 * OAuth 2.0 Authentication for QuickBooks Online
 * Uses the OAuth2 library for Google Apps Script
 */

// OAuth constants
const OAUTH_CLIENT_ID = PropertiesService.getUserProperties().getProperty('QBO_CLIENT_ID') || '';
const OAUTH_CLIENT_SECRET = PropertiesService.getUserProperties().getProperty('QBO_CLIENT_SECRET') || '';
const OAUTH_CALLBACK_FUNCTION = 'authCallback';
const OAUTH_SCOPE = 'com.intuit.quickbooks.accounting';

// Intuit OAuth URLs
const INTUIT_AUTH_URL = 'https://appcenter.intuit.com/connect/oauth2';
const INTUIT_TOKEN_URL = 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer';
const INTUIT_REVOKE_URL = 'https://developer.api.intuit.com/v2/oauth2/tokens/revoke';
const INTUIT_DISCOVERY_URL = 'https://developer.api.intuit.com/.well-known/openid_configuration';

/**
 * Gets the OAuth2 service instance
 */
function getOAuthService() {
  return OAuth2.createService('QuickBooksOnline')
    .setAuthorizationBaseUrl(INTUIT_AUTH_URL)
    .setTokenUrl(INTUIT_TOKEN_URL)
    .setClientId(OAUTH_CLIENT_ID)
    .setClientSecret(OAUTH_CLIENT_SECRET)
    .setCallbackFunction(OAUTH_CALLBACK_FUNCTION)
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope(OAUTH_SCOPE)
    .setParam('access_type', 'offline')
    .setParam('prompt', 'consent')
    .setTokenHeaders({
      'Authorization': 'Basic ' + Utilities.base64Encode(OAUTH_CLIENT_ID + ':' + OAUTH_CLIENT_SECRET),
      'Accept': 'application/json',
      'Content-Type': 'application/x-www-form-urlencoded'
    });
}

/**
 * OAuth callback handler
 */
function authCallback(request) {
  try {
    const service = getOAuthService();
    const isAuthorized = service.handleCallback(request);
    
    if (isAuthorized) {
      // Extract and store the realmId from the callback
      const realmId = request.parameter.realmId;
      if (realmId) {
        PropertiesService.getUserProperties().setProperty('QBO_REALM_ID', realmId);
        logAction('auth_callback_success', { realmId: realmId });
      }
      
      return HtmlService.createHtmlOutput(getSuccessHTML());
    } else {
      logAction('auth_callback_failed', { 
        error: 'Authorization failed',
        parameters: Object.keys(request.parameter || {})
      });
      return HtmlService.createHtmlOutput(getFailureHTML('Authorization failed'));
    }
  } catch (error) {
    console.error('Auth callback error:', error);
    logAction('auth_callback_error', { 
      error: error.toString(),
      stack: error.stack 
    });
    return HtmlService.createHtmlOutput(getFailureHTML(error.toString()));
  }
}

/**
 * Initiates the OAuth flow
 */
function getAuthorizationUrl() {
  try {
    const service = getOAuthService();
    
    if (!OAUTH_CLIENT_ID || !OAUTH_CLIENT_SECRET) {
      throw new Error('OAuth credentials not configured. Please set up your QuickBooks app credentials in Settings.');
    }
    
    const authUrl = service.getAuthorizationUrl();
    
    logAction('get_auth_url', { 
      hasCredentials: !!(OAUTH_CLIENT_ID && OAUTH_CLIENT_SECRET)
    });
    
    return authUrl;
  } catch (error) {
    console.error('Error getting auth URL:', error);
    throw error;
  }
}

/**
 * Checks if the user is connected to QuickBooks
 */
function isConnected() {
  try {
    const service = getOAuthService();
    const hasAccess = service.hasAccess();
    const realmId = PropertiesService.getUserProperties().getProperty('QBO_REALM_ID');
    
    return hasAccess && !!realmId;
  } catch (error) {
    console.error('Error checking connection:', error);
    return false;
  }
}

/**
 * Gets connection details for the UI
 */
function getConnectionDetails() {
  try {
    const service = getOAuthService();
    const hasAccess = service.hasAccess();
    const realmId = PropertiesService.getUserProperties().getProperty('QBO_REALM_ID');
    const companyName = PropertiesService.getUserProperties().getProperty('QBO_COMPANY_NAME');
    
    if (!hasAccess) {
      return {
        connected: false,
        message: 'Not connected to QuickBooks'
      };
    }
    
    const token = service.getAccessToken();
    const expiresIn = service.getExpirationTime();
    const expiresAt = expiresIn ? new Date(expiresIn) : null;
    
    return {
      connected: true,
      realmId: realmId,
      companyName: companyName || 'Unknown Company',
      expiresAt: expiresAt ? expiresAt.toISOString() : null,
      tokenType: token ? 'Bearer' : null
    };
  } catch (error) {
    console.error('Error getting connection details:', error);
    return {
      connected: false,
      error: error.toString()
    };
  }
}

/**
 * Disconnects from QuickBooks
 */
function disconnect() {
  try {
    const service = getOAuthService();
    const token = service.getAccessToken();
    
    if (token) {
      // Revoke the token at Intuit
      try {
        const revokeUrl = INTUIT_REVOKE_URL;
        const response = UrlFetchApp.fetch(revokeUrl, {
          method: 'POST',
          headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(OAUTH_CLIENT_ID + ':' + OAUTH_CLIENT_SECRET),
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
          },
          payload: 'token=' + token,
          muteHttpExceptions: true
        });
        
        logAction('disconnect_revoke', {
          status: response.getResponseCode(),
          success: response.getResponseCode() === 200
        });
      } catch (revokeError) {
        console.error('Error revoking token:', revokeError);
      }
    }
    
    // Clear local storage
    service.reset();
    PropertiesService.getUserProperties().deleteProperty('QBO_REALM_ID');
    PropertiesService.getUserProperties().deleteProperty('QBO_COMPANY_NAME');
    
    logAction('disconnect_success');
    
    return {
      success: true,
      message: 'Successfully disconnected from QuickBooks'
    };
  } catch (error) {
    console.error('Error disconnecting:', error);
    logAction('disconnect_error', { error: error.toString() });
    
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Refreshes the access token if needed
 */
function refreshAccessToken() {
  try {
    const service = getOAuthService();
    
    if (service.hasAccess()) {
      // Check if token needs refresh (within 30 minutes of expiry)
      const expirationTime = service.getExpirationTime();
      const now = new Date().getTime();
      const thirtyMinutes = 30 * 60 * 1000;
      
      if (expirationTime && (expirationTime - now) < thirtyMinutes) {
        // Force refresh
        service.refresh();
        logAction('token_refreshed', {
          newExpiration: new Date(service.getExpirationTime()).toISOString()
        });
      }
      
      return true;
    }
    
    return false;
  } catch (error) {
    console.error('Error refreshing token:', error);
    logAction('token_refresh_error', { error: error.toString() });
    return false;
  }
}

/**
 * Makes an authenticated request to QuickBooks API
 */
function makeAuthenticatedRequest(url, options = {}) {
  try {
    const service = getOAuthService();
    
    if (!service.hasAccess()) {
      throw new Error('Not authenticated. Please connect to QuickBooks first.');
    }
    
    // Refresh token if needed
    refreshAccessToken();
    
    const token = service.getAccessToken();
    const realmId = PropertiesService.getUserProperties().getProperty('QBO_REALM_ID');
    
    if (!realmId) {
      throw new Error('No QuickBooks company selected. Please reconnect.');
    }
    
    // Build request options
    const requestOptions = {
      ...options,
      headers: {
        'Authorization': 'Bearer ' + token,
        'Accept': 'application/json',
        'Content-Type': options.method === 'POST' ? 'application/json' : 'application/x-www-form-urlencoded',
        ...(options.headers || {})
      },
      muteHttpExceptions: true
    };
    
    // Make the request
    const response = UrlFetchApp.fetch(url, requestOptions);
    const responseCode = response.getResponseCode();
    
    // Handle 401 - token expired or invalid
    if (responseCode === 401) {
      logAction('auth_401_error', { url: url });
      
      // Try to refresh and retry once
      service.refresh();
      const newToken = service.getAccessToken();
      requestOptions.headers['Authorization'] = 'Bearer ' + newToken;
      
      const retryResponse = UrlFetchApp.fetch(url, requestOptions);
      return retryResponse;
    }
    
    return response;
  } catch (error) {
    console.error('Authenticated request error:', error);
    throw error;
  }
}

/**
 * Gets the success HTML for OAuth callback
 */
function getSuccessHTML() {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            text-align: center;
            padding: 50px;
            background-color: #f0f0f0;
          }
          .container {
            background-color: white;
            padding: 40px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            max-width: 400px;
            margin: 0 auto;
          }
          .success-icon {
            color: #4CAF50;
            font-size: 60px;
            margin-bottom: 20px;
          }
          h1 {
            color: #333;
            margin-bottom: 10px;
          }
          p {
            color: #666;
            margin-bottom: 30px;
          }
          button {
            background-color: #4285F4;
            color: white;
            border: none;
            padding: 12px 30px;
            border-radius: 5px;
            font-size: 16px;
            cursor: pointer;
            transition: background-color 0.3s;
          }
          button:hover {
            background-color: #357abd;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="success-icon">✓</div>
          <h1>Successfully Connected!</h1>
          <p>You have successfully connected to QuickBooks Online.</p>
          <button onclick="window.close()">Close Window</button>
        </div>
        <script>
          // Try to close the window after 3 seconds
          setTimeout(() => {
            window.close();
          }, 3000);
        </script>
      </body>
    </html>
  `;
}

/**
 * Gets the failure HTML for OAuth callback
 */
function getFailureHTML(error) {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            text-align: center;
            padding: 50px;
            background-color: #f0f0f0;
          }
          .container {
            background-color: white;
            padding: 40px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            max-width: 400px;
            margin: 0 auto;
          }
          .error-icon {
            color: #f44336;
            font-size: 60px;
            margin-bottom: 20px;
          }
          h1 {
            color: #333;
            margin-bottom: 10px;
          }
          p {
            color: #666;
            margin-bottom: 20px;
          }
          .error-details {
            background-color: #f5f5f5;
            padding: 15px;
            border-radius: 5px;
            text-align: left;
            font-size: 14px;
            color: #666;
            margin-bottom: 30px;
          }
          button {
            background-color: #f44336;
            color: white;
            border: none;
            padding: 12px 30px;
            border-radius: 5px;
            font-size: 16px;
            cursor: pointer;
            transition: background-color 0.3s;
          }
          button:hover {
            background-color: #d32f2f;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="error-icon">✕</div>
          <h1>Connection Failed</h1>
          <p>Unable to connect to QuickBooks Online.</p>
          <div class="error-details">
            ${error ? `Error: ${error}` : 'Unknown error occurred'}
          </div>
          <button onclick="window.close()">Close Window</button>
        </div>
      </body>
    </html>
  `;
}

/**
 * Sets OAuth credentials (called from settings)
 */
function setOAuthCredentials(clientId, clientSecret) {
  try {
    if (!clientId || !clientSecret) {
      throw new Error('Both Client ID and Client Secret are required');
    }
    
    PropertiesService.getUserProperties().setProperty('QBO_CLIENT_ID', clientId);
    PropertiesService.getUserProperties().setProperty('QBO_CLIENT_SECRET', clientSecret);
    
    logAction('set_oauth_credentials', {
      hasClientId: !!clientId,
      hasClientSecret: !!clientSecret
    });
    
    return {
      success: true,
      message: 'OAuth credentials saved successfully'
    };
  } catch (error) {
    console.error('Error setting OAuth credentials:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Gets OAuth credentials (for settings UI)
 */
function getOAuthCredentials() {
  return {
    clientId: PropertiesService.getUserProperties().getProperty('QBO_CLIENT_ID') || '',
    hasClientSecret: !!PropertiesService.getUserProperties().getProperty('QBO_CLIENT_SECRET')
  };
}
