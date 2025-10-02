const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const USER_PROPS = PropertiesService.getUserProperties();

const QBO_KEYS = {
  REALM_ID: 'QBO_REALM_ID',
  ENVIRONMENT: 'ENVIRONMENT',
  MINOR: 'QBO_MINOR_VERSION',
  CLIENT_ID: 'QBO_CLIENT_ID',
  CLIENT_SECRET: 'QBO_CLIENT_SECRET'
};

function getConfig_() {
  return {
    clientId: SCRIPT_PROPS.getProperty(QBO_KEYS.CLIENT_ID) || '',
    clientSecret: SCRIPT_PROPS.getProperty(QBO_KEYS.CLIENT_SECRET) || '',
    env: SCRIPT_PROPS.getProperty(QBO_KEYS.ENVIRONMENT) || 'sandbox',
    minor: SCRIPT_PROPS.getProperty(QBO_KEYS.MINOR) || '75',
    realmId: USER_PROPS.getProperty(QBO_KEYS.REALM_ID) || ''
  };
}

function setConfig_(kv) {
  Object.keys(kv || {}).forEach(key => {
    const val = kv[key];
    if (val != null) {
      SCRIPT_PROPS.setProperty(key, String(val));
    }
  });
}

function getQboService_() {
  const cfg = getConfig_();
  if (!cfg.clientId || !cfg.clientSecret) {
    throw new Error('Missing QBO client credentials. Set script properties QBO_CLIENT_ID and QBO_CLIENT_SECRET.');
  }
  return OAuth2.createService('intuit-qbo')
    .setAuthorizationBaseUrl('https://appcenter.intuit.com/connect/oauth2')
    .setTokenUrl('https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer')
    .setClientId(cfg.clientId)
    .setClientSecret(cfg.clientSecret)
    .setCallbackFunction('authCallback')
    .setScope('com.intuit.quickbooks.accounting')
    .setParam('access_type', 'offline')
    .setParam('prompt', 'consent')
    .setPropertyStore(USER_PROPS)
    .setLock(LockService.getUserLock());
}

function authCallback(e) {
  const svc = getQboService_();
  const handled = svc.handleCallback(e);
  if (handled) {
    const realmId = e && e.parameter && e.parameter.realmId ? e.parameter.realmId : '';
    if (realmId) {
      USER_PROPS.setProperty(QBO_KEYS.REALM_ID, realmId);
      try {
        svc.getStorage().setValue(QBO_KEYS.REALM_ID, realmId);
      } catch (err) {
        console.error('Failed to persist realmId into OAuth storage', err);
      }
    }
    return HtmlService.createHtmlOutput('Success. You can close this tab.');
  }
  return HtmlService.createHtmlOutput('Access denied. You can close this tab.');
}

function qboStatus() {
  const cfg = getConfig_();
  try {
    const svc = getQboService_();
    return {
      isAuthed: svc.hasAccess(),
      authUrl: svc.getAuthorizationUrl(),
      realmId: cfg.realmId,
      env: cfg.env,
      minor: cfg.minor
    };
  } catch (error) {
    return {
      isAuthed: false,
      authUrl: '',
      realmId: cfg.realmId,
      env: cfg.env,
      minor: cfg.minor,
      error: error.message
    };
  }
}

function isUserAuthorized() {
  try {
    return getQboService_().hasAccess();
  } catch (error) {
    return false;
  }
}

function getAuthorizationUrl() {
  return getQboService_().getAuthorizationUrl();
}

function logout() {
  try {
    const svc = getQboService_();
    svc.reset();
  } catch (error) {
    console.error('Failed to reset OAuth service during logout', error);
  }
  USER_PROPS.deleteProperty(QBO_KEYS.REALM_ID);
}

function QBO_adminSetConfig(props) {
  setConfig_(props);
}

function getUserProps() {
  return USER_PROPS;
}
