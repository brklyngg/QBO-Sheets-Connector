/**
 * Minimal Spreadsheet entry points for exercising the QuickBooks connection.
 */

const SCRIPT_VERSION = '1.0.1';
const SIDEBAR_TITLE = 'QBO Connection Tester';
const SIDEBAR_WIDTH = 320;

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(`QBO Connection v${SCRIPT_VERSION}`)
    .addItem('Open Connection Tester', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const template = HtmlService.createTemplateFromFile('UI');
  template.initialStateJson = JSON.stringify(loadSidebarData());

  const htmlOutput = template.evaluate()
    .setTitle(SIDEBAR_TITLE)
    .setWidth(SIDEBAR_WIDTH);

  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function loadSidebarData() {
  return {
    status: getConnectionStatus(),
    redirectUri: getOAuthRedirectUri()
  };
}

function saveConnectionSettings(settings) {
  if (!settings) {
    throw new Error('No settings payload supplied.');
  }

  const clientId = settings.clientId || '';
  const clientSecret = settings.clientSecret || '';
  const environment = settings.environment || 'sandbox';

  setOAuthCredentials(clientId, clientSecret, environment);
  return getConnectionStatus();
}

function startAuthorization() {
  return {
    authorizationUrl: getAuthorizationUrl()
  };
}

function disconnectFromQuickBooks() {
  disconnect();
  return getConnectionStatus();
}

function refreshConnectionStatus() {
  return getConnectionStatus();
}

function runConnectionTest() {
  return testQuickBooksConnection();
}

function getOAuthRedirectUri() {
  return 'https://script.google.com/macros/d/' + ScriptApp.getScriptId() + '/usercallback';
}

/**
 * Allows templated includes within UI.html.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
