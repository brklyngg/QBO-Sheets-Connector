function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('QBO Connector')
    .addItem('Open', 'showSidebar')
    .addItem('Run: P&L (YTD)', 'demoRunPnL_')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar_Reports')
    .setTitle('QBO Reports');
  SpreadsheetApp.getUi().showSidebar(html);
}

function uiGetStatus() {
  return qboStatus();
}

function uiRunReport(reportName, params, sheetName) {
  try {
    const result = QBO_API.reportToSheet(reportName, params || {}, sheetName || ('QBO_' + reportName));
    return {
      ok: true,
      msg: 'Report ' + reportName + ' → ' + result.sheet + ' (' + result.rows + '×' + result.cols + ')'
    };
  } catch (error) {
    return {
      ok: false,
      msg: error.message
    };
  }
}

function demoRunPnL_() {
  const timezone = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');
  const jan1 = Utilities.formatDate(new Date(new Date().getFullYear(), 0, 1), timezone, 'yyyy-MM-dd');
  const result = uiRunReport('ProfitAndLoss', {
    start_date: jan1,
    end_date: today,
    summarize_column_by: 'Month',
    accounting_method: 'Accrual'
  }, 'QBO_PnL_YTD');

  if (!result.ok) {
    throw new Error(result.msg);
  }
}
