# QuickBooks Online Connector for Google Sheets

A production-ready Google Sheets add-on that connects to QuickBooks Online for importing financial data with scheduling capabilities.

## Features

- 🔐 **OAuth 2.0 Authentication** - Secure connection to QuickBooks Online
- 📊 **Standard Reports** - Import P&L, Balance Sheet, Trial Balance, and more
- 🔍 **Custom Queries** - Write SQL-like queries to fetch specific data
- ⏰ **Scheduled Refreshes** - Automatic data updates (hourly, daily, weekly, monthly)
- 📝 **Comprehensive Logging** - Track all operations and troubleshoot issues
- 🎯 **Named Ranges** - Stable data placement for formulas
- 🚦 **Error Handling** - Automatic retries and clear error messages

## Supported Reports

- Profit and Loss
- Balance Sheet
- Trial Balance
- Transaction List by Date
- Sales by Customer

## Prerequisites

1. Google account with access to Google Sheets
2. QuickBooks Online account (or sandbox account for testing)
3. QuickBooks app credentials from [developer.intuit.com](https://developer.intuit.com)

## Installation

### Method 1: From Source (Recommended for Development)

1. Open Google Sheets
2. Go to **Extensions > Apps Script**
3. Delete any existing code
4. Create the following files and copy the contents:
   - `Code.gs`
   - `Auth.gs`
   - `QBO.gs`
   - `Datasets.gs`
   - `Scheduler.gs`
   - `Logging.gs`
   - `UI.gs`
   - `UI.html`
5. Copy `appsscript.json` to configure manifest
6. Save the project with a name like "QBO Connector"
7. Run `onInstall()` function once to initialize

### Method 2: Deploy as Add-on

1. Follow Method 1 to set up the code
2. In Apps Script editor, click **Deploy > Test deployments**
3. Select **Install add-on** for testing
4. For production deployment, follow [Google's add-on publishing guide](https://developers.google.com/workspace/add-ons/how-tos/publish-add-on-overview)

## Setup

### 1. Create QuickBooks App

1. Go to [developer.intuit.com](https://developer.intuit.com)
2. Sign in and go to **My Apps**
3. Click **Create an app**
4. Select **QuickBooks Online and Payments**
5. Configure your app:
   - **Redirect URI**: `https://script.google.com/macros/d/{SCRIPT_ID}/usercallback`
   - **Scopes**: Select `com.intuit.quickbooks.accounting`
6. Save your **Client ID** and **Client Secret**

### 2. Configure the Add-on

1. Open your Google Sheet
2. Go to **Extensions > QuickBooks Online Connector > Open Connector**
3. Click the **Settings** tab
4. Enter your Client ID and Client Secret
5. Click **Save Settings**

### 3. Connect to QuickBooks

1. In the sidebar, click **Connect**
2. Authorize the app in the QuickBooks popup
3. Select your company (or sandbox company)
4. You should see "Connected to QuickBooks" status

## Usage

### Creating a Dataset

1. Click **+ New Dataset** in the sidebar
2. Choose dataset type:
   - **Standard Report**: Pre-built QuickBooks reports
   - **Custom Query**: SQL-like queries for specific data
3. Configure parameters:
   - For reports: Select report type and date range
   - For queries: Write your query (e.g., `SELECT * FROM Customer`)
4. Set target sheet name (optional)
5. Click **Create Dataset**

### Running Datasets

1. Select a dataset from the list
2. Click **Run Now** to execute immediately
3. Monitor progress in the status bar
4. Data will appear in the specified sheet

### Scheduling Refreshes

1. Select a dataset
2. Enable **Scheduled refresh**
3. Choose frequency:
   - Hourly: Runs every hour
   - Daily: Runs at specified time
   - Weekly: Runs on specified day and time
   - Monthly: Runs on specified day of month
4. Click **Update Schedule**

### Custom Queries

Query syntax follows QuickBooks API query language:

```sql
SELECT * FROM Customer WHERE Active = true
SELECT * FROM Invoice WHERE TxnDate > '2024-01-01' ORDERBY TxnDate DESC
SELECT COUNT(*) FROM Item WHERE Type = 'Inventory'
```

Supported entities: Account, Bill, Customer, Employee, Invoice, Item, Payment, Vendor, and more.

## Logging

All operations are logged to the `QBO_Connector_Logs` sheet:
- API calls with response times
- Errors with detailed messages
- Dataset runs with row counts
- Schedule executions

Access logs via the **Logs** tab or **View Logs** menu item.

## Troubleshooting

### Common Issues

1. **"Not authenticated" error**
   - Click **Connect** to re-authenticate
   - Check Client ID and Secret in Settings

2. **"Rate limit exceeded" (429 error)**
   - Wait for the specified retry time
   - Reduce frequency of scheduled refreshes

3. **"No data returned"**
   - Check date ranges in report parameters
   - Verify entity exists for queries
   - Check QuickBooks data permissions

4. **Schedule not running**
   - Verify timezone settings match
   - Check Google Apps Script trigger quotas
   - Look for errors in logs

### Debug Mode

For detailed debugging:
1. Open Apps Script editor
2. View > Logs for script logs
3. Check QBO_Connector_Logs sheet

## Best Practices

1. **Data Limits**
   - Keep datasets under 2M cells (warning)
   - Hard limit: 8M cells per dataset
   - File limit: 10M cells total

2. **Scheduling**
   - Avoid scheduling multiple large datasets at the same time
   - Use hourly schedules sparingly to avoid rate limits
   - Monitor logs for failed scheduled runs

3. **Queries**
   - Always include MAXRESULTS to limit data
   - Use date filters to reduce data size
   - Test queries with small result sets first

4. **Security**
   - Never share Client Secret
   - Use sandbox for testing
   - Regularly review connected apps in QuickBooks

## API Limits

QuickBooks Online API limits:
- 500 requests per minute per realm
- 40 simultaneous connections per app
- Response timeout: 120 seconds

Google Apps Script limits:
- 6 minutes per execution
- 20 triggers per user
- 500KB property storage

## Support

For issues or questions:
1. Check logs for detailed error messages
2. Review [QuickBooks API documentation](https://developer.intuit.com/app/developer/qbo/docs/get-started)
3. Submit issues to the GitHub repository

## License

This project is provided as-is for use with QuickBooks Online and Google Sheets.

## Version History

- **1.0.0** - Initial release
  - OAuth 2.0 authentication
  - Standard reports and custom queries
  - Scheduling system
  - Comprehensive logging
