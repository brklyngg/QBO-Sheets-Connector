# QuickBooks Connection Tester

This branch contains a stripped-down Google Sheets project focused solely on validating the QuickBooks Online OAuth flow. The goal is to confirm that a Sheets sidebar can launch the Intuit consent screen, capture the callback, store tokens, and make a simple API request against your sandbox company.

## What You Get
- Minimal custom menu that opens a "QBO Connection Tester" sidebar
- Credential storage (Client ID, Client Secret, sandbox vs. production toggle)
- QuickBooks OAuth 2.0 flow via the official Apps Script OAuth2 library
- Connection status summary with stored realm ID and company name
- One API probe (`CompanyInfo`) to prove access with the saved token
- Quick disconnect/reset button so you can start over quickly

Everything unrelated to authentication (datasets, scheduling, logging, etc.) has been removed so there are no other moving parts while validating the connection.

## Setup
1. **Copy the project into Apps Script**
   - Open the Google Sheet you want to use for testing.
   - Go to **Extensions → Apps Script** and paste the contents of this branch into the Script Editor (`Code.js`, `Auth.js`, `QBO.js`, `UI.html`, `appsscript.json`).
   - Ensure the OAuth2 library is added (`Resources → Libraries → Script ID 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF`).

2. **Create a QuickBooks app (sandbox works great)**
   - https://developer.intuit.com → **My Apps** → **Create an app**.
   - Add the sheet's callback URL to the Redirect URI list: `https://script.google.com/macros/d/<SCRIPT_ID>/usercallback`.
   - Enable the `com.intuit.quickbooks.accounting` scope.
   - Grab the Client ID and Client Secret from the sandbox keys tab.

3. **Authorize the script**
   - Back in the sheet, reload and accept the new **QBO Connection** menu.
   - Choose **Open Connection Tester**, paste your Intuit credentials, select **Sandbox**, and click **Save Credentials**.
   - Click **Connect to QuickBooks**, complete the OAuth popup, then hit **Refresh Status** in the sidebar.

## Verifying the Flow
- The status pill will turn green once tokens are stored.
- Click **Test Company Info** to call `CompanyInfo` and confirm the sandbox company name and realm ID.
- Use **Disconnect** whenever you want to wipe tokens and try again.

## Next Steps
Once the connection is working end-to-end, we can graft the production-ready reporting features back in with confidence that OAuth is behaving as expected.

## Deploying updates
- Run `npm run clasp:push` from the project root whenever you are ready to push code to Apps Script. The helper script automatically bumps the `SCRIPT_VERSION` constant in `Code.js`, so the custom menu reads `QBO Connection vX.Y.Z` and confirms the latest build made it into the sheet.
