# Google Sheets Configuration (Apps Script Method)

## Spreadsheet Structure

The app uses **two sheets** inside the same spreadsheet:

| Sheet name  | Purpose |
|-------------|---------|
| **Jumps**   | One row per jump log entry (jump #, date, location, equipment, notes, …) |
| **Equipment** | Auto-created by the Apps Script. Stores all equipment components (harnesses, canopies, rigs) and settings as JSON. This is what enables multi-device sync. |

The **Equipment** sheet is read and written by the app automatically — you should not edit it by hand.

## Setup Instructions

> **Already set up?** If you previously deployed the Apps Script, you must redeploy it (new version) to pick up the new `getEquipment` / `saveEquipment` actions. The Equipment sheet tab will be created automatically on first sync.

1. **Create a Google Spreadsheet:**
   - Go to [Google Sheets](https://sheets.google.com)
   - Create a new spreadsheet
   - Rename it to "Skydiving Logbook"
   - Rename the first sheet to "Jumps"
   - Add headers in row 1: `Jump Number`, `Date`, `Location`, `Equipment`, `Notes`, `Timestamp`

2. **Get your Spreadsheet ID:**
   - Copy the ID from your spreadsheet URL: `https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit`
   - The SPREADSHEET_ID is the long string between `/d/` and `/edit`

3. **Create a Google Apps Script:**
   - While in your spreadsheet, go to `Extensions` → `Apps Script`
   - Delete any existing code in the script editor
   - Copy and paste the code from `config/apps-script.js`
   - Save the script (give it a name like "Skydiving Logbook API")

4. **Deploy the Apps Script:**
   - In the Apps Script editor, click `Deploy` → `New deployment`
   - Choose type: `Web app`
   - Execute as: `Me`
   - Who has access: `Anyone`
   - Click `Deploy`
   - Copy the Web app URL that's generated

5. **Configure the App:**
   - Edit `config/sheets-config.json`
   - Replace `YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL` with your Web app URL
   - Replace `YOUR_SPREADSHEET_ID` with your spreadsheet ID

## Example Configuration

```json
{
  "webAppUrl": "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec",
  "spreadsheetId": "YOUR_SPREADSHEET_ID"
}
```

## Troubleshooting
- If sync fails, check your Web app URL and spreadsheet ID
- Ensure the Apps Script is deployed with "Anyone" access
- Make sure your spreadsheet has a "Jumps" sheet with the correct headers
- Check browser console for error messages

## Why Apps Script?
Google changed their API policy - API keys only work for reading data, but writing requires OAuth2 authentication. Apps Script provides a simpler alternative for personal projects.