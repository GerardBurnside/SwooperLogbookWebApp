# Google Sheets Configuration (OAuth Method)

## Overview

The app uses **Google Sign-In (OAuth 2.0)** to connect directly to the Google Sheets API. When you sign in with your Google account, the app creates a spreadsheet in your Google Drive and syncs your data automatically.

**No Apps Script deployment needed** — just sign in and go.

## Spreadsheet Structure

The app auto-creates a spreadsheet with **two sheets**:

| Sheet name  | Purpose |
|-------------|---------|
| **Jumps**   | One row per jump log entry (jump #, date, location, equipment, notes, …) |
| **Equipment** | Stores all equipment components (harnesses, canopies, locations) and settings as JSON. |

The **Equipment** sheet is managed automatically — do not edit it by hand.

## Setup Instructions

There are two roles:

- **App owner (deployer):** Creates the Google Cloud project and sets the OAuth Client ID once. Steps 1–5 below.
- **End users:** Just open the app and click **Sign in with Google**. No Cloud Console access needed — the owner adds them as test users.

### Step 1: Create a Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Click **Select a project** → **New Project**
3. Name it (e.g. "Swooper Logbook") and click **Create**

### Step 2: Enable the Google Sheets API

1. In the Cloud Console, go to **APIs & Services** → **Library**
2. Search for **Google Sheets API**
3. Click on it and press **Enable**

### Step 3: Configure the OAuth Consent Screen

1. Go to **APIs & Services** → **OAuth consent screen**
2. Select **External** user type → **Create**
3. Fill in the required fields:
   - **App name:** Swooper Logbook
   - **User support email:** your email
   - **Developer contact email:** your email
4. Click **Save and Continue**
5. On the **Scopes** page, click **Add or Remove Scopes**
   - Search for `drive.file` and check **`../auth/drive.file`**
   - Click **Update** then **Save and Continue**
6. On the **Test users** page:
   - Click **+ Add Users**
   - Add the Google account emails for yourself and anyone else who will use the app
   - Click **Save and Continue**
7. Click **Back to Dashboard**

> **Note:** In "Testing" mode you can add up to 100 test users. For public use with >100 users, you must submit for [Google verification](https://support.google.com/cloud/answer/9110914) (requires a privacy policy URL).

### Step 4: Create OAuth Client ID

1. Go to **APIs & Services** → **Credentials**
2. Click **+ Create Credentials** → **OAuth client ID**
3. Application type: **Web application**
4. Name: "Swooper Logbook Web"
5. Under **Authorized JavaScript origins**, add all domains where the app will be hosted:
   - `http://localhost` (for development)
   - `http://localhost:8080` (or whatever port you use)
   - `https://your-domain.com` (your production domain)
6. Leave **Authorized redirect URIs** empty (not needed for GIS)
7. Click **Create**
8. **Copy the Client ID**

### Step 5: Set the Client ID in the App

Open `js/auth.js` and paste your Client ID into the constant at the top of the file:

```js
const OAUTH_CLIENT_ID = 'YOUR_CLIENT_ID.apps.googleusercontent.com';
```

Deploy the app. That's it — users just click **Sign in with Google**.

### For End Users

1. The app owner adds your Google account email as a test user (Step 3.6 above)
2. Open the app → **Settings** → **Google Sheet Integration**
3. Click **Sign in with Google**
4. Grant the requested permissions
5. The app creates a spreadsheet in your Drive and starts syncing!

## Migrating from Apps Script

If you previously used the Apps Script deployment method:

1. Open **Settings** → **Google Sheet Integration**
2. You'll see a migration banner
3. Enter your OAuth Client ID and click **Sign in with Google**
4. The app will create a new spreadsheet and push all your local data to it
5. Your old Apps Script config is automatically cleared
6. You can delete the old Apps Script deployment from your Google Sheets

> All your data is preserved locally — the migration just creates a new cloud sheet.

## Troubleshooting

- **"Sign-in failed" or popup closes immediately:** Verify your OAuth Client ID is correct and that your domain is listed in Authorized JavaScript origins
- **"Access blocked" in Google popup:** Add your Google account as a test user in the OAuth consent screen
- **Sync stops working after ~1 hour:** The app should silently refresh the token. If it doesn't, click the sync button to re-authenticate
- **Want to use a different Google account:** Sign out first, then sign in with the other account
- Check browser console for error messages

## Self-Hosting

If you're self-hosting, you have two options:

1. **Recommended:** Create your own Google Cloud project (Steps 1–5 above) and set `OAUTH_CLIENT_ID` in `js/auth.js` before deploying.
2. **Alternative:** Leave `OAUTH_CLIENT_ID` empty. The app will show a Client ID input field in the integration modal so each user can enter one manually (useful for development).

## Legacy: Apps Script Method

The old Apps Script method (`config/apps-script.js`) is deprecated but still present in the codebase for reference. New installations should use the OAuth method described above.