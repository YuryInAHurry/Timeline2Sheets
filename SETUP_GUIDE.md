# Google Sheets & Service Account Setup Guide

This guide walks you through setting up Google Sheets API access and creating a service account for the Timeline to Logbook converter.

## Table of Contents
1. [Create a Google Cloud Project](#1-create-a-google-cloud-project)
2. [Enable Required APIs](#2-enable-required-apis)
3. [Create a Service Account](#3-create-a-service-account)
4. [Get Your Google Maps API Key](#4-get-your-google-maps-api-key)
5. [Create and Share Your Google Sheet](#5-create-and-share-your-google-sheet)
6. [Configure the Script](#6-configure-the-script)

---

## 1. Create a Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Sign in with your Google account
3. Click **Select a project** dropdown at the top
4. Click **NEW PROJECT**
5. Enter a project name (e.g., "Timeline Logbook Converter")
6. Click **CREATE**
7. Wait for the project to be created (you'll see a notification)
8. Make sure your new project is selected in the dropdown

---

## 2. Enable Required APIs

### Enable Google Sheets API

1. In the Google Cloud Console, go to **APIs & Services** > **Library**
   - Or visit: https://console.cloud.google.com/apis/library
2. Search for "Google Sheets API"
3. Click on **Google Sheets API**
4. Click **ENABLE**
5. Wait for it to enable (takes a few seconds)

### Enable Google Maps Geocoding API

1. Still in **APIs & Services** > **Library**
2. Search for "Geocoding API"
3. Click on **Geocoding API**
4. Click **ENABLE**

---

## 3. Create a Service Account

### 3.1 Create the Service Account

1. Go to **APIs & Services** > **Credentials**
   - Or visit: https://console.cloud.google.com/apis/credentials
2. Click **+ CREATE CREDENTIALS** at the top
3. Select **Service account**
4. Fill in the details:
   - **Service account name**: `timeline-converter` (or any name you prefer)
   - **Service account ID**: Will auto-fill (e.g., `timeline-converter@your-project.iam.gserviceaccount.com`)
   - **Description**: "Service account for Timeline to Logbook converter" (optional)
5. Click **CREATE AND CONTINUE**
6. For "Grant this service account access to project":
   - Skip this (click **CONTINUE**)
7. For "Grant users access to this service account":
   - Skip this (click **DONE**)

### 3.2 Create and Download the JSON Key

1. You'll see your service account listed under **Service Accounts**
2. Click on the service account you just created
3. Go to the **KEYS** tab
4. Click **ADD KEY** > **Create new key**
5. Select **JSON** format
6. Click **CREATE**
7. A JSON file will automatically download to your computer
8. **IMPORTANT**: Rename this file to `credentials.json` and move it to your project directory
9. ⚠️ **Keep this file secure!** Never share it or commit it to public repositories

### 3.3 Copy Your Service Account Email

1. On the service account details page, you'll see the email address
2. It looks like: `timeline-converter@your-project-id.iam.gserviceaccount.com`
3. **Copy this email address** - you'll need it to share your Google Sheet

---

## 4. Get Your Google Maps API Key

### 4.1 Create an API Key

1. Go to **APIs & Services** > **Credentials**
2. Click **+ CREATE CREDENTIALS**
3. Select **API key**
4. A popup will show your new API key
5. Click **COPY** to copy the key
6. Click **RESTRICT KEY** (recommended for security)

### 4.2 Restrict the API Key (Recommended)

1. Under **API restrictions**:
   - Select **Restrict key**
   - Check only: **Geocoding API**
2. Under **Application restrictions** (optional but recommended):
   - Select **IP addresses**
   - Add your computer's IP address
3. Click **SAVE**

### 4.3 Enable Billing (Required for Geocoding API)

⚠️ The Geocoding API requires a billing account, but includes $200 free credit per month.

1. Go to **Billing** in the Google Cloud Console
2. Click **LINK A BILLING ACCOUNT** or **CREATE BILLING ACCOUNT**
3. Follow the steps to add a payment method
4. Note: You get $200 free credit/month, which covers ~40,000 geocoding requests

**Cost Estimate**:
- Geocoding API: $5 per 1,000 requests
- Typical usage: 100-500 unique locations = $0.50-$2.50
- Your $200 monthly credit easily covers this

---

## 5. Create and Share Your Google Sheet

### 5.1 Create a New Google Sheet

1. Go to [Google Sheets](https://sheets.google.com)
2. Click **+ Blank** to create a new spreadsheet
3. Name your spreadsheet (e.g., "Vehicle Logbook 2024")
4. Create three tabs (sheets) at the bottom:
   - **Timeline Data** - Where raw imported data goes
   - **Final Report** - Where the processed logbook will appear
   - **Backup** (optional) - For keeping a copy of original data

### 5.2 Get Your Google Sheet ID

The Sheet ID is in the URL when you open your sheet:

```
https://docs.google.com/spreadsheets/d/1AbC2DeF3GhI4JkL5MnO6PqR7StU8VwX9YzA/edit
                                      └─────────────────┬──────────────────┘
                                              This is your Sheet ID
```

**Copy this ID** - you'll need it for the script configuration.

### 5.3 Share the Sheet with Your Service Account

🔥 **CRITICAL STEP** - The script cannot access your sheet without this!

1. Click the **Share** button (top right of your Google Sheet)
2. In the "Add people and groups" field, paste your **service account email**:
   ```
   timeline-converter@your-project-id.iam.gserviceaccount.com
   ```
3. Change permission from "Viewer" to **Editor**
4. **Uncheck** "Notify people" (the service account won't receive emails)
5. Click **Share** or **Done**

✅ You should see the service account email listed with "Editor" access

---

## 6. Configure the Script

Open `timeline_complete_pipeline.py` (or `timeline_complete_public.py`) and update these values:

```python
# Google API Credentials
GOOGLE_MAPS_API_KEY = "AIzaSyAbc123YourActualAPIKeyHere"  # From step 4
SHEETS_CREDENTIALS_PATH = "credentials.json"  # Downloaded in step 3
SPREADSHEET_ID = "1AbC2DeF3GhI4JkL5MnO6PqR7StU8VwX9YzA"  # From step 5.2
SERVICE_ACCOUNT_EMAIL = "timeline-converter@your-project-id.iam.gserviceaccount.com"  # From step 3.3
```

---

## Verification Checklist

Before running the script, verify:

- [ ] Google Sheets API is enabled in your Google Cloud project
- [ ] Geocoding API is enabled in your Google Cloud project
- [ ] You have a `credentials.json` file in your project directory
- [ ] Your Google Sheet has been shared with the service account email as "Editor"
- [ ] You've copied your Google Maps API key into the script
- [ ] You've copied your Google Sheet ID into the script
- [ ] You have a billing account linked (for Geocoding API)

---

## Testing Your Setup

Run this simple test to verify everything works:

```bash
python test_connection.py
```

If the test script isn't available, try running the main script:

```bash
python timeline_complete_pipeline.py
```

### Expected Success Output:
```
✅ Credentials loaded successfully
✅ API service initialized
✅ Successfully connected to spreadsheet!
```

### Common Errors:

**"403 Forbidden"**
- ❌ You didn't share the Google Sheet with the service account email
- ✅ Solution: Follow step 5.3 again

**"API key not valid"**
- ❌ Your Google Maps API key is incorrect or restricted
- ✅ Solution: Check step 4 and verify the key

**"Geocoding API has not been used in project"**
- ❌ Geocoding API isn't enabled or billing isn't set up
- ✅ Solution: Follow step 2 and enable billing in step 4.3

**"Unable to parse range"**
- ❌ Sheet tab names don't match what the script expects
- ✅ Solution: Create tabs named "Timeline Data" and "Final Report"

---

## Security Best Practices

### Protecting Your Credentials

1. **Never commit `credentials.json` to git**
   - Add to `.gitignore`:
     ```
     credentials.json
     *.json
     ```

2. **Never share your API keys publicly**
   - Don't post them in forums, GitHub issues, etc.

3. **Restrict your API key**
   - Follow step 4.2 to limit what the key can do

4. **Revoke and regenerate if exposed**
   - If you accidentally expose credentials:
     - Delete the service account key in Google Cloud Console
     - Create a new key
     - Regenerate your API key

### Minimal Permissions

The service account only needs:
- Read/write access to your specific Google Sheet
- No other Google Drive or project permissions are needed

---

## Troubleshooting

### "Error loading credentials.json"

**Problem**: File not found or incorrectly formatted

**Solutions**:
1. Verify the file is named exactly `credentials.json`
2. Check it's in the same directory as your script
3. Open the file - it should be valid JSON starting with `{`
4. Re-download from Google Cloud Console if corrupted

### "Sheet not found" or "Unable to read sheet"

**Problem**: Script can't access your Google Sheet

**Solutions**:
1. Verify the Sheet ID is correct (copy from URL)
2. Confirm you shared the sheet with the service account
3. Check the service account has "Editor" (not "Viewer") permission
4. Make sure tab names match: "Timeline Data" and "Final Report"

### "Geocoding API error" or "OVER_QUERY_LIMIT"

**Problem**: API quota or billing issues

**Solutions**:
1. Check billing is enabled in Google Cloud Console
2. Verify Geocoding API is enabled
3. Check you haven't exceeded your quota (unlikely with $200 credit)
4. Wait a few minutes and try again

### "Invalid value at 'requests'"

**Problem**: Data format issue when writing to sheet

**Solutions**:
1. Ensure your Timeline Data has proper headers in row 1
2. Check dates are in format: YYYY-MM-DD HH:MM:SS
3. Verify no cells have special characters that break formatting

---

## Cost Management

### Monitoring Your Usage

1. Go to **Google Cloud Console** > **Billing** > **Reports**
2. Filter by:
   - Service: Geocoding API
   - Time range: Current month
3. You'll see your usage and costs

### Staying Within Free Tier

- **$200 free credit per month** = ~40,000 geocoding requests
- Typical usage: 100-5,000 requests per run
- The script uses caching to minimize repeat requests
- **You're very unlikely to exceed the free tier**

### Setting Up Budget Alerts (Optional)

1. Go to **Billing** > **Budgets & alerts**
2. Click **CREATE BUDGET**
3. Set budget amount (e.g., $10)
4. Set alert thresholds (e.g., 50%, 90%, 100%)
5. Add your email to receive alerts

---

## Additional Resources

- [Google Cloud Console](https://console.cloud.google.com/)
- [Google Sheets API Documentation](https://developers.google.com/sheets/api)
- [Geocoding API Documentation](https://developers.google.com/maps/documentation/geocoding)
- [Service Account Documentation](https://cloud.google.com/iam/docs/service-accounts)

---

## Summary

You now have:
1. ✅ A Google Cloud Project with required APIs enabled
2. ✅ A service account with JSON credentials
3. ✅ A Google Maps API key for geocoding
4. ✅ A Google Sheet shared with your service account
5. ✅ A properly configured script ready to run

**Next Step**: Run your script and start converting your Timeline data to a CRA-compliant logbook!

```bash
python timeline_complete_pipeline.py
```
