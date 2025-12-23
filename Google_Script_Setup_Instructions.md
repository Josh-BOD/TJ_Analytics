# TrafficJunky Google Sheets Integration - Setup Instructions

## Overview
This Google Apps Script pulls campaign data from the TrafficJunky API directly into your Google Spreadsheet, eliminating the need to run Python scripts manually.

## Features
- ✅ Pull last 30 days of campaign data with one click
- ✅ Quick access to Last 7 Days, This Week, and This Month
- ✅ Custom date range selection
- ✅ Automatic data formatting (currency, numbers)
- ✅ Auto-updating timestamp
- ✅ Custom menu in Google Sheets for easy access
- ✅ Preserves your custom formulas (columns Q onwards) when updating data
- ✅ **EST Timezone Support** - All dates use EST timezone to match TrafficJunky platform

## Setup Instructions (Step-by-Step)

### Step 1: Open Your Google Spreadsheet
1. Go to your spreadsheet: https://docs.google.com/spreadsheets/d/1UWXzuWnM2dBbtDSgbofOnwO33UMyOQg8w-8E9jQaMXU/edit
2. Make sure you're logged in to your Google account

### Step 2: Open the Script Editor
1. In your Google Spreadsheet, click on **Extensions** in the top menu
2. Select **Apps Script**
3. This will open a new tab with the Google Apps Script editor

### Step 3: Add the Script
1. Delete any existing code in the editor (usually says "function myFunction() {}")
2. Open the file `TrafficJunkyGoogleScript.gs` from this folder
3. Copy ALL the code from that file
4. Paste it into the Google Apps Script editor

### Step 4: Configure the Sheet Name (Optional)
1. In the script, find this line near the top:
   ```javascript
   const SHEET_NAME = "Campaign Data";
   ```
2. Change "Campaign Data" to match the name of the sheet tab where you want the data to appear
3. If you want to use a different sheet name, either:
   - Rename your existing sheet to "Campaign Data", OR
   - Change the SHEET_NAME value in the script

### Step 5: Save the Script
1. Click the **disk icon** (💾) or press `Ctrl+S` (Windows) / `Cmd+S` (Mac)
2. Give your project a name like "TrafficJunky Data Puller"
3. Click **Save**

### Step 6: Run the Script for the First Time
1. Still in the Apps Script editor, find the function dropdown (usually shows "pullTrafficJunkyData")
2. Make sure "pullTrafficJunkyData" is selected
3. Click the **Run** button (▶️)
4. You'll see a permission dialog - this is normal!

### Step 7: Grant Permissions
1. Click **Review Permissions**
2. Choose your Google account
3. Click **Advanced** (at the bottom left)
4. Click **Go to TrafficJunky Data Puller (unsafe)** - Don't worry, this is your own script!
5. Click **Allow**

### Step 8: Go Back to Your Spreadsheet
1. Return to your Google Spreadsheet tab
2. Refresh the page (F5 or reload button)
3. You should now see a new menu called **"TrafficJunky"** in the top menu bar

## How to Use

The script provides two types of data:
1. **Aggregated Data** - Total performance for entire date range (for overview reporting)
2. **Daily Breakdown** - Day-by-day performance data (for trend analysis and pivot tables)

### Aggregated Data (Campaign Summary)

**Last 7 Days:**
1. Click **TrafficJunky** menu → **📊 Aggregated Data** → **📅 Last 7 Days**
2. Data will automatically pull for the past 7 days (ending yesterday)

**This Week:**
1. Click **TrafficJunky** menu → **📊 Aggregated Data** → **📆 This Week**
2. Data will pull from Monday of the current week to yesterday

**This Month:**
1. Click **TrafficJunky** menu → **📊 Aggregated Data** → **📊 This Month**
2. Data will pull from the 1st of the current month to yesterday

**Last 30 Days:**
1. Click **TrafficJunky** menu → **📊 Aggregated Data** → **📈 Last 30 Days**
2. Data will pull for the past 30 days (ending yesterday)

**Custom Date Range:**
1. Click **TrafficJunky** menu → **📊 Aggregated Data** → **🔧 Custom Date Range**
2. Enter start and end dates
3. Data will be fetched and displayed in "Campaign Data" sheet

### Daily Breakdown (For Pivot Tables & Trend Analysis)

The daily breakdown pulls data day-by-day and stores it in the **"RAW_DailyData-DNT"** sheet. This is perfect for:
- Creating pivot tables to analyze trends over time
- Tracking individual campaign performance by date
- Building time-series charts and dashboards
- Comparing day-over-day performance

**Important:** Daily breakdown only updates the selected date range (e.g., last 14 days), preserving all historical data beyond that range. This means you can refresh recent data without re-pulling everything!

**Update Last 7 Days:**
1. Click **TrafficJunky** menu → **📅 Daily Breakdown** → **Update Last 7 Days**
2. Updates/replaces data for the last 7 days only
3. All historical data older than 7 days is preserved

**Update Last 14 Days:** (Recommended for regular updates)
1. Click **TrafficJunky** menu → **📅 Daily Breakdown** → **Update Last 14 Days**
2. Updates/replaces data for the last 14 days only
3. Perfect for daily or weekly refresh routines

**Update This Month:**
1. Click **TrafficJunky** menu → **📅 Daily Breakdown** → **Update This Month**
2. Updates/replaces data from the 1st of the month to yesterday
3. Great for monthly reporting

**Custom Date Range:**
1. Click **TrafficJunky** menu → **📅 Daily Breakdown** → **Custom Date Range**
2. Enter start and end dates
3. Only the specified date range is updated/replaced

**Clear Daily Data:**
1. Click **TrafficJunky** menu → **📅 Daily Breakdown** → **Clear Daily Data**
2. Removes all daily breakdown data (use this to start fresh)

### Pull Latest Data (Last 30 Days)
1. Click **TrafficJunky** menu → **Pull Latest Data (Last 30 Days)**
2. Wait for the "Fetching data..." message
3. Data will appear in your sheet with all campaigns and statistics

### Pull Custom Date Range
1. Click **TrafficJunky** menu → **Pull Custom Date Range**
2. Enter start date in format: `YYYY-MM-DD` (e.g., 2024-10-01)
3. Click OK
4. Enter end date in format: `YYYY-MM-DD` (e.g., 2024-11-13)
5. Click OK
6. Data will be fetched and displayed

**Note:** If you enter today's date or a future date as the end date, the script will automatically adjust it to yesterday, since TrafficJunky API only provides data that's at least 1 day old.

### Clear Data
1. Click **TrafficJunky** menu → **🗑️ Clear Data**
2. Confirm you want to clear the sheet
3. Only columns A-P (API data) will be cleared - your formulas in column Q onwards are preserved

## Date Range Examples

Here's what each date range pulls (assuming today is November 14, 2025):

| Option | Start Date | End Date | Use Case |
|--------|-----------|----------|----------|
| **Last 7 Days** | Nov 6, 2025 | Nov 13, 2025 | Quick weekly performance check |
| **This Week** | Nov 10, 2025 (Mon) | Nov 13, 2025 | Current week's performance |
| **This Month** | Nov 1, 2025 | Nov 13, 2025 | Month-to-date performance |
| **Last 30 Days** | Oct 14, 2025 | Nov 13, 2025 | Monthly trends and analysis |
| **Custom** | Your choice | Your choice | Historical data or specific periods |

## Daily Breakdown Data Structure

The **"RAW_DailyData-DNT"** sheet contains the following columns:

| Column | Description |
|--------|-------------|
| Date | Date of the data (YYYY-MM-DD format) |
| Campaign ID | Unique identifier for each campaign |
| Campaign Name | Name of the campaign |
| Campaign Type | Type of campaign |
| Status | Current status |
| Impressions | Daily impressions |
| Clicks | Daily clicks |
| Conversions | Daily conversions |
| Cost | Daily cost |
| CTR | Click-through rate |
| CPM | Cost per thousand impressions |
| Last Updated | Timestamp of when data was pulled |

### Creating Pivot Tables with Daily Data

**Example 1: Campaign Performance Over Time**
1. Select all data in "RAW_DailyData-DNT" sheet
2. Go to **Insert → Pivot table**
3. Setup:
   - Rows: Campaign Name
   - Columns: Date
   - Values: Cost (SUM), Clicks (SUM), Conversions (SUM)
4. This shows each campaign's daily performance in a matrix

**Example 2: Daily Trend Analysis**
1. Create pivot table from "RAW_DailyData-DNT"
2. Setup:
   - Rows: Date
   - Values: Cost (SUM), Clicks (SUM), Conversions (SUM), Impressions (SUM)
3. Insert chart to visualize trends over time

**Example 3: Top Performing Days**
1. Create pivot table from "RAW_DailyData-DNT"
2. Setup:
   - Rows: Date
   - Values: Conversions (SUM), Cost (SUM)
3. Sort by Conversions descending

**Example 4: Campaign Comparison by Week**
1. Create pivot table from "RAW_DailyData-DNT"
2. Setup:
   - Rows: Campaign Name
   - Columns: Date (grouped by week)
   - Values: Cost (SUM), ROAS (calculated field)
3. Compare week-over-week performance

### How Daily Updates Work

When you run "Update Last 14 Days":
1. Script fetches data for each of the last 14 days
2. Removes any existing rows for those 14 days
3. Inserts the fresh data for those 14 days
4. **Preserves all data older than 14 days**

**Example Timeline:**
- You have historical data from Jan 1 - Nov 13
- You run "Update Last 7 Days" on Nov 14
- Result: Jan 1 - Nov 6 (unchanged) + Nov 7 - Nov 13 (refreshed)

This approach means:
✅ You can update recent data daily without re-pulling months of history  
✅ Historical data is preserved and never lost  
✅ Much faster than pulling all data every time  
✅ Reduces API calls and execution time

## Data Columns Explained

| Column | Description |
|--------|-------------|
| Campaign ID | Unique identifier for each campaign |
| Campaign Name | Name of the campaign |
| Campaign Type | Type of campaign (e.g., display, native) |
| Status | Current status (active, paused, etc.) |
| Daily Budget | Daily budget allocated to campaign |
| Daily Budget Left | Remaining daily budget |
| Ads Paused | Number of paused ads |
| Number of Bids | Total number of bids |
| Number of Creatives | Total number of creatives |
| Impressions | Total impressions |
| Clicks | Total clicks |
| Conversions | Total conversions |
| Cost | Total cost spent |
| CTR | Click-through rate (displayed as decimal, e.g., 7.99 = 7.99%) |
| CPM | Cost per thousand impressions |
| Last Updated | Timestamp of when data was pulled |

## Adding Custom Calculated Columns (CPA, ROAS, etc.)

The script only updates columns **A through P** (the API data). You can safely add your own formulas starting in **column Q** onwards, and they will be preserved when you refresh the data!

### Example: Add CPA (Cost Per Acquisition) in Column Q

**Step 1:** In cell Q1, add the header:
```
CPA
```

**Step 2:** In cell Q2, add the formula:
```
=IF(L2>0, M2/L2, 0)
```
This calculates: Cost ÷ Conversions

**Step 3:** Copy the formula down for all rows (drag the cell corner or double-click the fill handle)

### More Calculated Metrics You Can Add:

**Column R - ROAS (Return on Ad Spend):**
- Header: `ROAS`
- Formula in R2: `=IF(M2>0, [REVENUE]/M2, 0)` (replace [REVENUE] with your revenue column)

**Column S - CVR (Conversion Rate):**
- Header: `CVR`
- Formula in S2: `=IF(K2>0, L2/K2*100, 0)`
- This calculates: (Conversions ÷ Clicks) × 100

**Column T - CPC (Cost Per Click):**
- Header: `CPC`
- Formula in T2: `=IF(K2>0, M2/K2, 0)`
- This calculates: Cost ÷ Clicks

**Column U - CPM Check:**
- Header: `CPM Calculated`
- Formula in U2: `=IF(J2>0, M2/J2*1000, 0)`
- This calculates: (Cost ÷ Impressions) × 1000

### Important Notes:

✅ **Your formulas are safe** - They won't be deleted when you pull new data  
✅ **Start at column Q** - Columns A-P are reserved for API data  
✅ **Copy formulas down** - When new campaigns are added, drag your formulas to match the new rows  
✅ **Use relative references** - Use `M2` instead of `$M$2` so the formula adjusts for each row

## Setting Up Automatic Daily Updates (Optional)

If you want the data to update automatically every day:

1. In the Apps Script editor, click the **clock icon** (⏰) on the left sidebar - "Triggers"
2. Click **+ Add Trigger** (bottom right)
3. Configure:
   - Choose function: `pullTrafficJunkyData`
   - Choose event source: `Time-driven`
   - Choose type: `Day timer`
   - Choose time of day: Pick a time (e.g., "6am to 7am")
4. Click **Save**

Now your data will automatically update every day at the specified time!

## Troubleshooting

### Cost data doesn't match TrafficJunky platform
- **Fixed!** The script now uses EST timezone to match the TrafficJunky platform
- Make sure you've updated to the latest version of the script
- The script automatically calculates "yesterday" and date ranges based on EST time
- If you're in a different timezone (PST, CST, etc.), the script will still use EST dates

### "API returned status code 400" or "end date must be before tomorrow" error
- **This has been fixed!** The script now automatically adjusts any future dates to yesterday (in EST)
- If you still see this error, make sure you've updated to the latest version of the script
- The TrafficJunky API only provides data that's at least 1 day old (yesterday or earlier)

### "Sheet not found" error
- Make sure the SHEET_NAME in the script matches your sheet tab name exactly
- Check for typos and ensure capitalization matches

### No data appears
- Check the Logs: In Apps Script editor, click **Execution log** at the bottom
- Verify you have campaigns with data in the selected date range

### Permission errors
- You may need to re-authorize the script if you see permission errors
- Go back to Step 6-7 and re-run the authorization process

## Differences from Python Script

| Feature | Python Script | Google Script |
|---------|--------------|---------------|
| **Where it runs** | Your computer | Google's servers |
| **Data storage** | CSV file | Google Spreadsheet |
| **Scheduling** | Manual or cron job | Google Triggers (built-in) |
| **Formatting** | Basic CSV | Auto-formatted with colors |
| **Sharing** | Email CSV files | Share Google Sheet link |

## Next Steps

Once you have the data in Google Sheets, you can:
- Create pivot tables for analysis
- Build charts and dashboards
- Share the sheet with team members
- Export to other formats if needed
- Use Google Data Studio for advanced visualizations

## Support

If you run into issues:
1. Check the Execution log in Apps Script editor
2. Verify your API key is correct
3. Ensure date formats are correct
4. Check that the TrafficJunky API is accessible

