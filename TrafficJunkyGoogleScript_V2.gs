/**
 * TrafficJunky API Data Extractor for Google Sheets
 * This script pulls campaign data from TrafficJunky API and populates it into Google Sheets
 */

// Configuration
const API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039";
const API_URL = "https://api.trafficjunky.com/api/campaigns/bids/stats.json";
const SHEET_NAME = "RAW Data - DNT"; // Aggregated campaign data sheet
const DAILY_SHEET_NAME = "RAW_DailyData-DNT"; // Sheet for daily breakdown data
const API_TIMEZONE = "America/New_York"; // TrafficJunky API uses EST/EDT

// PostHog Configuration
const POSTHOG_API_KEY = "phx_eSPeTXSJRa8cXGXW7qqzQFr51uWjnGHpT6yRxlgvAWThQAv";  // Get from PostHog Settings > Personal API Keys
const POSTHOG_PROJECT_ID = "107054";             // From URL: us.posthog.com/project/{PROJECT_ID}/...
const POSTHOG_HOST = "https://us.posthog.com";
const CONVERSIONS_SHEET_NAME = "PostHog_Conversions";

/**
 * Helper function to get current date/time in EST timezone
 */
function getESTDate() {
  return new Date(new Date().toLocaleString("en-US", {timeZone: API_TIMEZONE}));
}

/**
 * Helper function to get yesterday in EST timezone
 */
function getESTYesterday() {
  const estNow = getESTDate();
  estNow.setDate(estNow.getDate() - 1);
  estNow.setHours(0, 0, 0, 0);
  return estNow;
}

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('TrafficJunky')
    .addSubMenu(ui.createMenu('📊 Aggregated Data')
      .addItem('📅 Last 7 Days', 'pullLast7Days')
      .addItem('📆 This Week', 'pullThisWeek')
      .addItem('📊 This Month', 'pullThisMonth')
      .addItem('📈 Last 30 Days', 'pullTrafficJunkyData')
      .addSeparator()
      .addItem('🔧 Custom Date Range', 'pullCustomDateRange')
      .addSeparator()
      .addItem('🗑️ Clear Aggregated Data', 'clearData'))
    .addSubMenu(ui.createMenu('📅 Daily Breakdown')
      .addItem('Update Last 7 Days', 'updateDailyLast7Days')
      .addItem('Update Last 14 Days', 'updateDailyLast14Days')
      .addItem('Update This Month', 'updateDailyThisMonth')
      .addSeparator()
      .addItem('Custom Date Range', 'updateDailyCustomRange')
      .addSeparator()
      .addItem('🗑️ Clear Daily Data', 'clearDailyData'))
    .addSubMenu(ui.createMenu('🔄 PostHog Conversions')
      .addItem('📅 Today (EST)', 'pullTodayConversions')
      .addItem('📆 Yesterday (EST)', 'pullYesterdayConversions')
      .addItem('📊 Last 7 Days', 'pullConversionsLast7Days')
      .addItem('📈 Last 30 Days', 'pullConversionsLast30Days')
      .addSeparator()
      .addItem('🔧 Custom Date Range', 'pullConversionsCustomRange')
      .addSeparator()
      .addItem('🗑️ Clear Conversion Data', 'clearConversionData'))
    .addSeparator()
    .addItem('🔍 Show API Data Structure', 'showAPIDataStructure')
    .addToUi();
}

/**
 * Main function to pull TrafficJunky data (Last 30 days)
 */
function pullTrafficJunkyData() {
  try {
    // Calculate dates in EST timezone - end date must be yesterday
    const endDate = getESTYesterday();
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 29); // 29 days before yesterday = 30 days total
    
    fetchAndWriteData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullTrafficJunkyData: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull data for last 7 days
 */
function pullLast7Days() {
  try {
    const endDate = getESTYesterday();
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 6); // 6 days before yesterday = 7 days total
    
    fetchAndWriteData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullLast7Days: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull data for this week (Monday to yesterday)
 */
function pullThisWeek() {
  try {
    const endDate = getESTYesterday();
    
    // Get Monday of current week
    const startDate = new Date(endDate);
    const day = startDate.getDay();
    const diff = startDate.getDate() - day + (day === 0 ? -6 : 1); // Adjust when day is Sunday
    startDate.setDate(diff);
    
    fetchAndWriteData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullThisWeek: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull data for this month (1st of month to yesterday)
 */
function pullThisMonth() {
  try {
    const endDate = getESTYesterday();
    
    // Get first day of current month in EST
    const startDate = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
    
    fetchAndWriteData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullThisMonth: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Function to pull data with custom date range
 */
function pullCustomDateRange() {
  const ui = SpreadsheetApp.getUi();
  
  // Prompt for start date
  const startDateResponse = ui.prompt(
    'Start Date',
    'Enter start date (YYYY-MM-DD):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (startDateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  // Prompt for end date
  const endDateResponse = ui.prompt(
    'End Date',
    'Enter end date (YYYY-MM-DD):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (endDateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  try {
    let startDate = new Date(startDateResponse.getResponseText());
    let endDate = new Date(endDateResponse.getResponseText());
    
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      ui.alert('Invalid date format. Please use YYYY-MM-DD format.');
      return;
    }
    
    // Validate and adjust dates
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);
    
    // If end date is today or in the future, adjust to yesterday
    if (endDate >= today) {
      endDate = yesterday;
      ui.alert('Note: End date adjusted to yesterday (' + formatDateForDisplay(yesterday) + ') as TrafficJunky API requires data to be at least 1 day old.');
    }
    
    // If start date is after end date, show error
    if (startDate > endDate) {
      ui.alert('Error: Start date cannot be after end date. Start: ' + formatDateForDisplay(startDate) + ', End: ' + formatDateForDisplay(endDate));
      return;
    }
    
    fetchAndWriteData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullCustomDateRange: " + error.toString());
    ui.alert('Error: ' + error.toString());
  }
}

/**
 * Fetches data from TrafficJunky API and writes to sheet
 */
function fetchAndWriteData(startDate, endDate) {
  const ui = SpreadsheetApp.getUi();
  
  // Validate and adjust dates to ensure API compliance (using EST timezone)
  const estToday = getESTDate();
  estToday.setHours(0, 0, 0, 0);
  const estYesterday = getESTYesterday();
  
  // Auto-adjust end date if it's today or in the future (in EST)
  // Compare by date only, not time
  if (endDate.getTime() >= estToday.getTime()) {
    endDate = estYesterday;
    Logger.log(`End date auto-adjusted to EST yesterday: ${formatDateForDisplay(endDate)}`);
  }
  
  // Show loading message
  ui.alert('Fetching data from TrafficJunky API (EST timezone)...');
  
  // Format dates as DD/MM/YYYY (TrafficJunky API format)
  const formattedStartDate = formatDate(startDate);
  const formattedEndDate = formatDate(endDate);
  
  Logger.log(`Fetching data from ${formattedStartDate} to ${formattedEndDate} (EST timezone)`);
  
  // Build API URL with parameters
  const url = `${API_URL}?api_key=${API_KEY}&startDate=${formattedStartDate}&endDate=${formattedEndDate}&limit=1000&offset=1`;
  
  try {
    // Make API request
    const response = UrlFetchApp.fetch(url, {
      'method': 'get',
      'contentType': 'application/json',
      'muteHttpExceptions': true
    });
    
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      throw new Error(`API returned status code ${responseCode}: ${response.getContentText()}`);
    }
    
    const jsonData = JSON.parse(response.getContentText());
    Logger.log(`Received data for ${Object.keys(jsonData).length} campaigns`);
    
    // Process and write data
    writeDataToSheet(jsonData, formattedStartDate, formattedEndDate);
    
    ui.alert('Success!', 'Data has been successfully imported from TrafficJunky API.', ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log("Error fetching data: " + error.toString());
    ui.alert('Error fetching data: ' + error.toString());
  }
}

/**
 * Processes API response and writes to Google Sheet
 */
function writeDataToSheet(apiData, startDate, endDate) {
  Logger.log(`writeDataToSheet called with startDate: ${startDate}, endDate: ${endDate}`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  Logger.log(`Sheet name to use: ${SHEET_NAME}`);
  
  // Create sheet if it doesn't exist
  if (!sheet) {
    Logger.log(`Sheet "${SHEET_NAME}" not found, creating new sheet`);
    sheet = ss.insertSheet(SHEET_NAME);
  } else {
    Logger.log(`Sheet "${SHEET_NAME}" found`);
  }
  
  // Define headers
  const headers = [
    'Campaign ID',
    'Campaign Name',
    'Campaign Type',
    'Status',
    'Daily Budget',
    'Daily Budget Left',
    'Ads Paused',
    'Number of Bids',
    'Number of Creatives',
    'Impressions',
    'Clicks',
    'Conversions',
    'Cost',
    'CTR',
    'CPM',
    'Last Updated'
  ];
  
  // Clear only the data columns (A-P), preserving any formulas in Q onwards
  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    // Clear columns A through P (16 columns) only
    sheet.getRange(1, 1, Math.max(lastRow, 1000), headers.length).clearContent();
  }
  
  // Write headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  
  // Process campaigns
  const rows = [];
  let campaigns = [];
  
  // Handle both list and dict responses from API
  if (Array.isArray(apiData)) {
    campaigns = apiData;
    Logger.log(`API returned array with ${campaigns.length} campaigns`);
  } else if (typeof apiData === 'object' && apiData !== null) {
    campaigns = Object.values(apiData);
    Logger.log(`API returned object with ${campaigns.length} campaigns`);
  } else {
    Logger.log(`ERROR: Unexpected API response format: ${typeof apiData}`);
    Logger.log(`API Data: ${JSON.stringify(apiData)}`);
    return;
  }
  
  Logger.log(`Processing ${campaigns.length} campaigns...`);
  
  // Convert campaigns to rows
  for (let campaign of campaigns) {
    if (campaign && typeof campaign === 'object') {
      const row = [
        campaign.campaignId || campaign.id || 'unknown',
        campaign.campaignName || '',
        campaign.campaignType || '',
        campaign.status || '',
        toNumeric(campaign.dailyBudget, 0),
        toNumeric(campaign.dailyBudgetLeft, 0),
        toNumeric(campaign.adsPaused, 0),
        toNumeric(campaign.numberOfBids, 0),
        toNumeric(campaign.numberOfCreative, 0),
        toNumeric(campaign.impressions, 0),
        toNumeric(campaign.clicks, 0),
        toNumeric(campaign.conversions, 0),
        toNumeric(campaign.cost, 0),
        toNumeric(campaign.CTR, 0),
        toNumeric(campaign.CPM, 0),
        new Date()
      ];
      rows.push(row);
    }
  }
  
  Logger.log(`Created ${rows.length} rows to write`);
  
  // Write data to sheet
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    
    // Format numeric columns
    const lastRow = rows.length + 1;
    
    // Format currency columns (Daily Budget, Daily Budget Left, Cost, CPM)
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat('$#,##0.00'); // Daily Budget
    sheet.getRange(2, 6, rows.length, 1).setNumberFormat('$#,##0.00'); // Daily Budget Left
    sheet.getRange(2, 13, rows.length, 1).setNumberFormat('$#,##0.00'); // Cost
    sheet.getRange(2, 15, rows.length, 1).setNumberFormat('$#,##0.00'); // CPM
    
    // Format CTR column (API returns as percentage value, so just show 2 decimals)
    sheet.getRange(2, 14, rows.length, 1).setNumberFormat('0.00');
    
    // Format number columns (Impressions, Clicks, Conversions, etc.)
    sheet.getRange(2, 7, rows.length, 1).setNumberFormat('#,##0'); // Ads Paused
    sheet.getRange(2, 8, rows.length, 1).setNumberFormat('#,##0'); // Number of Bids
    sheet.getRange(2, 9, rows.length, 1).setNumberFormat('#,##0'); // Number of Creatives
    sheet.getRange(2, 10, rows.length, 1).setNumberFormat('#,##0'); // Impressions
    sheet.getRange(2, 11, rows.length, 1).setNumberFormat('#,##0'); // Clicks
    sheet.getRange(2, 12, rows.length, 1).setNumberFormat('#,##0'); // Conversions
    
    // Format timestamp column
    sheet.getRange(2, 16, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
    
    Logger.log(`Successfully wrote ${rows.length} campaigns to sheet`);
  } else {
    sheet.getRange(2, 1).setValue('No data found for the selected date range');
    Logger.log("No campaigns found in API response");
  }
}

/**
 * Helper function to convert values to numeric safely
 */
function toNumeric(value, defaultValue = 0) {
  if (value === null || value === undefined || value === '') {
    return defaultValue;
  }
  
  const num = Number(value);
  return isNaN(num) ? defaultValue : num;
}

/**
 * Helper function to format date as DD/MM/YYYY (for API)
 */
function formatDate(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

/**
 * Helper function to format date as YYYY-MM-DD (for display)
 */
function formatDateForDisplay(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${year}-${month}-${day}`;
}

/**
 * Clears API data from columns A-P only (preserves formulas in Q onwards)
 */
function clearData() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Clear Data',
    'Are you sure you want to clear API data from columns A-P? (Formulas in column Q onwards will be preserved)',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 0) {
        // Clear only columns A through P (16 columns)
        sheet.getRange(1, 1, lastRow, 16).clearContent();
      }
      ui.alert('API data cleared successfully. Formulas in column Q onwards are preserved.');
    }
  }
}

// ============================================================================
// DAILY BREAKDOWN FUNCTIONS
// ============================================================================

/**
 * Update daily data for last 7 days
 */
function updateDailyLast7Days() {
  try {
    const endDate = getESTYesterday();
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 6); // 6 days before yesterday = 7 days total
    
    updateDailyData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in updateDailyLast7Days: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Update daily data for last 14 days
 */
function updateDailyLast14Days() {
  try {
    const endDate = getESTYesterday();
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 13); // 13 days before yesterday = 14 days total
    
    updateDailyData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in updateDailyLast14Days: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Update daily data for this month
 */
function updateDailyThisMonth() {
  try {
    const endDate = getESTYesterday();
    
    // Get first day of current month in EST
    const startDate = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
    
    updateDailyData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in updateDailyThisMonth: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Update daily data for custom date range
 */
function updateDailyCustomRange() {
  const ui = SpreadsheetApp.getUi();
  
  // Prompt for start date
  const startDateResponse = ui.prompt(
    'Start Date',
    'Enter start date (YYYY-MM-DD):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (startDateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  // Prompt for end date
  const endDateResponse = ui.prompt(
    'End Date',
    'Enter end date (YYYY-MM-DD):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (endDateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  try {
    let startDate = new Date(startDateResponse.getResponseText());
    let endDate = new Date(endDateResponse.getResponseText());
    
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      ui.alert('Invalid date format. Please use YYYY-MM-DD format.');
      return;
    }
    
    // Validate and adjust dates
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);
    
    // If end date is today or in the future, adjust to yesterday
    if (endDate >= today) {
      endDate = yesterday;
      ui.alert('Note: End date adjusted to yesterday (' + formatDateForDisplay(yesterday) + ') as TrafficJunky API requires data to be at least 1 day old.');
    }
    
    // If start date is after end date, show error
    if (startDate > endDate) {
      ui.alert('Error: Start date cannot be after end date. Start: ' + formatDateForDisplay(startDate) + ', End: ' + formatDateForDisplay(endDate));
      return;
    }
    
    updateDailyData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in updateDailyCustomRange: " + error.toString());
    ui.alert('Error: ' + error.toString());
  }
}

/**
 * Main function to fetch and update daily breakdown data
 * This function only updates data for the specified date range, preserving historical data
 */
function updateDailyData(startDate, endDate) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Validate and adjust dates (using EST timezone)
  const estToday = getESTDate();
  estToday.setHours(0, 0, 0, 0);
  const estYesterday = getESTYesterday();
  
  if (endDate.getTime() >= estToday.getTime()) {
    endDate = estYesterday;
    Logger.log(`End date auto-adjusted to EST yesterday: ${formatDateForDisplay(endDate)}`);
  }
  
  ui.alert('Fetching daily breakdown data from TrafficJunky API (EST timezone)...');
  
  // Get or create daily sheet
  let sheet = ss.getSheetByName(DAILY_SHEET_NAME);
  const isNewSheet = !sheet;
  
  if (!sheet) {
    sheet = ss.insertSheet(DAILY_SHEET_NAME);
  }
  
  // Define headers
  const headers = [
    'Date',
    'Campaign ID',
    'Campaign Name',
    'Campaign Type',
    'Status',
    'Impressions',
    'Clicks',
    'Conversions',
    'Cost',
    'CTR',
    'CPM',
    'Last Updated'
  ];
  
  // Always ensure headers are present
  if (isNewSheet || sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() !== 'Date') {
    // Write headers
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    sheet.setFrozenRows(1);
    Logger.log('Headers written to daily sheet');
  }
  
  // Fetch data day by day to get daily breakdown
  const dailyData = [];
  const currentDate = new Date(startDate);
  
  while (currentDate <= endDate) {
    const dateStr = formatDate(currentDate);
    Logger.log(`Fetching data for ${dateStr}`);
    
    try {
      const url = `${API_URL}?api_key=${API_KEY}&startDate=${dateStr}&endDate=${dateStr}&limit=1000&offset=1`;
      
      const response = UrlFetchApp.fetch(url, {
        'method': 'get',
        'contentType': 'application/json',
        'muteHttpExceptions': true
      });
      
      const responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
        const jsonData = JSON.parse(response.getContentText());
        
        let campaigns = [];
        if (Array.isArray(jsonData)) {
          campaigns = jsonData;
        } else if (typeof jsonData === 'object') {
          campaigns = Object.values(jsonData);
        }
        
        // Process each campaign for this date
        for (let campaign of campaigns) {
          if (campaign && typeof campaign === 'object') {
            const row = [
              formatDateForDisplay(currentDate),
              campaign.campaignId || campaign.id || 'unknown',
              campaign.campaignName || '',
              campaign.campaignType || '',
              campaign.status || '',
              toNumeric(campaign.impressions, 0),
              toNumeric(campaign.clicks, 0),
              toNumeric(campaign.conversions, 0),
              toNumeric(campaign.cost, 0),
              toNumeric(campaign.CTR, 0),
              toNumeric(campaign.CPM, 0),
              new Date()
            ];
            dailyData.push(row);
          }
        }
      }
    } catch (error) {
      Logger.log(`Error fetching data for ${dateStr}: ${error.toString()}`);
    }
    
    // Move to next day
    currentDate.setDate(currentDate.getDate() + 1);
    Utilities.sleep(100); // Small delay to avoid rate limiting
  }
  
  if (dailyData.length > 0) {
    // Remove existing data for the date range being updated
    removeDataForDateRange(sheet, startDate, endDate);
    
    // Append new data
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, dailyData.length, 12).setValues(dailyData);
    
    // Format columns
    formatDailySheet(sheet, lastRow + 1, dailyData.length);
    
    // Sort by date descending, then by campaign name
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12);
    dataRange.sort([{column: 1, ascending: false}, {column: 3, ascending: true}]);
    
    ui.alert('Success!', `Updated ${dailyData.length} daily records from ${formatDateForDisplay(startDate)} to ${formatDateForDisplay(endDate)}.`, ui.ButtonSet.OK);
    Logger.log(`Successfully updated ${dailyData.length} daily records`);
  } else {
    ui.alert('No data found for the selected date range.');
    Logger.log("No data found in API response");
  }
}

/**
 * Remove existing data for a specific date range from the daily sheet
 * Only clears columns A-L to preserve formulas in column M onwards
 */
function removeDataForDateRange(sheet, startDate, endDate) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // Only header row
  
  const startDateStr = formatDateForDisplay(startDate);
  const endDateStr = formatDateForDisplay(endDate);
  
  // Get all dates in column A
  const dateRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const dates = dateRange.getValues();
  
  // Find rows to clear (only columns A-L, preserving formulas in M onwards)
  let rowsCleared = 0;
  for (let i = 0; i < dates.length; i++) {
    const rowDate = dates[i][0];
    if (rowDate >= startDateStr && rowDate <= endDateStr) {
      const actualRow = i + 2; // +2 because: array is 0-indexed, and we start from row 2
      // Clear only columns A through L (12 columns)
      sheet.getRange(actualRow, 1, 1, 12).clearContent();
      rowsCleared++;
    }
  }
  
  Logger.log(`Cleared ${rowsCleared} rows (columns A-L only) for date range ${startDateStr} to ${endDateStr}`);
}

/**
 * Format the daily data sheet
 */
function formatDailySheet(sheet, startRow, numRows) {
  if (numRows === 0) return;
  
  // Format currency columns (Cost, CPM)
  sheet.getRange(startRow, 9, numRows, 1).setNumberFormat('$#,##0.00'); // Cost
  sheet.getRange(startRow, 11, numRows, 1).setNumberFormat('$#,##0.00'); // CPM
  
  // Format CTR column (as decimal)
  sheet.getRange(startRow, 10, numRows, 1).setNumberFormat('0.00');
  
  // Format number columns
  sheet.getRange(startRow, 6, numRows, 1).setNumberFormat('#,##0'); // Impressions
  sheet.getRange(startRow, 7, numRows, 1).setNumberFormat('#,##0'); // Clicks
  sheet.getRange(startRow, 8, numRows, 1).setNumberFormat('#,##0'); // Conversions
  
  // Format date column
  sheet.getRange(startRow, 1, numRows, 1).setNumberFormat('yyyy-mm-dd');
  
  // Format timestamp column
  sheet.getRange(startRow, 12, numRows, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

/**
 * Clear all daily data (only columns A-L, preserving formulas in M onwards)
 */
function clearDailyData() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Clear Daily Data',
    'Are you sure you want to clear all daily breakdown data (columns A-L)? Formulas in column M onwards will be preserved.',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(DAILY_SHEET_NAME);
    
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        // Clear only columns A through L (12 columns), preserving formulas in M onwards
        sheet.getRange(2, 1, lastRow - 1, 12).clearContent();
      }
      ui.alert('Daily data cleared successfully (columns A-L). Formulas in column M onwards are preserved.');
    } else {
      ui.alert('Daily data sheet not found.');
    }
  }
}

// ============================================================================
// API DATA STRUCTURE VIEWER
// ============================================================================

/**
 * Shows the raw API data structure in a new sheet
 */
function showAPIDataStructure() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  ui.alert('Fetching sample data from API to show structure...');
  
  try {
    // Get sample data (last 3 days to get quick response)
    const endDate = getESTYesterday();
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 2); // Just 3 days
    
    const formattedStartDate = formatDate(startDate);
    const formattedEndDate = formatDate(endDate);
    
    const url = `${API_URL}?api_key=${API_KEY}&startDate=${formattedStartDate}&endDate=${formattedEndDate}&limit=5&offset=1`;
    
    const response = UrlFetchApp.fetch(url, {
      'method': 'get',
      'contentType': 'application/json',
      'muteHttpExceptions': true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`API error: ${response.getContentText()}`);
    }
    
    const jsonData = JSON.parse(response.getContentText());
    
    // Get or create sheet
    let sheet = ss.getSheetByName('API_Data_Structure');
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    sheet = ss.insertSheet('API_Data_Structure');
    
    // Get first campaign
    const firstCampaign = Array.isArray(jsonData) ? jsonData[0] : Object.values(jsonData)[0];
    
    if (!firstCampaign) {
      ui.alert('No data returned from API');
      return;
    }
    
    // Write overview
    sheet.getRange('A1').setValue('TrafficJunky API Response Structure');
    sheet.getRange('A1').setFontSize(14).setFontWeight('bold');
    
    sheet.getRange('A2').setValue(`Date Range: ${formattedStartDate} to ${formattedEndDate}`);
    sheet.getRange('A3').setValue(`Total Campaigns Returned: ${Object.keys(jsonData).length}`);
    
    // Field structure table
    sheet.getRange('A5').setValue('FIELD NAME');
    sheet.getRange('B5').setValue('SAMPLE VALUE');
    sheet.getRange('C5').setValue('DATA TYPE');
    sheet.getRange('D5').setValue('DESCRIPTION');
    sheet.getRange('A5:D5').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    
    let row = 6;
    
    // Main campaign fields
    for (let key in firstCampaign) {
      if (key === 'bids' || key === 'spots') continue; // Handle separately
      
      const value = firstCampaign[key];
      const type = typeof value;
      
      sheet.getRange(row, 1).setValue(key);
      sheet.getRange(row, 2).setValue(String(value).substring(0, 100)); // Limit length
      sheet.getRange(row, 3).setValue(type);
      
      // Add descriptions
      const descriptions = {
        'campaignId': 'Unique campaign identifier',
        'campaignName': 'Campaign name',
        'campaignType': 'Type of campaign (bid, deal, etc.)',
        'status': 'Campaign status (active, paused, etc.)',
        'dailyBudget': 'Daily budget limit',
        'dailyBudgetLeft': 'Remaining daily budget',
        'clicks': 'Total clicks',
        'impressions': 'Total impressions',
        'conversions': 'Total conversions',
        'cost': 'Total cost spent',
        'CTR': 'Click-through rate (%)',
        'CPM': 'Cost per thousand impressions',
        'adsPaused': 'Number of paused ads',
        'numberOfCreative': 'Number of creatives',
        'numberOfBids': 'Number of bid placements',
        'numberOfTimeTargets': 'Number of time targets'
      };
      
      sheet.getRange(row, 4).setValue(descriptions[key] || '');
      row++;
    }
    
    // Add section for bids array
    row++;
    sheet.getRange(row, 1).setValue('BIDS ARRAY (Country Targeting)');
    sheet.getRange(row, 1, 1, 4).merge().setFontWeight('bold').setBackground('#f4b400');
    row++;
    
    if (firstCampaign.bids && firstCampaign.bids.length > 0) {
      const firstBid = firstCampaign.bids[0];
      
      sheet.getRange(row, 1).setValue('Bid Field');
      sheet.getRange(row, 2).setValue('Sample Value');
      sheet.getRange(row, 3).setValue('Type');
      sheet.getRange(row, 4).setValue('Description');
      sheet.getRange(row, 1, 1, 4).setFontWeight('bold').setBackground('#e0e0e0');
      row++;
      
      const bidDescriptions = {
        'placementId': 'Unique placement ID',
        'bid': 'Bid amount',
        'countryCode': 'Country code (e.g., US, UK, CA)',
        'countryName': 'Full country name',
        'regionCode': 'State/province code',
        'regionName': 'State/province name',
        'city': 'City name'
      };
      
      for (let key in firstBid) {
        sheet.getRange(row, 1).setValue(key);
        sheet.getRange(row, 2).setValue(String(firstBid[key]));
        sheet.getRange(row, 3).setValue(typeof firstBid[key]);
        sheet.getRange(row, 4).setValue(bidDescriptions[key] || '');
        row++;
      }
      
      // Show all bids for this campaign
      row++;
      sheet.getRange(row, 1).setValue(`This campaign has ${firstCampaign.bids.length} bid(s) targeting:`);
      sheet.getRange(row, 1, 1, 4).merge();
      row++;
      
      sheet.getRange(row, 1).setValue('Country Code');
      sheet.getRange(row, 2).setValue('Country Name');
      sheet.getRange(row, 3).setValue('Bid Amount');
      sheet.getRange(row, 4).setValue('Placement ID');
      sheet.getRange(row, 1, 1, 4).setFontWeight('bold').setBackground('#e0e0e0');
      row++;
      
      for (let bid of firstCampaign.bids) {
        sheet.getRange(row, 1).setValue(bid.countryCode);
        sheet.getRange(row, 2).setValue(bid.countryName);
        sheet.getRange(row, 3).setValue(bid.bid);
        sheet.getRange(row, 4).setValue(bid.placementId);
        row++;
      }
    }
    
    // Add section for spots array
    row++;
    sheet.getRange(row, 1).setValue('SPOTS ARRAY (Ad Placements)');
    sheet.getRange(row, 1, 1, 4).merge().setFontWeight('bold').setBackground('#34a853');
    row++;
    
    if (firstCampaign.spots && firstCampaign.spots.length > 0) {
      sheet.getRange(row, 1).setValue('Spot ID');
      sheet.getRange(row, 2).setValue('Spot Name');
      sheet.getRange(row, 1, 1, 2).setFontWeight('bold').setBackground('#e0e0e0');
      row++;
      
      for (let spot of firstCampaign.spots) {
        sheet.getRange(row, 1).setValue(spot.id);
        sheet.getRange(row, 2).setValue(spot.name);
        row++;
      }
    }
    
    // Add important notes
    row += 2;
    sheet.getRange(row, 1).setValue('IMPORTANT NOTES:');
    sheet.getRange(row, 1, 1, 4).merge().setFontWeight('bold').setBackground('#ea4335').setFontColor('white');
    row++;
    
    const notes = [
      '1. The API returns AGGREGATED stats per campaign (not broken down by country)',
      '2. Country data is in the "bids" array showing which countries are TARGETED',
      '3. To get performance BY COUNTRY, you would need to use Selenium web scraping',
      '4. All stats (clicks, impressions, cost, conversions) are campaign totals across all countries',
      '5. CTR is returned as a percentage (e.g., 5.5 = 5.5%)',
      '6. Dates must be in DD/MM/YYYY format and end date must be yesterday or earlier'
    ];
    
    for (let note of notes) {
      sheet.getRange(row, 1).setValue(note);
      sheet.getRange(row, 1, 1, 4).merge();
      row++;
    }
    
    // Format the sheet
    sheet.autoResizeColumns(1, 4);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(4, 350);
    sheet.setFrozenRows(5);
    
    // Switch to the new sheet
    ss.setActiveSheet(sheet);
    
    ui.alert('API Data Structure', `Sheet created! Check the "API_Data_Structure" tab.\n\nKey findings:\n• ${Object.keys(firstCampaign).length} main fields per campaign\n• Country data available in "bids" array\n• Stats are aggregated (not per-country)`, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error in showAPIDataStructure: ${error.toString()}`);
    ui.alert('Error: ' + error.toString());
  }
}

// ============================================================================
// POSTHOG CONVERSIONS FUNCTIONS
// ============================================================================

/**
 * Builds the HogQL query for PostHog conversions with dynamic date range
 * @param {string} startDateStr - Start date in YYYY-MM-DD format
 * @param {string} endDateStr - End date in YYYY-MM-DD format
 * @returns {string} The HogQL query string
 */
function buildPostHogQuery(startDateStr, endDateStr) {
  // Note: The query uses EST timezone filtering via toTimeZone
  const query = `
WITH first_events AS (
    SELECT 
        e.person_id,
        person.properties.first_joined_at_epoch AS first_joined_at_epoch,
        person.properties.email AS user_email,
        toTimeZone(e.timestamp, 'America/New_York') AS timestamp_est,
        person.properties.$initial_current_url AS initial_current_url,
        person.properties.$initial_referring_domain AS initial_referring_domain,
        ROW_NUMBER() OVER (PARTITION BY e.person_id ORDER BY e.timestamp ASC) AS rn
    FROM events e
    WHERE e.event IN (
        'sticky_subscription_activated',
        'chargebee_subscription_created',
        'balance_add_first_yearly_credits',
        'balance_add_monthly_credits',
        'upgate_subscription_activated'
    )
)
SELECT
    formatDateTime(toTimeZone(first_joined_at_epoch, 'America/New_York'), '%Y-%m-%d %H:%i:%s') AS first_joined_at,
    formatDateTime(timestamp_est, '%Y-%m-%d %H:%i:%s') AS timestamp_est,
    user_email,
    coalesce(
        nullIf(extractURLParameter(initial_current_url, 'campaign'), ''),
        nullIf(extractURLParameter(initial_current_url, 'Campaign'), '')
    ) AS campaign_id,
    coalesce(
        nullIf(extractURLParameter(initial_current_url, 'ClickID'), ''),
        nullIf(extractURLParameter(initial_current_url, 'clickid'), '')
    ) AS click_id,
    coalesce(
        nullIf(extractURLParameter(initial_current_url, 'Tracker'), ''),
        nullIf(extractURLParameter(initial_current_url, 'tracker'), '')
    ) AS tracker,
    coalesce(
        nullIf(extractURLParameter(initial_current_url, 'N_CLID'), ''),
        nullIf(extractURLParameter(initial_current_url, 'aclid'), '')
    ) AS n_clid,
    initial_referring_domain,
    initial_current_url
FROM first_events
WHERE rn = 1
  AND toDate(timestamp_est) >= toDate('${startDateStr}')
  AND toDate(timestamp_est) <= toDate('${endDateStr}')
  AND (
      initial_current_url LIKE '%ref=TrafficJunky%' 
      OR initial_current_url LIKE '%trafficjunky%'
      OR initial_referring_domain LIKE '%.youporn.%'
      OR initial_referring_domain LIKE '%.pornhub.%'
  )
ORDER BY timestamp_est DESC
LIMIT 2000;
  `.trim();
  
  return query;
}

/**
 * Fetches conversions from PostHog API
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @returns {Object} Response data from PostHog
 */
function fetchPostHogConversions(startDate, endDate) {
  const startDateStr = formatDateForDisplay(startDate);
  const endDateStr = formatDateForDisplay(endDate);
  
  Logger.log(`Fetching PostHog conversions from ${startDateStr} to ${endDateStr}`);
  
  const query = buildPostHogQuery(startDateStr, endDateStr);
  
  const url = `${POSTHOG_HOST}/api/projects/${POSTHOG_PROJECT_ID}/query/`;
  
  const payload = {
    "query": {
      "kind": "HogQLQuery",
      "query": query
    },
    "name": `TJ Conversions ${startDateStr} to ${endDateStr}`
  };
  
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + POSTHOG_API_KEY
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log(`PostHog API response code: ${responseCode}`);
    
    if (responseCode !== 200) {
      throw new Error(`PostHog API returned status ${responseCode}: ${responseText}`);
    }
    
    return JSON.parse(responseText);
    
  } catch (error) {
    Logger.log(`Error fetching PostHog data: ${error.toString()}`);
    throw error;
  }
}

/**
 * Writes PostHog conversion data to the sheet
 * @param {Object} data - PostHog API response data
 * @param {string} startDateStr - Start date string for display
 * @param {string} endDateStr - End date string for display
 */
function writeConversionsToSheet(data, startDateStr, endDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONVERSIONS_SHEET_NAME);
  
  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(CONVERSIONS_SHEET_NAME);
    Logger.log(`Created new sheet: ${CONVERSIONS_SHEET_NAME}`);
  }
  
  // Define headers
  const headers = [
    'First Joined At (EST)',
    'Conversion Time (EST)',
    'User Email',
    'Campaign ID',
    'Click ID',
    'Tracker',
    'N_CLID',
    'Referring Domain',
    'Initial URL',
    'Last Updated'
  ];
  
  // Clear the sheet
  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    sheet.getRange(1, 1, Math.max(lastRow, 1000), headers.length).clearContent();
  }
  
  // Write headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.setFrozenRows(1);
  
  // Process results
  const results = data.results || [];
  const columns = data.columns || [];
  
  Logger.log(`Processing ${results.length} conversion records`);
  Logger.log(`Columns: ${JSON.stringify(columns)}`);
  
  if (results.length === 0) {
    sheet.getRange(2, 1).setValue('No conversions found for the selected date range');
    sheet.getRange(2, 2).setValue(`${startDateStr} to ${endDateStr}`);
    return 0;
  }
  
  // Map column indices
  const colIndex = {};
  columns.forEach((col, idx) => {
    colIndex[col] = idx;
  });
  
  // Convert results to rows
  const rows = [];
  const now = new Date();
  
  for (let result of results) {
    const row = [
      result[colIndex['first_joined_at']] || '',
      result[colIndex['timestamp_est']] || '',
      result[colIndex['user_email']] || '',
      result[colIndex['campaign_id']] || '',
      result[colIndex['click_id']] || '',
      result[colIndex['tracker']] || '',
      result[colIndex['n_clid']] || '',
      result[colIndex['initial_referring_domain']] || '',
      result[colIndex['initial_current_url']] || '',
      now
    ];
    rows.push(row);
  }
  
  // Write data
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    
    // Format timestamp column
    sheet.getRange(2, 10, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    
    // Auto-resize columns (but limit URL column width)
    sheet.autoResizeColumns(1, 8);
    sheet.setColumnWidth(9, 300); // Initial URL column - limit width
    sheet.autoResizeColumn(10);
    
    Logger.log(`Successfully wrote ${rows.length} conversions to sheet`);
  }
  
  return rows.length;
}

/**
 * Main function to fetch and display PostHog conversions
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 */
function fetchAndWriteConversions(startDate, endDate) {
  const ui = SpreadsheetApp.getUi();
  
  // Check if PostHog is configured
  if (POSTHOG_API_KEY === "YOUR_POSTHOG_PERSONAL_API_KEY" || POSTHOG_PROJECT_ID === "YOUR_PROJECT_ID") {
    ui.alert('PostHog Not Configured', 
      'Please configure your PostHog credentials in the script:\n\n' +
      '1. Open Extensions > Apps Script\n' +
      '2. Find POSTHOG_API_KEY and POSTHOG_PROJECT_ID at the top\n' +
      '3. Replace with your actual values from PostHog\n\n' +
      'API Key: PostHog Settings > Personal API Keys\n' +
      'Project ID: From your PostHog URL (us.posthog.com/project/YOUR_ID/...)', 
      ui.ButtonSet.OK);
    return;
  }
  
  const startDateStr = formatDateForDisplay(startDate);
  const endDateStr = formatDateForDisplay(endDate);
  
  ui.alert(`Fetching PostHog conversions from ${startDateStr} to ${endDateStr} (EST)...`);
  
  try {
    const data = fetchPostHogConversions(startDate, endDate);
    const count = writeConversionsToSheet(data, startDateStr, endDateStr);
    
    ui.alert('Success!', 
      `Found ${count} TrafficJunky conversions from ${startDateStr} to ${endDateStr}.\n\n` +
      `Data written to "${CONVERSIONS_SHEET_NAME}" sheet.`, 
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error in fetchAndWriteConversions: ${error.toString()}`);
    ui.alert('Error', `Failed to fetch PostHog data:\n\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Pull today's conversions (in EST timezone)
 */
function pullTodayConversions() {
  try {
    const today = getESTDate();
    today.setHours(0, 0, 0, 0);
    
    fetchAndWriteConversions(today, today);
    
  } catch (error) {
    Logger.log("Error in pullTodayConversions: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull yesterday's conversions (in EST timezone)
 */
function pullYesterdayConversions() {
  try {
    const yesterday = getESTYesterday();
    
    fetchAndWriteConversions(yesterday, yesterday);
    
  } catch (error) {
    Logger.log("Error in pullYesterdayConversions: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull last 7 days of conversions
 */
function pullConversionsLast7Days() {
  try {
    const endDate = getESTDate();
    endDate.setHours(0, 0, 0, 0);
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 6); // 6 days before today = 7 days total
    
    fetchAndWriteConversions(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullConversionsLast7Days: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull last 30 days of conversions
 */
function pullConversionsLast30Days() {
  try {
    const endDate = getESTDate();
    endDate.setHours(0, 0, 0, 0);
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 29); // 29 days before today = 30 days total
    
    fetchAndWriteConversions(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullConversionsLast30Days: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull conversions for custom date range
 */
function pullConversionsCustomRange() {
  const ui = SpreadsheetApp.getUi();
  
  // Prompt for start date
  const startDateResponse = ui.prompt(
    'Start Date (EST)',
    'Enter start date (YYYY-MM-DD):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (startDateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  // Prompt for end date
  const endDateResponse = ui.prompt(
    'End Date (EST)',
    'Enter end date (YYYY-MM-DD):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (endDateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  try {
    let startDate = new Date(startDateResponse.getResponseText());
    let endDate = new Date(endDateResponse.getResponseText());
    
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      ui.alert('Invalid date format. Please use YYYY-MM-DD format.');
      return;
    }
    
    // If start date is after end date, show error
    if (startDate > endDate) {
      ui.alert('Error: Start date cannot be after end date.');
      return;
    }
    
    fetchAndWriteConversions(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullConversionsCustomRange: " + error.toString());
    ui.alert('Error: ' + error.toString());
  }
}

/**
 * Clear conversion data from the PostHog_Conversions sheet
 */
function clearConversionData() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Clear Conversion Data',
    'Are you sure you want to clear all PostHog conversion data?',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONVERSIONS_SHEET_NAME);
    
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 10).clearContent();
      }
      ui.alert('Conversion data cleared successfully.');
    } else {
      ui.alert('Conversion data sheet not found.');
    }
  }
}


