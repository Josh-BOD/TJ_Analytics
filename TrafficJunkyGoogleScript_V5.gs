/**
 * TrafficJunky API Data Extractor for Google Sheets - V5
 * This script pulls campaign data from TrafficJunky API and populates it into Google Sheets
 * V5: Updated PostHog query to use person.properties.ref/source instead of URL parameters
 *      Filters OUT header traffic (source='header')
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

// RedTrack Configuration
const REDTRACK_API_KEY = "M2YeuO6VDcxs5sLqJkaI";  // Get from RedTrack Settings
const REDTRACK_API_URL = "https://api.redtrack.io/conversions";
const REDTRACK_SHEET_NAME = "RedTrack_Conversions";
const REDTRACK_CONVERSION_TYPE = "Purchase";  // Conversion type to filter (e.g., purchase, lead, etc.)
const REDTRACK_CAMPAIGN_IDS = [
  "68fee7d85d1066083925def9",
  "692e7bedbc9773c782761c97",
  "6913ebf5e14b3217fa1bfa97"
];  // Campaign IDs to filter conversions

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
 * Helper function to get first day of current month in EST timezone
 */
function getESTFirstOfMonth() {
  const estNow = getESTDate();
  return new Date(estNow.getFullYear(), estNow.getMonth(), 1);
}

/**
 * Helper function to get first day of last month in EST timezone
 */
function getESTFirstOfLastMonth() {
  const estNow = getESTDate();
  return new Date(estNow.getFullYear(), estNow.getMonth() - 1, 1);
}

/**
 * Helper function to get last day of last month in EST timezone
 */
function getESTLastOfLastMonth() {
  const estNow = getESTDate();
  return new Date(estNow.getFullYear(), estNow.getMonth(), 0);
}

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Data Connector')
    .addSubMenu(ui.createMenu('🔄 All Data Sources')
      .addItem('📅 Today (EST)', 'pullAllDataToday')
      .addItem('📆 Yesterday (EST)', 'pullAllDataYesterday')
      .addItem('📊 Last 7 Days (EST)', 'pullAllDataLast7Days')
      .addItem('📊 Last 14 Days (EST)', 'pullAllDataLast14Days')
      .addItem('📈 Last 30 Days (EST)', 'pullAllDataLast30Days')
      .addItem('📅 This Month (EST)', 'pullAllDataThisMonth')
      .addItem('📅 Last Month (EST)', 'pullAllDataLastMonth')
      .addSeparator()
      .addItem('🔧 Custom Date Range (EST)', 'pullAllDataCustomRange')
      .addSeparator()
      .addItem('🗑️ Clear All Data', 'clearAllData'))
    .addSubMenu(ui.createMenu('📊 TJ Aggregated Data')
      .addItem('📅 Today (EST)', 'pullTJToday')
      .addItem('📆 Yesterday (EST)', 'pullTJYesterday')
      .addItem('📊 Last 7 Days (EST)', 'pullLast7Days')
      .addItem('📊 Last 14 Days (EST)', 'pullTJLast14Days')
      .addItem('📈 Last 30 Days (EST)', 'pullTrafficJunkyData')
      .addItem('📅 This Month (EST)', 'pullThisMonth')
      .addItem('📅 Last Month (EST)', 'pullTJLastMonth')
      .addSeparator()
      .addItem('🔧 Custom Date Range (EST)', 'pullCustomDateRange')
      .addSeparator()
      .addItem('🗑️ Clear Aggregated Data', 'clearData'))
    .addSubMenu(ui.createMenu('📅 TJ Daily Breakdown')
      .addItem('📅 Today (EST)', 'updateDailyToday')
      .addItem('📆 Yesterday (EST)', 'updateDailyYesterday')
      .addItem('📊 Last 7 Days (EST)', 'updateDailyLast7Days')
      .addItem('📊 Last 14 Days (EST)', 'updateDailyLast14Days')
      .addItem('📈 Last 30 Days (EST)', 'updateDailyLast30Days')
      .addItem('📅 This Month (EST)', 'updateDailyThisMonth')
      .addItem('📅 Last Month (EST)', 'updateDailyLastMonth')
      .addSeparator()
      .addItem('🔧 Custom Date Range (EST)', 'updateDailyCustomRange')
      .addSeparator()
      .addItem('🗑️ Clear Daily Data', 'clearDailyData'))
    .addSubMenu(ui.createMenu('🔄 PostHog Conversions')
      .addItem('📅 Today (EST)', 'pullTodayConversions')
      .addItem('📆 Yesterday (EST)', 'pullYesterdayConversions')
      .addItem('📊 Last 7 Days (EST)', 'pullConversionsLast7Days')
      .addItem('📊 Last 14 Days (EST)', 'pullConversionsLast14Days')
      .addItem('📈 Last 30 Days (EST)', 'pullConversionsLast30Days')
      .addItem('📅 This Month (EST)', 'pullConversionsThisMonth')
      .addItem('📅 Last Month (EST)', 'pullConversionsLastMonth')
      .addSeparator()
      .addItem('🔧 Custom Date Range (EST)', 'pullConversionsCustomRange')
      .addSeparator()
      .addItem('🗑️ Clear Conversion Data', 'clearConversionData'))
    .addSubMenu(ui.createMenu('📈 RedTrack Conversions')
      .addItem('📅 Today (EST)', 'pullRedTrackToday')
      .addItem('📆 Yesterday (EST)', 'pullRedTrackYesterday')
      .addItem('📊 Last 7 Days (EST)', 'pullRedTrackLast7Days')
      .addItem('📊 Last 14 Days (EST)', 'pullRedTrackLast14Days')
      .addItem('📈 Last 30 Days (EST)', 'pullRedTrackLast30Days')
      .addItem('📅 This Month (EST)', 'pullRedTrackThisMonth')
      .addItem('📅 Last Month (EST)', 'pullRedTrackLastMonth')
      .addSeparator()
      .addItem('🔧 Custom Date Range (EST)', 'pullRedTrackCustomRange')
      .addSeparator()
      .addItem('🗑️ Clear RedTrack Data', 'clearRedTrackData'))
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
 * Pull TJ data for today (EST)
 */
function pullTJToday() {
  try {
    const today = getESTDate();
    today.setHours(0, 0, 0, 0);
    
    fetchAndWriteData(today, today);
    
  } catch (error) {
    Logger.log("Error in pullTJToday: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull TJ data for yesterday (EST)
 */
function pullTJYesterday() {
  try {
    const yesterday = getESTYesterday();
    
    fetchAndWriteData(yesterday, yesterday);
    
  } catch (error) {
    Logger.log("Error in pullTJYesterday: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull TJ data for last 14 days
 */
function pullTJLast14Days() {
  try {
    const endDate = getESTYesterday();
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 13);
    
    fetchAndWriteData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullTJLast14Days: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull TJ data for last month
 */
function pullTJLastMonth() {
  try {
    const startDate = getESTFirstOfLastMonth();
    const endDate = getESTLastOfLastMonth();
    
    fetchAndWriteData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullTJLastMonth: " + error.toString());
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
    
    // Validate dates (allow current day for real-time stats)
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
  
  // Allow current day pulls for real-time stats (EST timezone)
  Logger.log(`Fetching TJ data - allowing current day pulls`);
  
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
 * Helper function to format date as MM/DD/YYYY (US format for RedTrack API)
 */
function formatDateUS(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${month}/${day}/${year}`;
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
 * Update daily data for today (EST)
 */
function updateDailyToday() {
  try {
    const today = getESTDate();
    today.setHours(0, 0, 0, 0);
    
    updateDailyData(today, today);
    
  } catch (error) {
    Logger.log("Error in updateDailyToday: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Update daily data for yesterday (EST)
 */
function updateDailyYesterday() {
  try {
    const yesterday = getESTYesterday();
    
    updateDailyData(yesterday, yesterday);
    
  } catch (error) {
    Logger.log("Error in updateDailyYesterday: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Update daily data for last 30 days
 */
function updateDailyLast30Days() {
  try {
    const endDate = getESTYesterday();
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 29);
    
    updateDailyData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in updateDailyLast30Days: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Update daily data for last month
 */
function updateDailyLastMonth() {
  try {
    const startDate = getESTFirstOfLastMonth();
    const endDate = getESTLastOfLastMonth();
    
    updateDailyData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in updateDailyLastMonth: " + error.toString());
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
    
    // Validate dates (allow current day for real-time stats)
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
  
  // Allow current day pulls for real-time stats (EST timezone)
  Logger.log(`Fetching daily data - allowing current day pulls`);
  
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
 * Fetch and write daily data without UI alerts (for aggregate data pull)
 */
function fetchAllDataForDaily(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Allow current day pulls for real-time stats (EST timezone)
  Logger.log(`Fetching daily data for aggregate - allowing current day pulls`);
  
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
    Logger.log(`Fetching daily data for ${dateStr}`);
    
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
      Logger.log(`Error fetching daily data for ${dateStr}: ${error.toString()}`);
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
    
    Logger.log(`Successfully updated ${dailyData.length} daily records`);
    return dailyData.length;
  } else {
    Logger.log("No daily data found in API response");
    return 0;
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
 * DEDUPLICATED: Only returns the first conversion event per unique user email
 * @param {string} startDateStr - Start date in YYYY-MM-DD format
 * @param {string} endDateStr - End date in YYYY-MM-DD format
 * @returns {string} The HogQL query string
 */
function buildPostHogQuery(startDateStr, endDateStr) {
  // V5: Updated to use person.properties.ref and filter out header traffic
  // Deduplication done client-side in Google Apps Script
  const query = `
SELECT
    formatDateTime(toTimeZone(person.properties.first_joined_at_epoch, 'America/New_York'), '%Y-%m-%d %H:%i:%s') AS first_joined_at,
    formatDateTime(toTimeZone(e.timestamp, 'America/New_York'), '%Y-%m-%d %H:%i:%s') AS timestamp_est,
    person.properties.email AS user_email,
    person.properties.ref AS ref,
    person.properties.source AS source,
    coalesce(
        nullIf(extractURLParameter(person.properties.$initial_current_url, 'campaign'), ''),
        nullIf(extractURLParameter(person.properties.$initial_current_url, 'Campaign'), '')
    ) AS campaign_id,
    coalesce(
        nullIf(extractURLParameter(person.properties.$initial_current_url, 'ClickID'), ''),
        nullIf(extractURLParameter(person.properties.$initial_current_url, 'clickid'), '')
    ) AS click_id,
    coalesce(
        nullIf(extractURLParameter(person.properties.$initial_current_url, 'Tracker'), ''),
        nullIf(extractURLParameter(person.properties.$initial_current_url, 'tracker'), '')
    ) AS tracker,
    coalesce(
        nullIf(extractURLParameter(person.properties.$initial_current_url, 'N_CLID'), ''),
        nullIf(extractURLParameter(person.properties.$initial_current_url, 'aclid'), '')
    ) AS n_clid,
    person.properties.$initial_referring_domain AS initial_referring_domain,
    person.properties.$initial_current_url AS initial_current_url
FROM events e
WHERE e.event IN (
    'sticky_subscription_activated',
    'chargebee_subscription_created',
    'balance_add_first_yearly_credits',
    'balance_add_monthly_credits',
    'upgate_subscription_activated'
)
AND toDate(toTimeZone(e.timestamp, 'America/New_York')) >= toDate('${startDateStr}')
AND toDate(toTimeZone(e.timestamp, 'America/New_York')) <= toDate('${endDateStr}')
AND LOWER(person.properties.ref) = 'trafficjunky'
AND (person.properties.source IS NULL OR LOWER(person.properties.source) != 'header')
ORDER BY e.timestamp ASC
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
  
  Logger.log(`=== PostHog Fetch Start ===`);
  Logger.log(`Start Date: ${startDateStr}, End Date: ${endDateStr}`);
  Logger.log(`Start Date object: ${startDate}, End Date object: ${endDate}`);
  
  // Calculate the number of days
  const msPerDay = 1000 * 60 * 60 * 24;
  const daysDiff = Math.ceil((endDate.getTime() - startDate.getTime()) / msPerDay) + 1;
  
  Logger.log(`Days difference calculated: ${daysDiff}`);
  
  // If more than 1 day, fetch day by day to avoid timeout
  if (daysDiff > 1) {
    Logger.log(`*** Using day-by-day fetch for ${daysDiff} days ***`);
    return fetchPostHogConversionsByDay(startDate, endDate);
  }
  
  Logger.log(`*** Using single fetch for 1 day ***`);
  // Single day fetch
  return fetchPostHogConversionsSingle(startDateStr, endDateStr);
}

/**
 * Fetch PostHog conversions for a single date range (internal helper)
 */
function fetchPostHogConversionsSingle(startDateStr, endDateStr) {
  Logger.log(`  [Single] Fetching ${startDateStr} to ${endDateStr}`);
  
  const query = buildPostHogQuery(startDateStr, endDateStr);
  
  Logger.log(`  [Single] Query built, length: ${query.length} chars`);
  
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
    Logger.log(`  [Single] Calling PostHog API...`);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log(`  [Single] Response code: ${responseCode}, Response length: ${responseText.length} chars`);
    
    if (responseCode !== 200) {
      Logger.log(`  [Single] ERROR Response: ${responseText.substring(0, 500)}`);
      throw new Error(`PostHog API returned status ${responseCode}: ${responseText}`);
    }
    
    const parsed = JSON.parse(responseText);
    Logger.log(`  [Single] Parsed OK. Results: ${parsed.results ? parsed.results.length : 'N/A'}`);
    
    return parsed;
    
  } catch (error) {
    Logger.log(`  [Single] EXCEPTION: ${error.toString()}`);
    throw error;
  }
}

/**
 * Fetch PostHog conversions day by day with INCREMENTAL WRITING to sheet
 * Data appears in real-time as each day is fetched
 * DEDUPLICATION: Only keeps first conversion per unique email (done client-side)
 */
function fetchPostHogConversionsByDay(startDate, endDate) {
  Logger.log(`=== Day-by-Day Fetch Starting (Incremental Write + Dedup) ===`);
  Logger.log(`From: ${startDate} To: ${endDate}`);
  
  // Prepare the sheet first
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONVERSIONS_SHEET_NAME);
  
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
    'Ref',
    'Source',
    'Last Updated'
  ];
  
  // Clear existing data
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  
  // Write headers if needed
  if (lastRow === 0 || sheet.getRange(1, 1).getValue() !== headers[0]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    sheet.setFrozenRows(1);
  }
  
  // Skip initial flush - let it happen naturally with first write
  
  let columns = null;
  const currentDate = new Date(startDate);
  let totalFetched = 0;
  let totalFromAPI = 0;
  let daysProcessed = 0;
  let errorsEncountered = 0;
  let currentRow = 2; // Start writing from row 2
  const now = new Date();
  const seenEmails = new Set(); // Track emails for deduplication
  
  while (currentDate <= endDate) {
    const dateStr = formatDateForDisplay(currentDate);
    daysProcessed++;
    
    Logger.log(`--- Day ${daysProcessed}: ${dateStr} ---`);
    
    try {
      const dayData = fetchPostHogConversionsSingle(dateStr, dateStr);
      
      Logger.log(`  API returned: columns=${dayData.columns ? dayData.columns.length : 'none'}, results=${dayData.results ? dayData.results.length : 'none'}`);
      
      // Store columns from first successful response
      if (!columns && dayData.columns) {
        columns = dayData.columns;
        Logger.log(`  Captured columns: ${JSON.stringify(columns)}`);
      }
      
      // IMMEDIATELY write this day's results to the sheet (with deduplication by email)
      if (dayData.results && Array.isArray(dayData.results) && dayData.results.length > 0 && columns) {
        const colIndex = {};
        columns.forEach((col, idx) => { colIndex[col] = idx; });
        
        totalFromAPI += dayData.results.length;
        
        // Filter out duplicate emails (keep first occurrence)
        const rows = [];
        for (const result of dayData.results) {
          const email = result[colIndex['user_email']] || '';
          
          // Skip if we've already seen this email
          if (email && seenEmails.has(email)) {
            continue;
          }
          
          // Mark email as seen
          if (email) {
            seenEmails.add(email);
          }
          
          rows.push([
            result[colIndex['first_joined_at']] || '',
            result[colIndex['timestamp_est']] || '',
            email,
            result[colIndex['campaign_id']] || '',
            result[colIndex['click_id']] || '',
            result[colIndex['tracker']] || '',
            result[colIndex['n_clid']] || '',
            result[colIndex['initial_referring_domain']] || '',
            result[colIndex['initial_current_url']] || '',
            result[colIndex['ref']] || '',
            result[colIndex['source']] || '',
            now
          ]);
        }
        
        // Write to sheet immediately (only unique emails)
        if (rows.length > 0) {
          sheet.getRange(currentRow, 1, rows.length, headers.length).setValues(rows);
          
          // Only flush every 3 days to reduce overhead
          if (daysProcessed % 3 === 0) {
            SpreadsheetApp.flush();
          }
          
          currentRow += rows.length;
          totalFetched += rows.length;
        }
        
        Logger.log(`  ✓ Wrote ${rows.length} unique conversions for ${dateStr} (${dayData.results.length} from API, running total: ${totalFetched})`);
      } else if (dayData.results && dayData.results.length === 0) {
        Logger.log(`  ⚠ No conversions for ${dateStr}`);
      } else {
        Logger.log(`  ⚠ No results array in response for ${dateStr}`);
      }
      
    } catch (error) {
      errorsEncountered++;
      Logger.log(`  ✗ ERROR fetching ${dateStr}: ${error.toString()}`);
      // Continue with other days even if one fails
    }
    
    // Move to next day
    currentDate.setDate(currentDate.getDate() + 1);
    
    // Minimal delay to avoid rate limiting
    Utilities.sleep(50);
  }
  
  // Final flush to ensure all data is displayed
  SpreadsheetApp.flush();
  
  Logger.log(`=== Day-by-Day Fetch Complete ===`);
  Logger.log(`Days processed: ${daysProcessed}, Errors: ${errorsEncountered}`);
  Logger.log(`Total from API: ${totalFromAPI}, Unique conversions written: ${totalFetched}`);
  
  // Return summary (data already written to sheet)
  return {
    columns: columns || [],
    results: [], // Empty - data already written incrementally
    totalWritten: totalFetched
  };
}

/**
 * Writes PostHog conversion data to the sheet
 * @param {Object} data - PostHog API response data
 * @param {string} startDateStr - Start date string for display
 * @param {string} endDateStr - End date string for display
 */
function writeConversionsToSheet(data, startDateStr, endDateStr) {
  // Check if data was already written incrementally (day-by-day fetch)
  if (data.totalWritten !== undefined) {
    Logger.log(`Data already written incrementally: ${data.totalWritten} conversions`);
    return data.totalWritten;
  }
  
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
    'Ref',
    'Source',
    'Last Updated'
  ];
  
  // Clear only the data rows (faster than clearing large range)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  
  // Write headers only if they don't exist
  if (lastRow === 0 || sheet.getRange(1, 1).getValue() !== headers[0]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    sheet.setFrozenRows(1);
  }
  
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
      result[colIndex['ref']] || '',
      result[colIndex['source']] || '',
      now
    ];
    rows.push(row);
  }
  
  // Write data in batches to avoid timeout
  if (rows.length > 0) {
    const BATCH_SIZE = 500;
    for (let i = 0; i < rows.length; i += BATCH_SIZE) {
      const batch = rows.slice(i, i + BATCH_SIZE);
      const startRow = 2 + i;
      sheet.getRange(startRow, 1, batch.length, headers.length).setValues(batch);
      Logger.log(`Wrote batch ${Math.floor(i/BATCH_SIZE) + 1}: rows ${startRow} to ${startRow + batch.length - 1}`);
      
      // Flush changes to avoid timeout on large writes
      if (i + BATCH_SIZE < rows.length) {
        SpreadsheetApp.flush();
      }
    }
    
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
 * Pull last 14 days of PostHog conversions
 */
function pullConversionsLast14Days() {
  try {
    const endDate = getESTDate();
    endDate.setHours(0, 0, 0, 0);
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 13);
    
    fetchAndWriteConversions(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullConversionsLast14Days: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull this month's PostHog conversions
 */
function pullConversionsThisMonth() {
  try {
    const endDate = getESTDate();
    endDate.setHours(0, 0, 0, 0);
    
    const startDate = getESTFirstOfMonth();
    
    fetchAndWriteConversions(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullConversionsThisMonth: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull last month's PostHog conversions
 */
function pullConversionsLastMonth() {
  try {
    const startDate = getESTFirstOfLastMonth();
    const endDate = getESTLastOfLastMonth();
    
    fetchAndWriteConversions(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullConversionsLastMonth: " + error.toString());
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
        sheet.getRange(2, 1, lastRow - 1, 12).clearContent();
      }
      ui.alert('Conversion data cleared successfully.');
    } else {
      ui.alert('Conversion data sheet not found.');
    }
  }
}

// ============================================================================
// REDTRACK CONVERSIONS FUNCTIONS
// ============================================================================

/**
 * Fetches conversions from RedTrack API
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @returns {Object} Response data from RedTrack
 */
function fetchRedTrackConversions(startDate, endDate) {
  // Format dates as YYYY-MM-DD for RedTrack API
  const startDateStr = formatDateForDisplay(startDate);
  const endDateStr = formatDateForDisplay(endDate);
  
  Logger.log(`Fetching RedTrack conversions from ${startDateStr} to ${endDateStr}`);
  
  const options = {
    'method': 'get',
    'contentType': 'application/json',
    'muteHttpExceptions': true
  };
  
  const perPage = 10000;
  let allConversions = [];
  let page = 1;
  let totalReported = 0;
  let hasMore = true;
  
  try {
    // Paginate through all results
    while (hasMore) {
      const url = `${REDTRACK_API_URL}?api_key=${REDTRACK_API_KEY}&date_from=${startDateStr}&date_to=${endDateStr}&type=${REDTRACK_CONVERSION_TYPE}&page=${page}&per=${perPage}`;
      
      Logger.log(`Fetching page ${page}...`);
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode !== 200) {
        throw new Error(`RedTrack API returned status ${responseCode}: ${responseText}`);
      }
      
      const data = JSON.parse(responseText);
      
      // Get total on first page
      if (page === 1) {
        totalReported = data.total || 0;
        Logger.log(`RedTrack reports ${totalReported} total conversions`);
      }
      
      // Extract items from response
      let pageItems = [];
      if (data && data.items && Array.isArray(data.items)) {
        pageItems = data.items;
      } else if (Array.isArray(data)) {
        pageItems = data;
      }
      
      Logger.log(`Page ${page}: fetched ${pageItems.length} conversions`);
      
      allConversions = allConversions.concat(pageItems);
      
      // Check if we need more pages
      if (pageItems.length < perPage || allConversions.length >= totalReported) {
        hasMore = false;
      } else {
        page++;
        // Safety limit to prevent infinite loops
        if (page > 100) {
          Logger.log('Warning: Reached page limit of 100, stopping pagination');
          hasMore = false;
        }
      }
    }
    
    Logger.log(`Total conversions fetched from API: ${allConversions.length} (total reported: ${totalReported})`);
    
    // Filter locally by conversion type (in case API filter didn't work perfectly)
    const typeFiltered = allConversions.filter(conv => {
      const convType = conv.type || conv.conversion_type || '';
      // Case-insensitive comparison
      return convType.toLowerCase() === REDTRACK_CONVERSION_TYPE.toLowerCase();
    });
    
    Logger.log(`Conversions after filtering by type '${REDTRACK_CONVERSION_TYPE}': ${typeFiltered.length}`);
    
    // Filter locally by campaign IDs
    const filteredConversions = typeFiltered.filter(conv => {
      // Check common field names for campaign ID
      const convCampaignId = conv.campaign_id || conv.campaignId || conv.cid || conv.campaign || '';
      return REDTRACK_CAMPAIGN_IDS.includes(convCampaignId);
    });
    
    Logger.log(`Conversions after filtering by campaign IDs: ${filteredConversions.length}`);
    
    return filteredConversions;
    
  } catch (error) {
    Logger.log(`Error fetching RedTrack data: ${error.toString()}`);
    throw error;
  }
}

/**
 * Writes RedTrack conversion data to the sheet
 * @param {Array} data - RedTrack API response data (array of conversions)
 * @param {string} startDateStr - Start date string for display
 * @param {string} endDateStr - End date string for display
 */
function writeRedTrackToSheet(data, startDateStr, endDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(REDTRACK_SHEET_NAME);
  
  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(REDTRACK_SHEET_NAME);
    Logger.log(`Created new sheet: ${REDTRACK_SHEET_NAME}`);
  }
  
  // Handle both array and object responses
  let conversions = [];
  if (Array.isArray(data)) {
    conversions = data;
  } else if (data && typeof data === 'object') {
    // If it's an object with a data/conversions property
    conversions = data.data || data.conversions || data.results || [data];
  }
  
  Logger.log(`Processing ${conversions.length} RedTrack conversion records`);
  
  // Clear the sheet
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1 && lastCol > 0) {
    // Clear data rows only (row 2 onwards)
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  } else if (lastRow === 1) {
    // Only headers exist, nothing to clear
  } else if (lastRow > 0) {
    // Clear everything if structure is unclear
    sheet.clearContents();
  }
  
  if (conversions.length === 0) {
    sheet.getRange(1, 1).setValue('No conversions found for the selected date range');
    sheet.getRange(1, 2).setValue(`${startDateStr} to ${endDateStr}`);
    return 0;
  }
  
  // Dynamically get headers from the first conversion record
  const firstConversion = conversions[0];
  const headers = Object.keys(firstConversion);
  headers.push('Last Updated'); // Add our timestamp column
  
  // Write headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#34a853').setFontColor('white');
  sheet.setFrozenRows(1);
  
  // Convert conversions to rows
  const rows = [];
  const now = new Date();
  
  for (let conversion of conversions) {
    const row = headers.slice(0, -1).map(header => {
      const value = conversion[header];
      // Handle nested objects by converting to string
      if (value !== null && typeof value === 'object') {
        return JSON.stringify(value);
      }
      return value !== undefined ? value : '';
    });
    row.push(now); // Add timestamp
    rows.push(row);
  }
  
  // Write data
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    
    // Format timestamp column (last column)
    sheet.getRange(2, headers.length, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    
    // Auto-resize columns (limit to first 15 to avoid too wide)
    const colsToResize = Math.min(headers.length, 15);
    sheet.autoResizeColumns(1, colsToResize);
    
    Logger.log(`Successfully wrote ${rows.length} RedTrack conversions to sheet`);
  }
  
  return rows.length;
}

/**
 * Main function to fetch and display RedTrack conversions
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 */
function fetchAndWriteRedTrack(startDate, endDate) {
  const ui = SpreadsheetApp.getUi();
  
  // Check if RedTrack is configured
  if (REDTRACK_API_KEY === "YOUR_REDTRACK_API_KEY") {
    ui.alert('RedTrack Not Configured', 
      'Please configure your RedTrack API key in the script:\n\n' +
      '1. Open Extensions > Apps Script\n' +
      '2. Find REDTRACK_API_KEY at the top\n' +
      '3. Replace with your actual API key from RedTrack settings', 
      ui.ButtonSet.OK);
    return;
  }
  
  const startDateStr = formatDateForDisplay(startDate);
  const endDateStr = formatDateForDisplay(endDate);
  
  ui.alert(`Fetching RedTrack conversions from ${startDateStr} to ${endDateStr} (EST)...`);
  
  try {
    const data = fetchRedTrackConversions(startDate, endDate);
    const count = writeRedTrackToSheet(data, startDateStr, endDateStr);
    
    ui.alert('Success!', 
      `Found ${count} RedTrack conversions from ${startDateStr} to ${endDateStr}.\n\n` +
      `Data written to "${REDTRACK_SHEET_NAME}" sheet.`, 
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error in fetchAndWriteRedTrack: ${error.toString()}`);
    ui.alert('Error', `Failed to fetch RedTrack data:\n\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Pull today's RedTrack conversions (in EST timezone)
 */
function pullRedTrackToday() {
  try {
    const today = getESTDate();
    today.setHours(0, 0, 0, 0);
    
    fetchAndWriteRedTrack(today, today);
    
  } catch (error) {
    Logger.log("Error in pullRedTrackToday: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull yesterday's RedTrack conversions (in EST timezone)
 */
function pullRedTrackYesterday() {
  try {
    const yesterday = getESTYesterday();
    
    fetchAndWriteRedTrack(yesterday, yesterday);
    
  } catch (error) {
    Logger.log("Error in pullRedTrackYesterday: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull last 7 days of RedTrack conversions
 */
function pullRedTrackLast7Days() {
  try {
    const endDate = getESTDate();
    endDate.setHours(0, 0, 0, 0);
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 6); // 6 days before today = 7 days total
    
    fetchAndWriteRedTrack(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullRedTrackLast7Days: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull last 30 days of RedTrack conversions
 */
function pullRedTrackLast30Days() {
  try {
    const endDate = getESTDate();
    endDate.setHours(0, 0, 0, 0);
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 29); // 29 days before today = 30 days total
    
    fetchAndWriteRedTrack(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullRedTrackLast30Days: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull last 14 days of RedTrack conversions
 */
function pullRedTrackLast14Days() {
  try {
    const endDate = getESTDate();
    endDate.setHours(0, 0, 0, 0);
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 13);
    
    fetchAndWriteRedTrack(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullRedTrackLast14Days: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull this month's RedTrack conversions
 */
function pullRedTrackThisMonth() {
  try {
    const endDate = getESTDate();
    endDate.setHours(0, 0, 0, 0);
    
    const startDate = getESTFirstOfMonth();
    
    fetchAndWriteRedTrack(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullRedTrackThisMonth: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull last month's RedTrack conversions
 */
function pullRedTrackLastMonth() {
  try {
    const startDate = getESTFirstOfLastMonth();
    const endDate = getESTLastOfLastMonth();
    
    fetchAndWriteRedTrack(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullRedTrackLastMonth: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull RedTrack conversions for custom date range
 */
function pullRedTrackCustomRange() {
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
    
    fetchAndWriteRedTrack(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullRedTrackCustomRange: " + error.toString());
    ui.alert('Error: ' + error.toString());
  }
}

/**
 * Clear RedTrack data from the RedTrack_Conversions sheet
 */
function clearRedTrackData() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Clear RedTrack Data',
    'Are you sure you want to clear RedTrack conversion data? (Formulas in columns beyond the data will be preserved)',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(REDTRACK_SHEET_NAME);
    
    if (sheet) {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      
      if (lastRow > 1 && lastCol > 0) {
        // Clear only data rows (row 2 onwards), preserving headers and formulas beyond data columns
        sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      }
      ui.alert('RedTrack data cleared successfully. Headers and formulas beyond data columns are preserved.');
    } else {
      ui.alert('RedTrack data sheet not found.');
    }
  }
}

// ============================================================================
// AGGREGATE DATA FUNCTIONS - Pull from all sources at once
// ============================================================================

/**
 * Main function to fetch data from all sources
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 */
function fetchAllData(startDate, endDate) {
  const ui = SpreadsheetApp.getUi();
  const startDateStr = formatDateForDisplay(startDate);
  const endDateStr = formatDateForDisplay(endDate);
  
  ui.alert(`Fetching data from all sources for ${startDateStr} to ${endDateStr} (EST)...\n\nThis may take a moment.`);
  
  let results = {
    tjAggregated: { success: false, message: '' },
    tjDaily: { success: false, message: '' },
    posthog: { success: false, message: '' },
    redtrack: { success: false, message: '' }
  };
  
  // 1. Fetch TrafficJunky Aggregated Data
  try {
    Logger.log('Fetching TrafficJunky Aggregated data...');
    fetchAndWriteData(startDate, endDate);
    results.tjAggregated = { success: true, message: 'TJ Aggregated data loaded' };
  } catch (error) {
    results.tjAggregated = { success: false, message: `TJ Aggregated Error: ${error.toString()}` };
    Logger.log(`TrafficJunky Aggregated error: ${error.toString()}`);
  }
  
  // 2. Fetch TrafficJunky Daily Breakdown Data
  try {
    Logger.log('Fetching TrafficJunky Daily Breakdown data...');
    fetchAllDataForDaily(startDate, endDate);
    results.tjDaily = { success: true, message: 'TJ Daily Breakdown data loaded' };
  } catch (error) {
    results.tjDaily = { success: false, message: `TJ Daily Error: ${error.toString()}` };
    Logger.log(`TrafficJunky Daily error: ${error.toString()}`);
  }
  
  // 3. Fetch PostHog Conversions
  try {
    Logger.log('Fetching PostHog conversions...');
    const phData = fetchPostHogConversions(startDate, endDate);
    const phCount = writeConversionsToSheet(phData, startDateStr, endDateStr);
    results.posthog = { success: true, message: `PostHog: ${phCount} conversions` };
  } catch (error) {
    results.posthog = { success: false, message: `PostHog Error: ${error.toString()}` };
    Logger.log(`PostHog error: ${error.toString()}`);
  }
  
  // 4. Fetch RedTrack Conversions
  try {
    Logger.log('Fetching RedTrack conversions...');
    const rtData = fetchRedTrackConversions(startDate, endDate);
    const rtCount = writeRedTrackToSheet(rtData, startDateStr, endDateStr);
    results.redtrack = { success: true, message: `RedTrack: ${rtCount} conversions` };
  } catch (error) {
    results.redtrack = { success: false, message: `RedTrack Error: ${error.toString()}` };
    Logger.log(`RedTrack error: ${error.toString()}`);
  }
  
  // Show summary
  const summary = [
    `Data fetch complete for ${startDateStr} to ${endDateStr}:`,
    '',
    `${results.tjAggregated.success ? '✓' : '✗'} ${results.tjAggregated.message}`,
    `${results.tjDaily.success ? '✓' : '✗'} ${results.tjDaily.message}`,
    `${results.posthog.success ? '✓' : '✗'} ${results.posthog.message}`,
    `${results.redtrack.success ? '✓' : '✗'} ${results.redtrack.message}`
  ].join('\n');
  
  ui.alert('All Data Sources', summary, ui.ButtonSet.OK);
}

/**
 * Pull all data for today (EST)
 */
function pullAllDataToday() {
  try {
    const today = getESTDate();
    today.setHours(0, 0, 0, 0);
    
    fetchAllData(today, today);
    
  } catch (error) {
    Logger.log("Error in pullAllDataToday: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull all data for yesterday (EST)
 */
function pullAllDataYesterday() {
  try {
    const yesterday = getESTYesterday();
    
    fetchAllData(yesterday, yesterday);
    
  } catch (error) {
    Logger.log("Error in pullAllDataYesterday: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull all data for last 7 days
 */
function pullAllDataLast7Days() {
  try {
    const endDate = getESTDate();
    endDate.setHours(0, 0, 0, 0);
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 6);
    
    fetchAllData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullAllDataLast7Days: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull all data for last 30 days
 */
function pullAllDataLast30Days() {
  try {
    const endDate = getESTDate();
    endDate.setHours(0, 0, 0, 0);
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 29);
    
    fetchAllData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullAllDataLast30Days: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull all data for last 14 days
 */
function pullAllDataLast14Days() {
  try {
    const endDate = getESTDate();
    endDate.setHours(0, 0, 0, 0);
    
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 13);
    
    fetchAllData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullAllDataLast14Days: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull all data for this month
 */
function pullAllDataThisMonth() {
  try {
    const endDate = getESTDate();
    endDate.setHours(0, 0, 0, 0);
    
    const startDate = getESTFirstOfMonth();
    
    fetchAllData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullAllDataThisMonth: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull all data for last month
 */
function pullAllDataLastMonth() {
  try {
    const startDate = getESTFirstOfLastMonth();
    const endDate = getESTLastOfLastMonth();
    
    fetchAllData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullAllDataLastMonth: " + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Pull all data for custom date range
 */
function pullAllDataCustomRange() {
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
    
    if (startDate > endDate) {
      ui.alert('Error: Start date cannot be after end date.');
      return;
    }
    
    fetchAllData(startDate, endDate);
    
  } catch (error) {
    Logger.log("Error in pullAllDataCustomRange: " + error.toString());
    ui.alert('Error: ' + error.toString());
  }
}

/**
 * Clear data from all data sources (preserves headers and formulas beyond data columns)
 */
function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Clear All Data',
    'Are you sure you want to clear data from ALL sources?\n\n' +
    '• TJ Aggregated Data\n' +
    '• TJ Daily Breakdown\n' +
    '• PostHog Conversions\n' +
    '• RedTrack Conversions\n\n' +
    'Headers and formulas beyond data columns will be preserved.',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let cleared = [];
    let errors = [];
    
    // Clear TJ Aggregated Data
    try {
      const tjSheet = ss.getSheetByName(SHEET_NAME);
      if (tjSheet) {
        const lastRow = tjSheet.getLastRow();
        if (lastRow > 0) {
          tjSheet.getRange(1, 1, lastRow, 16).clearContent();
        }
        cleared.push('TJ Aggregated Data');
      }
    } catch (e) {
      errors.push('TJ Aggregated: ' + e.toString());
    }
    
    // Clear TJ Daily Breakdown Data
    try {
      const dailySheet = ss.getSheetByName(DAILY_SHEET_NAME);
      if (dailySheet) {
        const lastRow = dailySheet.getLastRow();
        if (lastRow > 1) {
          // Clear only columns A-L (12 columns), preserving headers and formulas beyond
          dailySheet.getRange(2, 1, lastRow - 1, 12).clearContent();
        }
        cleared.push('TJ Daily Breakdown');
      }
    } catch (e) {
      errors.push('TJ Daily: ' + e.toString());
    }
    
    // Clear PostHog Conversions
    try {
      const phSheet = ss.getSheetByName(CONVERSIONS_SHEET_NAME);
      if (phSheet) {
        const lastRow = phSheet.getLastRow();
        if (lastRow > 1) {
          phSheet.getRange(2, 1, lastRow - 1, 12).clearContent();
        }
        cleared.push('PostHog Conversions');
      }
    } catch (e) {
      errors.push('PostHog: ' + e.toString());
    }
    
    // Clear RedTrack Conversions
    try {
      const rtSheet = ss.getSheetByName(REDTRACK_SHEET_NAME);
      if (rtSheet) {
        const lastRow = rtSheet.getLastRow();
        const lastCol = rtSheet.getLastColumn();
        if (lastRow > 1 && lastCol > 0) {
          rtSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
        }
        cleared.push('RedTrack Conversions');
      }
    } catch (e) {
      errors.push('RedTrack: ' + e.toString());
    }
    
    // Show summary
    let message = '';
    if (cleared.length > 0) {
      message += 'Cleared:\n• ' + cleared.join('\n• ');
    }
    if (errors.length > 0) {
      message += '\n\nErrors:\n• ' + errors.join('\n• ');
    }
    if (cleared.length === 0 && errors.length === 0) {
      message = 'No data sheets found to clear.';
    }
    
    ui.alert('Clear All Data', message, ui.ButtonSet.OK);
  }
}


