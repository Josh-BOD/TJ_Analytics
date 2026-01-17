/**
 * TJ Control Panel V5 - Google Apps Script
 * 
 * A dashboard for bid optimization with multi-period stats (Today, Yesterday, 7-Day).
 * Uses BID-LEVEL stats for all periods - granular per-bid T/Y/7D comparisons.
 * 
 * V4 ADDITIONS:
 * - Pivot View sheet: Hierarchical view by Strategy > Sub Strategy > Campaign
 * - Edit columns at end of pivot for bid/budget changes
 * - Update functions that work from pivot view
 * 
 * V3 FEATURES:
 * - Campaign budget management: View and update daily budgets
 * - Budget columns: Daily Budget, Budget Left, New Budget
 * - Budget Logs sheet: Track all budget changes
 * 
 * V2 FEATURES:
 * - Daily Stats sheet: 7 days of campaign-level daily data
 * - Dashboard sheet: Campaign selector with CPA/Spend/eCPM/CTR trends chart
 * 
 * SETUP INSTRUCTIONS:
 * 1. Create a new Google Sheet (or use existing with BidManager)
 * 2. Go to Extensions > Apps Script
 * 3. Create a new file and paste this entire script
 * 4. Save the script (Ctrl+S / Cmd+S)
 * 5. Refresh your Google Sheet
 * 6. You should see a "Control Panel" menu appear
 * 
 * REQUIREMENTS:
 * - Legend sheet with Campaign IDs in Column A
 * - Bid Logs sheet (created by BidManager when bids are updated)
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const CP_API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039";
const CP_API_BASE_URL = "https://api.trafficjunky.com/api";
const CP_SHEET_NAME = "Control Panel";
const CP_DAILY_SHEET_NAME = "Daily Stats";
const CP_DASHBOARD_SHEET_NAME = "Dashboard";
const CP_PIVOT_SHEET_NAME = "Pivot View";
const CP_API_TIMEZONE = "America/New_York";

// ============================================================================
// MENU SETUP
// ============================================================================

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Check if BidManager menu already exists, add Control Panel as separate menu
  ui.createMenu('Control Panel')
    .addItem('🔄 Refresh All Data', 'refreshControlPanel')
    .addSeparator()
    .addSubMenu(ui.createMenu('💰 Bid Management')
      .addItem('📋 Copy Bids to New Column', 'cpCopyBidsToNew')
      .addItem('📈 Calculate Bid Changes', 'cpCalculateBidChanges')
      .addItem('🚀 UPDATE BIDS IN TJ', 'cpUpdateBids'))
    .addSubMenu(ui.createMenu('💵 Budget Management')
      .addItem('📋 Copy Budgets to New Column', 'cpCopyBudgetsToNew')
      .addItem('📈 Calculate Budget Changes', 'cpCalculateBudgetChanges')
      .addItem('🚀 UPDATE BUDGETS IN TJ', 'cpUpdateBudgets'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Pivot View (V5)')
      .addItem('📊 Build/Refresh Pivot View', 'cpBuildPivotView')
      .addItem('📋 Copy Bids (Pivot)', 'cpCopyBidsPivot')
      .addItem('📋 Copy Budgets (Pivot)', 'cpCopyBudgetsPivot')
      .addItem('🚀 UPDATE FROM PIVOT', 'cpUpdateFromPivot'))
    .addSubMenu(ui.createMenu('📊 Daily Dashboard')
      .addItem('📥 Pull Daily Stats (7 Days)', 'cpPullDailyStats')
      .addItem('📈 Build/Refresh Dashboard', 'cpBuildDashboard'))
    .addSeparator()
    .addItem('🗑️ Clear Data', 'cpClearData')
    .addSeparator()
    .addItem('🐛 Debug: Test Parallel Fetch', 'cpDebugParallelFetch')
    .addToUi();
}

/**
 * onEdit trigger - handles checkbox clicks in Pivot View to navigate to Dashboard
 * When a checkbox in column Z (Dashboard) of Pivot View is checked:
 * 1. Gets the campaign info from that row
 * 2. Sets the Dashboard dropdown (B3) to that campaign
 * 3. Navigates to the Dashboard sheet
 * 4. Unchecks the checkbox
 */
function onEdit(e) {
  // Only process if we have event info
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  
  // Only process edits in Pivot View sheet
  if (sheetName !== CP_PIVOT_SHEET_NAME) return;
  
  // Only process edits in column Z (column 26) - the checkbox column
  const col = e.range.getColumn();
  if (col !== 26) return;
  
  // Only process if checkbox was checked (value = true)
  const value = e.value;
  if (value !== 'TRUE' && value !== true) return;
  
  const row = e.range.getRow();
  
  // Skip header row (1) and totals row (2) - data starts at row 3
  if (row < 3) return;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get campaign name from column C and campaign ID from column AF (32)
    const campaignName = sheet.getRange(row, 3).getDisplayValue();  // Column C (may be hyperlink)
    const campaignId = sheet.getRange(row, 32).getValue();          // Column AF
    
    // Format for dropdown: "Campaign Name (ID)"
    const dropdownValue = `${campaignName} (${campaignId})`;
    
    // Get Dashboard sheet
    const dashSheet = ss.getSheetByName(CP_DASHBOARD_SHEET_NAME);
    if (!dashSheet) {
      SpreadsheetApp.getUi().alert('Dashboard sheet not found. Please run "Build/Refresh Dashboard" first.');
      // Uncheck the checkbox
      e.range.setValue(false);
      return;
    }
    
    // Set the dropdown value in B3
    dashSheet.getRange('B3').setValue(dropdownValue);
    
    // Uncheck the checkbox
    e.range.setValue(false);
    
    // Navigate to Dashboard sheet
    ss.setActiveSheet(dashSheet);
    dashSheet.getRange('B3').activate();
    
  } catch (error) {
    Logger.log('onEdit error: ' + error.toString());
    // Uncheck the checkbox on error
    try {
      e.range.setValue(false);
    } catch (e2) {
      // Ignore
    }
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Get current date/time in EST timezone
 */
function cpGetESTDate() {
  return new Date(new Date().toLocaleString("en-US", {timeZone: CP_API_TIMEZONE}));
}

/**
 * Format date to DD/MM/YYYY for API
 */
function cpFormatDateForAPI(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

/**
 * Format date to YYYY-MM-DD for display
 */
function cpFormatDateForDisplay(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${year}-${month}-${day}`;
}

/**
 * Get date ranges for Today, Yesterday, and 7-Day periods (EST)
 */
function cpGetDateRanges() {
  const now = cpGetESTDate();
  const today = new Date(now);
  today.setHours(0, 0, 0, 0);
  
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  
  const sevenDaysAgo = new Date(today);
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
  
  return {
    today: {
      start: cpFormatDateForAPI(today),
      end: cpFormatDateForAPI(today),
      label: 'Today'
    },
    yesterday: {
      start: cpFormatDateForAPI(yesterday),
      end: cpFormatDateForAPI(yesterday),
      label: 'Yesterday'
    },
    sevenDay: {
      start: cpFormatDateForAPI(sevenDaysAgo),
      end: cpFormatDateForAPI(yesterday),  // Excludes today
      label: '7-Day'
    }
  };
}

/**
 * Convert value to numeric safely
 */
function cpToNumeric(value, defaultValue = 0) {
  if (value === null || value === undefined || value === '') return defaultValue;
  const num = Number(value);
  return isNaN(num) ? defaultValue : num;
}

/**
 * Get or create sheet by name
 */
function cpGetOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`Created new sheet: ${sheetName}`);
  }
  
  return sheet;
}

/**
 * Extract device and OS from spot name and campaign name
 */
function cpGetDeviceOS(spotName, campaignName) {
  // Device from spot name
  let device = 'Desk';
  if (spotName.includes('Mobile')) device = 'Mob';
  else if (spotName.includes('Tablet')) device = 'Tab';
  
  // OS from campaign name
  let os = 'All';
  const upper = campaignName.toUpperCase();
  if (upper.includes('_IOS_') || upper.includes('_IOS') || upper.includes('-IOS_') || upper.includes('-IOS-')) {
    os = 'iOS';
  } else if (upper.includes('_AND_') || upper.includes('_AND') || upper.includes('-AND_') || upper.includes('-AND-')) {
    os = 'Android';
  }
  
  return `${device} - ${os}`;
}

/**
 * Extract countries from geos object
 */
function cpExtractCountries(geos) {
  if (!geos || typeof geos !== 'object') return '';
  
  const countries = [];
  for (const geoKey in geos) {
    const geo = geos[geoKey];
    if (geo && geo.countryCode && !countries.includes(geo.countryCode)) {
      countries.push(geo.countryCode);
    }
  }
  
  if (countries.length <= 5) {
    return countries.join(', ');
  } else {
    return `${countries.slice(0, 3).join(', ')} (+${countries.length - 3} more)`;
  }
}

/**
 * Extract format from spot name
 * e.g., "Pornhub PC - Preroll" -> "Preroll"
 *       "Pornhub Mobile - Banner 300x250" -> "Banner"
 *       "Xvideos Desktop - Interstitial" -> "Interstitial"
 */
function cpExtractFormat(spotName) {
  if (!spotName || typeof spotName !== 'string') return '';
  
  // Format is typically after the last " - "
  const parts = spotName.split(' - ');
  if (parts.length >= 2) {
    const formatPart = parts[parts.length - 1].trim();
    
    // Clean up format - remove size specs like "300x250"
    // and extract just the format name
    const formatMatch = formatPart.match(/^([A-Za-z]+)/);
    if (formatMatch) {
      return formatMatch[1];
    }
    return formatPart;
  }
  
  return '';
}

// ============================================================================
// API FETCH FUNCTIONS
// ============================================================================

/**
 * DEBUG: Test parallel vs sequential fetching
 * Run this from menu to diagnose issues
 */
function cpDebugParallelFetch() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get first 3 campaign IDs for testing
  const legendSheet = ss.getSheetByName('Legend');
  if (!legendSheet) {
    ui.alert('Error', 'No Legend sheet found', ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = Math.min(legendSheet.getLastRow(), 5);  // Max 4 campaigns for test
  if (lastRow < 2) {
    ui.alert('Error', 'No campaign IDs in Legend', ui.ButtonSet.OK);
    return;
  }
  
  const ids = legendSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const campaignIds = ids.map(row => String(row[0]).trim()).filter(id => id);
  
  Logger.log('=== DEBUG: Parallel vs Sequential Fetch Test ===');
  Logger.log(`Testing with ${campaignIds.length} campaigns: ${campaignIds.join(', ')}`);
  
  // Get today's date for stats
  const dateRanges = cpGetDateRanges();
  const testDate = dateRanges.today.start;
  Logger.log(`Test date: ${testDate}`);
  
  // ========== TEST 1: Sequential Fetch ==========
  Logger.log('\n--- TEST 1: Sequential Fetch ---');
  const seqResults = {};
  
  for (const campaignId of campaignIds) {
    const url = `${CP_API_BASE_URL}/bids/${campaignId}.json?api_key=${CP_API_KEY}&startDate=${testDate}&endDate=${testDate}`;
    Logger.log(`SEQ: Fetching ${campaignId}...`);
    
    try {
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const code = resp.getResponseCode();
      const text = resp.getContentText();
      
      Logger.log(`  Response code: ${code}`);
      Logger.log(`  Response length: ${text.length} chars`);
      
      if (code === 200) {
        const data = JSON.parse(text);
        const bidIds = Object.keys(data);
        Logger.log(`  Bids found: ${bidIds.length}`);
        
        // Check first bid's stats
        if (bidIds.length > 0) {
          const firstBid = data[bidIds[0]];
          const stats = firstBid.stats || {};
          Logger.log(`  First bid (${bidIds[0]}) stats: impressions=${stats.impressions}, revenue=${stats.revenue}, ecpm=${stats.ecpm}`);
          seqResults[campaignId] = { bidCount: bidIds.length, firstBidStats: stats };
        }
      } else {
        Logger.log(`  ERROR: ${text.substring(0, 200)}`);
        seqResults[campaignId] = { error: code };
      }
    } catch (e) {
      Logger.log(`  EXCEPTION: ${e}`);
      seqResults[campaignId] = { exception: e.toString() };
    }
    
    Utilities.sleep(200);
  }
  
  // ========== TEST 2: Parallel Fetch (fetchAll) ==========
  Logger.log('\n--- TEST 2: Parallel Fetch (fetchAll) ---');
  
  const requests = campaignIds.map(id => ({
    url: `${CP_API_BASE_URL}/bids/${id}.json?api_key=${CP_API_KEY}&startDate=${testDate}&endDate=${testDate}`,
    muteHttpExceptions: true
  }));
  
  Logger.log(`PAR: Sending ${requests.length} requests in parallel...`);
  
  let parResults = {};
  try {
    const responses = UrlFetchApp.fetchAll(requests);
    Logger.log(`PAR: Got ${responses.length} responses`);
    
    for (let i = 0; i < campaignIds.length; i++) {
      const campaignId = campaignIds[i];
      const resp = responses[i];
      const code = resp.getResponseCode();
      const text = resp.getContentText();
      
      Logger.log(`\nPAR Campaign ${campaignId}:`);
      Logger.log(`  Response code: ${code}`);
      Logger.log(`  Response length: ${text.length} chars`);
      
      if (code === 200) {
        const data = JSON.parse(text);
        const bidIds = Object.keys(data);
        Logger.log(`  Bids found: ${bidIds.length}`);
        
        if (bidIds.length > 0) {
          const firstBid = data[bidIds[0]];
          const stats = firstBid.stats || {};
          Logger.log(`  First bid (${bidIds[0]}) stats: impressions=${stats.impressions}, revenue=${stats.revenue}, ecpm=${stats.ecpm}`);
          parResults[campaignId] = { bidCount: bidIds.length, firstBidStats: stats };
        }
      } else {
        Logger.log(`  ERROR: ${text.substring(0, 200)}`);
        parResults[campaignId] = { error: code };
      }
    }
  } catch (e) {
    Logger.log(`PAR EXCEPTION: ${e}`);
    parResults = { exception: e.toString() };
  }
  
  // ========== COMPARISON ==========
  Logger.log('\n--- COMPARISON ---');
  for (const campaignId of campaignIds) {
    const seq = seqResults[campaignId] || {};
    const par = parResults[campaignId] || {};
    
    Logger.log(`Campaign ${campaignId}:`);
    Logger.log(`  SEQ: ${JSON.stringify(seq)}`);
    Logger.log(`  PAR: ${JSON.stringify(par)}`);
    
    if (seq.bidCount !== par.bidCount) {
      Logger.log(`  ⚠️ MISMATCH: bid count differs!`);
    }
    if (JSON.stringify(seq.firstBidStats) !== JSON.stringify(par.firstBidStats)) {
      Logger.log(`  ⚠️ MISMATCH: stats differ!`);
    }
  }
  
  Logger.log('\n=== DEBUG COMPLETE ===');
  Logger.log('Check View > Logs for detailed output');
  
  ui.alert('Debug Complete', 
    'Check View > Logs (or Executions) for detailed comparison.\n\n' +
    `Tested ${campaignIds.length} campaigns with date ${testDate}`,
    ui.ButtonSet.OK);
}

/**
 * Fetch current bids from /api/bids/{campaignId}.json
 * OPTIMIZED: Uses small parallel batches with conservative delays to avoid rate limits
 * Returns object keyed by bid_id
 */
function cpFetchCurrentBids(campaignIds) {
  const allBids = {};
  const BATCH_SIZE = 3; // Extra small (3 campaigns = 6 requests) since we fetch campaign+bids per ID
  const MAX_RETRIES = 3;
  
  Logger.log(`Fetching bids for ${campaignIds.length} campaigns in batches of ${BATCH_SIZE}...`);
  
  // Process in batches
  for (let batchStart = 0; batchStart < campaignIds.length; batchStart += BATCH_SIZE) {
    const batchIds = campaignIds.slice(batchStart, batchStart + BATCH_SIZE);
    const batchNum = Math.floor(batchStart/BATCH_SIZE) + 1;
    const totalBatches = Math.ceil(campaignIds.length / BATCH_SIZE);
    Logger.log(`  Batch ${batchNum}/${totalBatches}: ${batchIds.length} campaigns`);
    
    // Track which IDs need retry
    let pendingIds = [...batchIds];
    let retryCount = 0;
    const campaignNames = {};
    
    // Retry loop for this batch
    while (pendingIds.length > 0 && retryCount < MAX_RETRIES) {
      if (retryCount > 0) {
        const backoffMs = Math.pow(2, retryCount) * 2000; // Longer backoff: 4s, 8s, 16s
        Logger.log(`    Retry ${retryCount}/${MAX_RETRIES} for ${pendingIds.length} campaigns after ${backoffMs}ms...`);
        Utilities.sleep(backoffMs);
      }
      
      // Build requests for pending IDs
      const campaignRequests = pendingIds.map(id => ({
        url: `${CP_API_BASE_URL}/campaigns/${id}.json?api_key=${CP_API_KEY}`,
        muteHttpExceptions: true
      }));
      
      const bidRequests = pendingIds.map(id => ({
        url: `${CP_API_BASE_URL}/bids/${id}.json?api_key=${CP_API_KEY}`,
        muteHttpExceptions: true
      }));
      
      const allRequests = [...campaignRequests, ...bidRequests];
      
      try {
        const responses = UrlFetchApp.fetchAll(allRequests);
        const campaignResponses = responses.slice(0, pendingIds.length);
        const bidResponses = responses.slice(pendingIds.length);
        
        const stillPendingIds = [];
        
        for (let i = 0; i < pendingIds.length; i++) {
          const campaignId = pendingIds[i];
          const campaignResp = campaignResponses[i];
          const bidResp = bidResponses[i];
          
          // Check for rate limiting on either request
          if (campaignResp.getResponseCode() === 429 || bidResp.getResponseCode() === 429) {
            stillPendingIds.push(campaignId);
            continue;
          }
          
          // Get campaign name
          if (campaignResp.getResponseCode() === 200) {
            try {
              const data = JSON.parse(campaignResp.getContentText());
              campaignNames[campaignId] = data.campaign_name || '';
            } catch (e) {
              campaignNames[campaignId] = '';
            }
          } else {
            campaignNames[campaignId] = '';
          }
          
          // Get bids
          if (bidResp.getResponseCode() === 200) {
            try {
              const data = JSON.parse(bidResp.getContentText());
              if (typeof data === 'object' && data !== null) {
                for (const bidId of Object.keys(data)) {
                  const bid = data[bidId];
                  if (bid && typeof bid === 'object') {
                    bid.campaign_id = campaignId;
                    bid.campaign_name = campaignNames[campaignId];
                    allBids[bidId] = bid;
                  }
                }
              }
            } catch (e) {
              // Skip parse errors
            }
          }
        }
        
        pendingIds = stillPendingIds;
        
      } catch (e) {
        Logger.log(`    Batch error: ${e}`);
        break;
      }
      
      retryCount++;
    }
    
    if (pendingIds.length > 0) {
      Logger.log(`    Failed ${pendingIds.length} campaigns after retries`);
    }
    
    // Extra long delay for bids (2 requests per campaign = higher rate limit impact)
    if (batchStart + BATCH_SIZE < campaignIds.length) {
      Utilities.sleep(4000); // 4 seconds between batches (6 requests per 4s = ~90 req/min)
    }
  }
  
  Logger.log(`Total bids fetched: ${Object.keys(allBids).length}`);
  return allBids;
}

/**
 * Fetch BID-LEVEL stats for a single date
 * OPTIMIZED: Uses parallel requests with fetchAll + retry logic
 * Uses /api/bids/{campaignId}.json with startDate/endDate params
 * Returns object keyed by bid_id
 */
function cpFetchBidStats(campaignIds, startDate, endDate, periodLabel) {
  Logger.log(`Fetching ${periodLabel} BID stats (${startDate} to ${endDate}) in parallel...`);
  
  const allStats = {};
  const BATCH_SIZE = 5; // Very small batches to stay under rate limit
  const MAX_RETRIES = 3;
  
  // Process in batches
  for (let batchStart = 0; batchStart < campaignIds.length; batchStart += BATCH_SIZE) {
    const batchIds = campaignIds.slice(batchStart, batchStart + BATCH_SIZE);
    const batchNum = Math.floor(batchStart/BATCH_SIZE) + 1;
    const totalBatches = Math.ceil(campaignIds.length / BATCH_SIZE);
    
    // Track which IDs need retry
    let pendingIds = [...batchIds];
    let retryCount = 0;
    
    while (pendingIds.length > 0 && retryCount < MAX_RETRIES) {
      if (retryCount > 0) {
        const backoffMs = Math.pow(2, retryCount) * 2000; // Longer backoff: 4s, 8s, 16s
        Utilities.sleep(backoffMs);
      }
      
      const requests = pendingIds.map(id => ({
        url: `${CP_API_BASE_URL}/bids/${id}.json?api_key=${CP_API_KEY}&startDate=${startDate}&endDate=${endDate}`,
        muteHttpExceptions: true
      }));
      
      try {
        const responses = UrlFetchApp.fetchAll(requests);
        const stillPendingIds = [];
        
        for (let i = 0; i < pendingIds.length; i++) {
          const campaignId = pendingIds[i];
          const resp = responses[i];
          const code = resp.getResponseCode();
          
          if (code === 429) {
            stillPendingIds.push(campaignId);
            continue;
          }
          
          if (code === 200) {
            try {
              const data = JSON.parse(resp.getContentText());
              if (typeof data === 'object' && data !== null) {
                for (const bidId in data) {
                  const bid = data[bidId];
                  const stats = bid.stats || {};
                  allStats[bidId] = {
                    impressions: cpToNumeric(stats.impressions, 0),
                    clicks: cpToNumeric(stats.clicks, 0),
                    conversions: cpToNumeric(stats.conversions, 0),
                    cost: cpToNumeric(stats.revenue, 0),
                    ecpm: cpToNumeric(stats.ecpm, 0),
                    ctr: cpToNumeric(stats.ctr, 0)
                  };
                }
              }
            } catch (e) {
              // Skip parse errors
            }
          }
        }
        
        pendingIds = stillPendingIds;
        
      } catch (e) {
        Logger.log(`Batch error for ${periodLabel}: ${e}`);
        break;
      }
      
      retryCount++;
    }
    
    // Longer delay between batches to stay under rate limit
    if (batchStart + BATCH_SIZE < campaignIds.length) {
      Utilities.sleep(3000); // 3 seconds between batches
    }
  }
  
  Logger.log(`Got ${periodLabel} stats for ${Object.keys(allStats).length} bids`);
  return allStats;
}

/**
 * Fetch BID-LEVEL stats for each day in a 7-day period
 * OPTIMIZED: Uses parallel batching with retry logic for rate limits
 * Returns object keyed by bid_id with totals AND active day count
 */
function cpFetch7DayBidStats(campaignIds) {
  Logger.log('Fetching 7-day bid stats (PARALLEL batching with retry)...');
  
  const now = cpGetESTDate();
  const today = new Date(now);
  today.setHours(0, 0, 0, 0);
  
  // Build array of dates for past 7 days (excluding today)
  const dates = [];
  for (let i = 1; i <= 7; i++) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    dates.push(cpFormatDateForAPI(d));
  }
  
  Logger.log(`Fetching stats for dates: ${dates.join(', ')}`);
  Logger.log(`Total requests needed: ${dates.length} days × ${campaignIds.length} campaigns = ${dates.length * campaignIds.length} requests`);
  
  // Build ALL requests upfront (7 days × N campaigns)
  const allRequestsInfo = [];
  for (const dateStr of dates) {
    for (const campaignId of campaignIds) {
      allRequestsInfo.push({
        date: dateStr,
        campaignId: campaignId,
        url: `${CP_API_BASE_URL}/bids/${campaignId}.json?api_key=${CP_API_KEY}&startDate=${dateStr}&endDate=${dateStr}`
      });
    }
  }
  
  // Process in very small batches with longer delays (350 total requests)
  const BATCH_SIZE = 5; // Tiny batches to stay under rate limit
  const MAX_RETRIES = 3;
  const bidDailyStats = {};
  
  for (let batchStart = 0; batchStart < allRequestsInfo.length; batchStart += BATCH_SIZE) {
    const batchInfo = allRequestsInfo.slice(batchStart, batchStart + BATCH_SIZE);
    const batchNum = Math.floor(batchStart / BATCH_SIZE) + 1;
    const totalBatches = Math.ceil(allRequestsInfo.length / BATCH_SIZE);
    
    Logger.log(`  Batch ${batchNum}/${totalBatches} (${batchInfo.length} requests)...`);
    
    // Track pending requests for retry
    let pendingInfo = [...batchInfo];
    let retryCount = 0;
    
    while (pendingInfo.length > 0 && retryCount < MAX_RETRIES) {
      if (retryCount > 0) {
        const backoffMs = Math.pow(2, retryCount) * 2000; // Longer backoff: 4s, 8s, 16s
        Logger.log(`    Retry ${retryCount}/${MAX_RETRIES} for ${pendingInfo.length} requests after ${backoffMs}ms...`);
        Utilities.sleep(backoffMs);
      }
      
      const requests = pendingInfo.map(info => ({
        url: info.url,
        muteHttpExceptions: true
      }));
      
      try {
        const responses = UrlFetchApp.fetchAll(requests);
        const stillPendingInfo = [];
        
        for (let i = 0; i < pendingInfo.length; i++) {
          const info = pendingInfo[i];
          const resp = responses[i];
          const code = resp.getResponseCode();
          
          if (code === 429) {
            stillPendingInfo.push(info);
            continue;
          }
          
          if (code === 200) {
            try {
              const data = JSON.parse(resp.getContentText());
              if (typeof data === 'object' && data !== null) {
                for (const bidId in data) {
                  const bid = data[bidId];
                  const stats = bid.stats || {};
                  
                  if (!bidDailyStats[bidId]) {
                    bidDailyStats[bidId] = {};
                  }
                  
                  bidDailyStats[bidId][info.date] = {
                    impressions: cpToNumeric(stats.impressions, 0),
                    clicks: cpToNumeric(stats.clicks, 0),
                    conversions: cpToNumeric(stats.conversions, 0),
                    cost: cpToNumeric(stats.revenue, 0),
                    ecpm: cpToNumeric(stats.ecpm, 0),
                    ctr: cpToNumeric(stats.ctr, 0)
                  };
                }
              }
            } catch (e) {
              // Skip parse errors
            }
          }
        }
        
        pendingInfo = stillPendingInfo;
        
      } catch (e) {
        Logger.log(`    Batch ${batchNum} error: ${e}`);
        break;
      }
      
      retryCount++;
    }
    
    if (pendingInfo.length > 0) {
      Logger.log(`    ${pendingInfo.length} requests failed after retries`);
    }
    
    // Longer delay between batches to stay under rate limit (~60 req/min)
    if (batchStart + BATCH_SIZE < allRequestsInfo.length) {
      Utilities.sleep(4000); // 4 seconds between batches for 7-day stats
    }
  }
  
  // Calculate totals and active days per bid
  const result = {};
  
  for (const bidId in bidDailyStats) {
    const dailyStatsMap = bidDailyStats[bidId];
    
    let activeDays = 0;
    let totalImpressions = 0;
    let totalClicks = 0;
    let totalConversions = 0;
    let totalCost = 0;
    let totalEcpmSum = 0;
    let totalCtrSum = 0;
    
    for (const dateStr in dailyStatsMap) {
      const day = dailyStatsMap[dateStr];
      if (day.impressions > 0) {
        activeDays++;
        totalEcpmSum += day.ecpm;
        totalCtrSum += day.ctr;
      }
      totalImpressions += day.impressions;
      totalClicks += day.clicks;
      totalConversions += day.conversions;
      totalCost += day.cost;
    }
    
    const divisor = activeDays > 0 ? activeDays : 1;
    
    result[bidId] = {
      impressions: totalImpressions,
      clicks: totalClicks,
      conversions: totalConversions,
      cost: totalCost,
      avgImpressions: totalImpressions / divisor,
      avgClicks: totalClicks / divisor,
      avgConversions: totalConversions / divisor,
      avgCost: totalCost / divisor,
      ecpm: activeDays > 0 ? totalEcpmSum / activeDays : 0,
      ctr: activeDays > 0 ? totalCtrSum / activeDays : 0,
      activeDays: activeDays
    };
  }
  
  Logger.log(`Processed 7-day stats for ${Object.keys(result).length} bids`);
  return result;
}

/**
 * Fetch CAMPAIGN-LEVEL stats for a single date
 * Uses /api/campaigns/stats.json with startDate/endDate params
 * Returns object keyed by campaign_id
 */
function cpFetchCampaignStatsForDate(campaignIds, dateStr) {
  Logger.log(`Fetching campaign stats for ${dateStr}...`);
  
  const allStats = {};
  
  // Fetch stats for all campaigns with pagination params (matching V9)
  const url = `${CP_API_BASE_URL}/campaigns/stats.json?api_key=${CP_API_KEY}&startDate=${dateStr}&endDate=${dateStr}&limit=500&offset=1`;
  
  try {
    const resp = UrlFetchApp.fetch(url, { 
      method: 'get',
      contentType: 'application/json',
      muteHttpExceptions: true 
    });
    
    if (resp.getResponseCode() !== 200) {
      Logger.log(`Error fetching campaign stats: ${resp.getResponseCode()}`);
      return allStats;
    }
    
    const data = JSON.parse(resp.getContentText());
    
    // Handle both array and object responses (matching V9)
    let campaigns = [];
    if (Array.isArray(data)) {
      campaigns = data;
    } else if (typeof data === 'object' && data !== null) {
      campaigns = Object.values(data);
    }
    
    for (const campaign of campaigns) {
      if (campaign && typeof campaign === 'object') {
        const campId = String(campaign.campaign_id || campaign.campaignId || campaign.id || '');
        
        // Only include campaigns we're tracking
        if (campaignIds.includes(campId)) {
          allStats[campId] = {
            campaignId: campId,
            campaignName: campaign.campaign_name || campaign.campaignName || '',
            impressions: cpToNumeric(campaign.impressions, 0),
            clicks: cpToNumeric(campaign.clicks, 0),
            conversions: cpToNumeric(campaign.conversions, 0),
            cost: cpToNumeric(campaign.cost, 0),  // Campaign-level uses 'cost' not 'revenue'
            ecpm: cpToNumeric(campaign.ecpm || campaign.CPM, 0),
            ctr: cpToNumeric(campaign.ctr || campaign.CTR, 0)
          };
        }
      }
    }
    
  } catch (e) {
    Logger.log(`Error fetching campaign stats: ${e}`);
  }
  
  Logger.log(`Got stats for ${Object.keys(allStats).length} campaigns`);
  return allStats;
}

/**
 * Fetch campaign budget data from /api/campaigns/{campaignId}.json
 * OPTIMIZED: Uses parallel fetching with retry for rate limits (429 errors)
 * Returns object keyed by campaign_id with budget info
 */
function cpFetchCampaignBudgets(campaignIds) {
  Logger.log(`Fetching budgets for ${campaignIds.length} campaigns in parallel...`);
  
  const budgets = {};
  const BATCH_SIZE = 5; // Very small batches to stay under rate limit
  const MAX_RETRIES = 3;
  
  // Process in batches with retry logic
  for (let batchStart = 0; batchStart < campaignIds.length; batchStart += BATCH_SIZE) {
    const batchIds = campaignIds.slice(batchStart, batchStart + BATCH_SIZE);
    const batchNum = Math.floor(batchStart / BATCH_SIZE) + 1;
    
    Logger.log(`  Budget batch ${batchNum}: ${batchIds.length} campaigns`);
    
    // Build requests
    const requests = batchIds.map(id => ({
      url: `${CP_API_BASE_URL}/campaigns/${id}.json?api_key=${CP_API_KEY}`,
      muteHttpExceptions: true
    }));
    
    // Track which IDs need retry
    let pendingIds = [...batchIds];
    let pendingRequests = [...requests];
    let retryCount = 0;
    
    while (pendingIds.length > 0 && retryCount < MAX_RETRIES) {
      if (retryCount > 0) {
        const backoffMs = Math.pow(2, retryCount) * 2000; // Longer backoff: 4s, 8s, 16s // Exponential backoff: 2s, 4s, 8s
        Logger.log(`    Retry ${retryCount}/${MAX_RETRIES} after ${backoffMs}ms delay for ${pendingIds.length} campaigns...`);
        Utilities.sleep(backoffMs);
      }
      
      try {
        const responses = UrlFetchApp.fetchAll(pendingRequests);
        
        const stillPendingIds = [];
        const stillPendingRequests = [];
        
        for (let i = 0; i < pendingIds.length; i++) {
          const campaignId = pendingIds[i];
          const resp = responses[i];
          const code = resp.getResponseCode();
          
          if (code === 200) {
            try {
              const data = JSON.parse(resp.getContentText());
              budgets[campaignId] = {
                campaignId: campaignId,
                campaignName: data.campaign_name || '',
                dailyBudget: cpToNumeric(data.campaign_daily_budget, 0),
                budgetLeft: cpToNumeric(data.daily_budget_left, 0),
                status: data.status || ''
              };
            } catch (e) {
              // Parse error - don't retry
            }
          } else if (code === 429) {
            // Rate limited - add to retry queue
            stillPendingIds.push(campaignId);
            stillPendingRequests.push(pendingRequests[i]);
          } else {
            Logger.log(`    Error fetching ${campaignId}: ${code}`);
          }
        }
        
        pendingIds = stillPendingIds;
        pendingRequests = stillPendingRequests;
        
      } catch (e) {
        Logger.log(`    Batch error: ${e}`);
        break;
      }
      
      retryCount++;
    }
    
    if (pendingIds.length > 0) {
      Logger.log(`    Failed to fetch ${pendingIds.length} campaigns after retries`);
    }
    
    // Longer delay between batches to stay under rate limit
    if (batchStart + BATCH_SIZE < campaignIds.length) {
      Utilities.sleep(3000); // 3 seconds between batches
    }
  }
  
  Logger.log(`Fetched budgets for ${Object.keys(budgets).length}/${campaignIds.length} campaigns`);
  return budgets;
}

/**
 * Update campaign budget via API
 * PUT /api/campaigns/{campaignId}.json with dailyBudget parameter
 */
function cpUpdateCampaignBudget(campaignId, newBudget) {
  const url = `${CP_API_BASE_URL}/campaigns/${campaignId}.json?api_key=${CP_API_KEY}`;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    payload: JSON.stringify({ dailyBudget: newBudget.toString() }),
    muteHttpExceptions: true
  });
  
  return {
    success: response.getResponseCode() === 200,
    responseCode: response.getResponseCode(),
    responseText: response.getContentText()
  };
}

// ============================================================================
// DAILY STATS FUNCTIONS (V2)
// ============================================================================

/**
 * Pull daily stats for the last 7 days (including today)
 * Writes to Daily Stats sheet
 */
function cpPullDailyStats() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Get campaign IDs and names from Legend sheet
    let campaignIds = [];
    const campaignNames = {};  // campId -> name lookup
    const legendSheet = ss.getSheetByName('Legend');
    
    if (legendSheet) {
      const lastRow = legendSheet.getLastRow();
      if (lastRow >= 2) {
        // Get both ID (column A) and potentially name if available
        const legendData = legendSheet.getRange(2, 1, lastRow - 1, 2).getValues();
        for (const row of legendData) {
          const id = String(row[0]).trim();
          if (id) {
            campaignIds.push(id);
            if (row[1]) {
              campaignNames[id] = String(row[1]).trim();
            }
          }
        }
      }
    }
    
    if (campaignIds.length === 0) {
      ui.alert('Error', 'No campaign IDs found in Legend sheet (Column A).', ui.ButtonSet.OK);
      return;
    }
    
    Logger.log(`Pulling daily stats for ${campaignIds.length} campaigns...`);
    
    // Get last 7 days including today
    const now = cpGetESTDate();
    const today = new Date(now);
    today.setHours(0, 0, 0, 0);
    
    const dates = [];
    for (let i = 0; i < 7; i++) {
      const d = new Date(today);
      d.setDate(d.getDate() - i);
      dates.push({
        apiDate: cpFormatDateForAPI(d),
        displayDate: cpFormatDateForDisplay(d),
        dateObj: d
      });
    }
    
    // Reverse so oldest is first
    dates.reverse();
    
    Logger.log(`Dates: ${dates.map(d => d.displayDate).join(', ')}`);
    
    // Fetch stats for each day
    const allDailyStats = [];  // Array of {date, campaignId, campaignName, stats...}
    
    for (const dateInfo of dates) {
      Logger.log(`Fetching ${dateInfo.displayDate}...`);
      
      const dayStats = cpFetchCampaignStatsForDate(campaignIds, dateInfo.apiDate);
      
      // Add entries for each campaign found in API response
      for (const campId in dayStats) {
        const stats = dayStats[campId];
        
        // Calculate CPA
        const cpa = stats.conversions > 0 ? stats.cost / stats.conversions : 0;
        
        // Store campaign name for later use
        if (stats.campaignName) {
          campaignNames[campId] = stats.campaignName;
        }
        
        allDailyStats.push({
          date: dateInfo.displayDate,
          campaignId: stats.campaignId,
          campaignName: stats.campaignName,
          impressions: stats.impressions,
          clicks: stats.clicks,
          conversions: stats.conversions,
          spend: stats.cost,
          cpa: cpa,
          ecpm: stats.ecpm,
          ctr: stats.ctr
        });
      }
      
      // Add 0-data entries for campaigns with no stats that day
      for (const campId of campaignIds) {
        if (!dayStats[campId]) {
          allDailyStats.push({
            date: dateInfo.displayDate,
            campaignId: campId,
            campaignName: campaignNames[campId] || '',
            impressions: 0,
            clicks: 0,
            conversions: 0,
            spend: 0,
            cpa: 0,
            ecpm: 0,
            ctr: 0
          });
        }
      }
      
      Utilities.sleep(200);
    }
    
    // Write to Daily Stats sheet
    cpWriteDailyStats(allDailyStats);
    
    // Count actual data rows (with impressions > 0)
    const dataRows = allDailyStats.filter(s => s.impressions > 0).length;
    
    ui.alert('Success', 
      `Pulled daily stats for ${campaignIds.length} campaign(s) over 7 days.\n\n` +
      `Total entries: ${allDailyStats.length}\n` +
      `Entries with data: ${dataRows}\n` +
      `Date range: ${dates[0].displayDate} to ${dates[dates.length - 1].displayDate}`,
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    ui.alert('Error', `Failed to pull daily stats: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Write daily stats to sheet
 */
function cpWriteDailyStats(dailyStats) {
  const sheet = cpGetOrCreateSheet(CP_DAILY_SHEET_NAME);
  
  // Clear existing data
  sheet.clear();
  
  // Define headers
  const headers = [
    'Date',           // A
    'Campaign ID',    // B
    'Campaign Name',  // C
    'Impressions',    // D
    'Clicks',         // E
    'Conversions',    // F
    'Spend',          // G
    'CPA',            // H
    'eCPM',           // I
    'CTR'             // J
  ];
  
  // Write headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white')
    .setHorizontalAlignment('center');
  
  if (dailyStats.length === 0) {
    Logger.log('No daily stats to write');
    return;
  }
  
  // Sort by date then campaign name
  dailyStats.sort((a, b) => {
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    return a.campaignName.localeCompare(b.campaignName);
  });
  
  // Prepare data rows
  const dataRows = dailyStats.map(stat => [
    stat.date,
    stat.campaignId,
    stat.campaignName,
    stat.impressions,
    stat.clicks,
    stat.conversions,
    stat.spend,
    stat.cpa,
    stat.ecpm,
    stat.ctr
  ]);
  
  // Write data
  sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  
  // Format columns
  const numRows = dataRows.length;
  
  // Date (A) - text
  sheet.getRange(2, 1, numRows, 1).setNumberFormat('yyyy-mm-dd');
  
  // Campaign ID (B) - plain text
  sheet.getRange(2, 2, numRows, 1).setNumberFormat('@');
  
  // Impressions, Clicks, Conversions (D-F) - number
  sheet.getRange(2, 4, numRows, 3).setNumberFormat('#,##0');
  
  // Spend, CPA (G-H) - currency
  sheet.getRange(2, 7, numRows, 2).setNumberFormat('$#,##0.00');
  
  // eCPM (I) - currency
  sheet.getRange(2, 9, numRows, 1).setNumberFormat('$#,##0.000');
  
  // CTR (J) - percentage
  sheet.getRange(2, 10, numRows, 1).setNumberFormat('0.00"%"');
  
  // Alternating row colors
  for (let i = 2; i <= numRows + 1; i++) {
    if (i % 2 === 0) {
      sheet.getRange(i, 1, 1, headers.length).setBackground('#e3f2fd');
    } else {
      sheet.getRange(i, 1, 1, headers.length).setBackground('#ffffff');
    }
  }
  
  // Remove existing filter if present, then add new filter
  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }
  sheet.getRange(1, 1, numRows + 1, headers.length).createFilter();
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  
  // Make Campaign Name wider
  sheet.setColumnWidth(3, 300);
  
  Logger.log(`Wrote ${dataRows.length} rows to ${CP_DAILY_SHEET_NAME}`);
}

// ============================================================================
// DASHBOARD FUNCTIONS (V2)
// ============================================================================

/**
 * Build or refresh the Dashboard sheet with campaign selector and chart
 */
function cpBuildDashboard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if Daily Stats exists
  const dailySheet = ss.getSheetByName(CP_DAILY_SHEET_NAME);
  if (!dailySheet || dailySheet.getLastRow() < 2) {
    ui.alert('Error', 'Please run "Pull Daily Stats" first to get data.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // Get unique campaigns from Daily Stats
    const lastRow = dailySheet.getLastRow();
    const campaignData = dailySheet.getRange(2, 2, lastRow - 1, 2).getValues();
    
    // Build unique campaign list
    const campaigns = {};
    for (const row of campaignData) {
      const id = String(row[0]);
      const name = row[1];
      if (id && !campaigns[id]) {
        campaigns[id] = name;
      }
    }
    
    const campaignList = Object.keys(campaigns).map(id => ({
      id: id,
      name: campaigns[id],
      display: `${campaigns[id]} (${id})`
    }));
    
    Logger.log(`Found ${campaignList.length} unique campaigns`);
    
    // Get or create dashboard sheet
    let dashSheet = ss.getSheetByName(CP_DASHBOARD_SHEET_NAME);
    if (!dashSheet) {
      dashSheet = ss.insertSheet(CP_DASHBOARD_SHEET_NAME);
      Logger.log('Created Dashboard sheet');
    } else {
      dashSheet.clear();
      Logger.log('Cleared existing Dashboard sheet');
    }
    
    // Build dashboard layout
    // Row 1: Title
    dashSheet.getRange('A1').setValue('Campaign Performance Dashboard');
    dashSheet.getRange('A1').setFontSize(18).setFontWeight('bold');
    
    // Row 3: Campaign Selector
    dashSheet.getRange('A3').setValue('Select Campaign:');
    dashSheet.getRange('A3').setFontWeight('bold');
    
    // Create dropdown in B3:D3 (merged cells) with campaign names
    const dropdownValues = campaignList.map(c => c.display);
    const dropdownRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(dropdownValues, true)
      .setAllowInvalid(false)
      .build();
    // Merge B3:D3 for wider dropdown display
    dashSheet.getRange('B3:D3').merge();
    dashSheet.getRange('B3').setDataValidation(dropdownRule);
    dashSheet.getRange('B3').setValue(dropdownValues[0] || '');
    dashSheet.getRange('B3:D3').setBackground('#fff9c4');
    
    // Add TrafficJunky link next to dropdown (E3)
    // Uses the campaign ID extracted to H1 to build dynamic TJ link
    dashSheet.getRange('E3').setFormula('=IF(H1<>"", HYPERLINK("https://advertiser.trafficjunky.com/campaign/"&H1&"/tracking-spots-rules", "🔗 Open in TJ"), "")');
    dashSheet.getRange('E3').setFontColor('#1a73e8');
    
    // Row 5: Header for chart data
    dashSheet.getRange('A5').setValue('Date');
    dashSheet.getRange('B5').setValue('Conv');
    dashSheet.getRange('C5').setValue('Spend');
    dashSheet.getRange('D5').setValue('CPA');
    dashSheet.getRange('E5').setValue('eCPM');
    dashSheet.getRange('F5').setValue('CTR');
    dashSheet.getRange('A5:F5').setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    
    // Add FILTER formulas for each column
    // These will dynamically pull data based on the selected campaign
    // Use TEXT() to convert both sides to strings for reliable comparison
    // Campaign ID is in column B of Daily Stats, dropdown format is "Campaign Name (ID)"
    
    // Extract selected campaign ID to a helper cell (hidden)
    dashSheet.getRange('H1').setFormula(`=REGEXEXTRACT($B$3, "\\(([0-9]+)\\)$")`);
    dashSheet.getRange('H1').setFontColor('#ffffff');  // Hide it (white on white)
    
    // FILTER formulas using TEXT() for type-safe comparison
    // Daily Stats columns: A=Date, B=CampaignID, C=Name, D=Impressions, E=Clicks, F=Conversions, G=Spend, H=CPA, I=eCPM, J=CTR
    dashSheet.getRange('A6').setFormula(`=IFERROR(FILTER('${CP_DAILY_SHEET_NAME}'!A$2:A, TEXT('${CP_DAILY_SHEET_NAME}'!$B$2:$B,"0")=$H$1),"No data")`);
    dashSheet.getRange('B6').setFormula(`=IFERROR(FILTER('${CP_DAILY_SHEET_NAME}'!F$2:F, TEXT('${CP_DAILY_SHEET_NAME}'!$B$2:$B,"0")=$H$1),"No data")`);  // Conversions
    dashSheet.getRange('C6').setFormula(`=IFERROR(FILTER('${CP_DAILY_SHEET_NAME}'!G$2:G, TEXT('${CP_DAILY_SHEET_NAME}'!$B$2:$B,"0")=$H$1),"No data")`);  // Spend
    dashSheet.getRange('D6').setFormula(`=IFERROR(FILTER('${CP_DAILY_SHEET_NAME}'!H$2:H, TEXT('${CP_DAILY_SHEET_NAME}'!$B$2:$B,"0")=$H$1),"No data")`);  // CPA
    dashSheet.getRange('E6').setFormula(`=IFERROR(FILTER('${CP_DAILY_SHEET_NAME}'!I$2:I, TEXT('${CP_DAILY_SHEET_NAME}'!$B$2:$B,"0")=$H$1),"No data")`);  // eCPM
    dashSheet.getRange('F6').setFormula(`=IFERROR(FILTER('${CP_DAILY_SHEET_NAME}'!J$2:J, TEXT('${CP_DAILY_SHEET_NAME}'!$B$2:$B,"0")=$H$1),"No data")`);  // CTR
    
    // Format data columns (extend range to 20 rows for safety)
    dashSheet.getRange('A6:A25').setNumberFormat('yyyy-mm-dd');  // Date format
    dashSheet.getRange('B6:B25').setNumberFormat('#,##0');       // Conversions - whole number
    dashSheet.getRange('C6:C25').setNumberFormat('$#,##0.00');   // Spend
    dashSheet.getRange('D6:D25').setNumberFormat('$#,##0.00');   // CPA
    dashSheet.getRange('E6:E25').setNumberFormat('$#,##0.000');  // eCPM
    dashSheet.getRange('F6:F25').setNumberFormat('0.00"%"');     // CTR
    
    // Set column widths (B-D are merged for selector, so keep them normal width)
    dashSheet.setColumnWidth(1, 115);  // A - Date / Select label
    dashSheet.setColumnWidth(2, 120);  // B - Conv / Part of selector
    dashSheet.setColumnWidth(3, 100);  // C - Spend / Part of selector
    dashSheet.setColumnWidth(4, 80);   // D - CPA / Part of selector
    dashSheet.setColumnWidth(5, 80);   // E - eCPM
    dashSheet.setColumnWidth(6, 80);   // F - CTR
    
    // Create the chart
    // First, remove any existing charts
    const charts = dashSheet.getCharts();
    for (const chart of charts) {
      dashSheet.removeChart(chart);
    }
    
    // Create a combo chart with multiple series
    // Use row 5 as headers for the legend
    // Columns: A=Date, B=Conv, C=Spend, D=CPA, E=eCPM, F=CTR
    const chartBuilder = dashSheet.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(dashSheet.getRange('A5:F12'))  // Header row 5 + 7 days of data (rows 6-12)
      .setNumHeaders(1)  // Row 5 contains headers for legend labels
      .setPosition(3, 8, 0, 0)  // Start at H3 (moved over for new column)
      .setOption('title', 'Campaign Performance Trends')
      .setOption('width', 950)
      .setOption('height', 450)
      .setOption('legend', { position: 'bottom' })
      .setOption('hAxis', { title: 'Date', format: 'yyyy-MM-dd' })
      .setOption('series', {
        0: { targetAxisIndex: 1, type: 'line', color: '#9c27b0', pointSize: 6, dataLabel: 'value' },  // Conv - line, right axis
        1: { targetAxisIndex: 0, type: 'bars', color: '#4285f4', dataLabel: 'value' },  // Spend - bars, left axis
        2: { targetAxisIndex: 0, type: 'line', color: '#ea4335', pointSize: 6, dataLabel: 'value' },  // CPA - line
        3: { targetAxisIndex: 1, type: 'line', color: '#34a853', pointSize: 6, dataLabel: 'value' },  // eCPM - line
        4: { targetAxisIndex: 1, type: 'line', color: '#fbbc04', pointSize: 6, dataLabel: 'value' }   // CTR - line
      })
      .setOption('vAxes', {
        0: { title: 'Spend / CPA ($)', format: '$#,##0.00' },
        1: { title: 'Conv / eCPM / CTR', format: '#,##0.00' }
      })
      .setOption('useFirstColumnAsDomain', true)
      .setOption('annotations', {
        textStyle: { fontSize: 9, color: '#333333' },
        alwaysOutside: false
      });
    
    dashSheet.insertChart(chartBuilder.build());
    
    // Add summary section
    dashSheet.getRange('A14').setValue('Summary Statistics');
    dashSheet.getRange('A14').setFontSize(14).setFontWeight('bold');
    
    // Use TEXT() for type-safe comparison with the helper cell $H$1
    dashSheet.getRange('A15').setValue('Total Spend:');
    dashSheet.getRange('B15').setFormula(`=IFERROR(SUM(FILTER('${CP_DAILY_SHEET_NAME}'!G$2:G, TEXT('${CP_DAILY_SHEET_NAME}'!$B$2:$B,"0")=$H$1)),0)`);
    dashSheet.getRange('B15').setNumberFormat('$#,##0.00');
    
    dashSheet.getRange('A16').setValue('Total Conversions:');
    dashSheet.getRange('B16').setFormula(`=IFERROR(SUM(FILTER('${CP_DAILY_SHEET_NAME}'!F$2:F, TEXT('${CP_DAILY_SHEET_NAME}'!$B$2:$B,"0")=$H$1)),0)`);
    dashSheet.getRange('B16').setNumberFormat('#,##0');
    
    dashSheet.getRange('A17').setValue('Avg Daily Spend:');
    dashSheet.getRange('B17').setFormula(`=IFERROR(AVERAGE(FILTER('${CP_DAILY_SHEET_NAME}'!G$2:G, TEXT('${CP_DAILY_SHEET_NAME}'!$B$2:$B,"0")=$H$1)),0)`);
    dashSheet.getRange('B17').setNumberFormat('$#,##0.00');
    
    dashSheet.getRange('A18').setValue('Avg CPA:');
    dashSheet.getRange('B18').setFormula(`=IFERROR(B15/B16,0)`);
    dashSheet.getRange('B18').setNumberFormat('$#,##0.00');
    
    dashSheet.getRange('A19').setValue('Avg eCPM:');
    dashSheet.getRange('B19').setFormula(`=IFERROR(AVERAGE(FILTER('${CP_DAILY_SHEET_NAME}'!I$2:I, (TEXT('${CP_DAILY_SHEET_NAME}'!$B$2:$B,"0")=$H$1)*('${CP_DAILY_SHEET_NAME}'!I$2:I>0))),0)`);
    dashSheet.getRange('B19').setNumberFormat('$#,##0.000');
    
    dashSheet.getRange('A20').setValue('Avg CTR:');
    dashSheet.getRange('B20').setFormula(`=IFERROR(AVERAGE(FILTER('${CP_DAILY_SHEET_NAME}'!J$2:J, (TEXT('${CP_DAILY_SHEET_NAME}'!$B$2:$B,"0")=$H$1)*('${CP_DAILY_SHEET_NAME}'!J$2:J>0))),0)`);
    dashSheet.getRange('B20').setNumberFormat('0.00"%"');
    
    // Style summary
    dashSheet.getRange('A15:A20').setFontWeight('bold');
    
    // Add Bid Adjustments section (columns D-F)
    dashSheet.getRange('D14').setValue('Bid Adjustments');
    dashSheet.getRange('D14').setFontSize(14).setFontWeight('bold');
    
    // Headers for bid adjustments
    dashSheet.getRange('D15').setValue('Date');
    dashSheet.getRange('E15').setValue('Device');
    dashSheet.getRange('F15').setValue('New Bid');
    dashSheet.getRange('D15:F15').setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    
    // FILTER formula to pull bid adjustments from Bid Logs for this campaign
    // Bid Logs V3: A=Timestamp, B=Campaign ID, F=Spot Name, J=New CPM, M=Status (L is now Comment)
    // Match on Campaign ID (column B) where Status="SUCCESS"
    dashSheet.getRange('D16').setFormula(`=IFERROR(FILTER(TEXT('Bid Logs'!$A$2:$A$1000,"dd/mm/yy"), (TEXT('Bid Logs'!$B$2:$B$1000,"0")=$H$1)*('Bid Logs'!$M$2:$M$1000="SUCCESS")),"No adjustments")`);
    dashSheet.getRange('E16').setFormula(`=IFERROR(FILTER('Bid Logs'!$F$2:$F$1000, (TEXT('Bid Logs'!$B$2:$B$1000,"0")=$H$1)*('Bid Logs'!$M$2:$M$1000="SUCCESS")),"—")`);
    dashSheet.getRange('F16').setFormula(`=IFERROR(FILTER('Bid Logs'!$J$2:$J$1000, (TEXT('Bid Logs'!$B$2:$B$1000,"0")=$H$1)*('Bid Logs'!$M$2:$M$1000="SUCCESS")),"—")`);
    
    // Format bid adjustments
    dashSheet.getRange('F16:F30').setNumberFormat('$#,##0.000');
    
    // Set column widths for bid adjustments section
    dashSheet.setColumnWidth(4, 75);   // D - Date
    dashSheet.setColumnWidth(5, 85);   // E - Device
    dashSheet.setColumnWidth(6, 75);   // F - New Bid
    
    // Add count of adjustments
    dashSheet.getRange('D13').setFormula(`=IFERROR("("&ROWS(FILTER('Bid Logs'!$A$2:$A$1000, (TEXT('Bid Logs'!$B$2:$B$1000,"0")=$H$1)*('Bid Logs'!$M$2:$M$1000="SUCCESS")))&" total)","(0 total)")`);
    dashSheet.getRange('D13').setFontStyle('italic').setFontColor('#666666');
    
    // Instructions
    dashSheet.getRange('A22').setValue('Instructions: Select a campaign from the dropdown above to view its 7-day performance trends.');
    dashSheet.getRange('A22').setFontStyle('italic').setFontColor('#666666');
    
    // Activate dashboard
    ss.setActiveSheet(dashSheet);
    
    ui.alert('Success', 
      `Dashboard created with ${campaignList.length} campaign(s) available in the dropdown.\n\n` +
      'Select a campaign to view its performance trends.',
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    ui.alert('Error', `Failed to build dashboard: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ============================================================================
// MAIN REFRESH FUNCTION
// ============================================================================

/**
 * Main function to refresh the Control Panel
 */
function refreshControlPanel() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startTime = new Date();
  
  // Helper to log elapsed time
  const logTiming = (label, stepStart) => {
    const elapsed = ((new Date() - stepStart) / 1000).toFixed(1);
    const total = ((new Date() - startTime) / 1000).toFixed(1);
    Logger.log(`⏱️ ${label}: ${elapsed}s (total: ${total}s)`);
  };
  
  try {
    // Get campaign IDs from Legend sheet
    let campaignIds = [];
    const legendSheet = ss.getSheetByName('Legend');
    
    if (legendSheet) {
      const lastRow = legendSheet.getLastRow();
      if (lastRow >= 2) {
        const ids = legendSheet.getRange(2, 1, lastRow - 1, 1).getValues();
        campaignIds = ids.map(row => String(row[0]).trim()).filter(id => id && id !== '');
      }
    }
    
    if (campaignIds.length === 0) {
      ui.alert('Error', 'No campaign IDs found in Legend sheet (Column A).', ui.ButtonSet.OK);
      return;
    }
    
    Logger.log(`🚀 OPTIMIZED V5 - Processing ${campaignIds.length} campaigns`);
    Logger.log(`Campaign IDs: ${campaignIds.join(', ')}`);
    
    // Brief initial pause to ensure clean rate limit window
    Logger.log('  ⏸️ Initial 2s pause...');
    Utilities.sleep(2000);
    
    // Get date ranges
    const dateRanges = cpGetDateRanges();
    Logger.log(`Date ranges: Today=${dateRanges.today.start}, Yesterday=${dateRanges.yesterday.start}, 7D=${dateRanges.sevenDay.start} to ${dateRanges.sevenDay.end}`);
    
    // Step 1: Fetch current bids (PARALLEL)
    let stepStart = new Date();
    Logger.log('Step 1: Fetching current bids (PARALLEL)...');
    const currentBids = cpFetchCurrentBids(campaignIds);
    logTiming('Bids fetched', stepStart);
    
    // Cooldown between major operations to let rate limit reset
    Logger.log('  ⏸️ Cooldown 3s before stats...');
    Utilities.sleep(3000);
    
    // Step 2: Fetch BID-LEVEL stats for Today (PARALLEL)
    stepStart = new Date();
    Logger.log('Step 2: Fetching Today stats (PARALLEL)...');
    const todayStats = cpFetchBidStats(campaignIds, dateRanges.today.start, dateRanges.today.end, 'Today');
    logTiming('Today stats', stepStart);
    
    // Cooldown
    Logger.log('  ⏸️ Cooldown 2s...');
    Utilities.sleep(2000);
    
    // Step 3: Fetch BID-LEVEL stats for Yesterday (PARALLEL)
    stepStart = new Date();
    Logger.log('Step 3: Fetching Yesterday stats (PARALLEL)...');
    const yesterdayStats = cpFetchBidStats(campaignIds, dateRanges.yesterday.start, dateRanges.yesterday.end, 'Yesterday');
    logTiming('Yesterday stats', stepStart);
    
    // Longer cooldown before the biggest operation
    Logger.log('  ⏸️ Cooldown 5s before 7-day stats...');
    Utilities.sleep(5000);
    
    // Step 4: Fetch 7-day stats (PARALLEL BATCHED - was the biggest bottleneck!)
    stepStart = new Date();
    Logger.log('Step 4: Fetching 7-day stats (PARALLEL BATCHED)...');
    const sevenDayStats = cpFetch7DayBidStats(campaignIds);
    logTiming('7-day stats', stepStart);
    
    // Cooldown before budgets
    Logger.log('  ⏸️ Cooldown 3s before budgets...');
    Utilities.sleep(3000);
    
    // Step 5: Fetch campaign budgets (PARALLEL with retry)
    stepStart = new Date();
    Logger.log('Step 5: Fetching campaign budgets (PARALLEL)...');
    const campaignBudgets = cpFetchCampaignBudgets(campaignIds);
    logTiming('Budgets fetched', stepStart);
    
    // Step 6: Build Legend lookup
    const legendLookup = {};
    if (legendSheet) {
      const lastRow = legendSheet.getLastRow();
      if (lastRow >= 2) {
        const legendData = legendSheet.getRange(2, 1, lastRow - 1, 6).getValues();
        for (const row of legendData) {
          const campId = String(row[0]).trim();
          if (campId) {
            legendLookup[campId] = {
              strategy: row[2] || '',      // Column C
              subStrategy: row[3] || '',   // Column D
              keyword: row[4] || '',       // Column E
              format: row[5] || ''         // Column F
            };
          }
        }
      }
    }
    
    // Step 7: Merge data into rows
    stepStart = new Date();
    const rows = cpMergeBidData(currentBids, todayStats, yesterdayStats, sevenDayStats, legendLookup, campaignBudgets);
    logTiming('Data merged', stepStart);
    
    // Step 8: Write to sheet (OPTIMIZED - batch operations)
    stepStart = new Date();
    Logger.log('Step 8: Writing to sheet (BATCH OPTIMIZED)...');
    cpWriteToSheet(rows);
    logTiming('Sheet written', stepStart);
    
    // Final timing
    const totalSeconds = ((new Date() - startTime) / 1000).toFixed(1);
    Logger.log(`✅ COMPLETE! Total time: ${totalSeconds}s for ${rows.length} rows`);
    
    ui.alert('Success', 
      `Refreshed Control Panel with ${rows.length} bid entries from ${campaignIds.length} campaign(s).\n\n` +
      `⏱️ Total time: ${totalSeconds} seconds\n\n` +
      `Date ranges:\n` +
      `• Today: ${dateRanges.today.start}\n` +
      `• Yesterday: ${dateRanges.yesterday.start}\n` +
      `• 7-Day: ${dateRanges.sevenDay.start} to ${dateRanges.sevenDay.end}`,
      ui.ButtonSet.OK);
    
  } catch (error) {
    const totalSeconds = ((new Date() - startTime) / 1000).toFixed(1);
    Logger.log(`❌ Error after ${totalSeconds}s: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    ui.alert('Error', `Failed to refresh after ${totalSeconds}s: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Merge all data sources into row objects
 * Stats are now keyed by bid_id for granular per-bid T/Y/7D comparisons
 * 7D stats include active day count for accurate daily averages
 * V3: Now includes campaign budget data
 */
function cpMergeBidData(currentBids, todayStats, yesterdayStats, sevenDayStats, legendLookup, campaignBudgets) {
  const rows = [];
  
  for (const bidId in currentBids) {
    const bid = currentBids[bidId];
    const campaignId = String(bid.campaign_id || '');
    
    // Get BID-LEVEL stats for each period (keyed by bid_id)
    const tStats = todayStats[bidId] || {};
    const yStats = yesterdayStats[bidId] || {};
    const sdStats = sevenDayStats[bidId] || {};
    
    // Get legend data
    const legend = legendLookup[campaignId] || {};
    
    // Get campaign budget (V3)
    const budget = campaignBudgets[campaignId] || {};
    
    // Extract geo info
    const geos = bid.geos || {};
    const geoIds = Object.keys(geos);
    
    // Get 7D averages (already calculated based on active days)
    const sdActiveDays = sdStats.activeDays || 0;
    const sdSpendAvg = sdStats.avgCost || 0;
    const sdConvAvg = sdStats.avgConversions || 0;
    const sdTotalCost = sdStats.cost || 0;
    const sdTotalConv = sdStats.conversions || 0;
    
    // Calculate CPA
    const tConv = tStats.conversions || 0;
    const tSpend = tStats.cost || 0;
    const yConv = yStats.conversions || 0;
    const ySpend = yStats.cost || 0;
    
    const tCpa = tConv > 0 ? tSpend / tConv : 0;
    const yCpa = yConv > 0 ? ySpend / yConv : 0;
    const sdCpa = sdTotalConv > 0 ? sdTotalCost / sdTotalConv : 0;
    
    // Auto-extract format from spot name (e.g., "Pornhub PC - Preroll" -> "Preroll")
    const extractedFormat = cpExtractFormat(bid.spot_name || '');
    
    rows.push({
      tier1Strategy: legend.strategy || '',
      subStrategy: legend.subStrategy || '',
      campaignName: bid.campaign_name || '',
      campaignId: campaignId,
      format: extractedFormat || legend.format || '',  // Auto-extracted, fallback to legend
      country: cpExtractCountries(geos),
      deviceOS: cpGetDeviceOS(bid.spot_name || '', bid.campaign_name || ''),
      spotName: bid.spot_name || '',  // Actual spot name (e.g., "Pornhub PC - Preroll")
      currentBid: cpToNumeric(bid.bid, 0),
      newBid: '',  // User editable
      comment: '',  // User editable - logs with changes
      // V3: Budget columns (removed budgetLeft)
      dailyBudget: budget.dailyBudget || 0,
      newBudget: '',  // User editable
      // Stats columns
      tEcpm: tStats.ecpm || 0,
      yEcpm: yStats.ecpm || 0,
      sdEcpm: sdStats.ecpm || 0,
      tSpend: tSpend,
      ySpend: ySpend,
      sdSpend: sdSpendAvg,
      tCpa: tCpa,
      yCpa: yCpa,
      sdCpa: sdCpa,
      tConv: tConv,
      yConv: yConv,
      sdConv: sdConvAvg,
      tCtr: tStats.ctr || 0,
      yCtr: yStats.ctr || 0,
      sdCtr: sdStats.ctr || 0,
      sdActiveDays: sdActiveDays,
      spotId: bid.spot_id || '',
      bidId: bidId,
      geoId: geoIds.length === 1 ? geoIds[0] : (geoIds.length > 1 ? `${geoIds.length} geos` : ''),
      lastUpdated: new Date()
    });
  }
  
  // Debug: Log stats summary
  const rowsWithTodayData = rows.filter(r => r.tSpend > 0 || r.tConv > 0).length;
  const rowsWithYesterdayData = rows.filter(r => r.ySpend > 0 || r.yConv > 0).length;
  const rowsWithSevenDayData = rows.filter(r => r.sdActiveDays > 0).length;
  Logger.log(`Merge summary: ${rows.length} total rows, ${rowsWithTodayData} with Today data, ${rowsWithYesterdayData} with Yesterday data, ${rowsWithSevenDayData} with 7D data`);
  
  return rows;
}

/**
 * Write data to the Control Panel sheet
 * V3: Now includes budget columns
 */
function cpWriteToSheet(rows) {
  const sheet = cpGetOrCreateSheet(CP_SHEET_NAME);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear existing data
  sheet.clear();
  
  // Define headers (36 columns: A-AJ) - V3 with budget + comment
  const headers = [
    'Tier 1 Strategy',      // A
    'Sub Strategy',         // B
    'Campaign Name',        // C
    'Campaign ID',          // D
    'Format',               // E
    'Country',              // F
    'Device + iOS',         // G
    'Spot Name',            // H - Actual spot name (e.g., "Pornhub PC - Preroll")
    'Current eCPM Bid',     // I
    'New CPM Bid',          // J
    'Change %',             // K - Formula
    'Comment',              // L - User editable, logs with changes
    'T Bid Adjust',         // M - Formula
    'Date last bid Adjust', // N - Formula
    'Daily Budget',         // O - V3: Budget
    'New Budget',           // P - V3: Editable
    'T eCPM',               // Q
    'Y eCPM',               // R
    '7D eCPM',              // S
    'T Spend',              // T
    'Y Spend',              // U
    '7D Spend',             // V (avg per active day)
    'T CPA',                // W
    'Y CPA',                // X
    '7D CPA',               // Y
    'T Conv',               // Z
    'Y Conv',               // AA
    '7D Conv',              // AB (avg per active day)
    'T CTR',                // AC
    'Y CTR',                // AD
    '7D CTR',               // AE
    '7D Active Days',       // AF
    'Spot ID',              // AG
    'Bid ID',               // AH
    'Geo ID',               // AI
    'Last Updated'          // AJ
  ];
  
  // Write headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white')
    .setHorizontalAlignment('center');
  
  if (rows.length === 0) {
    Logger.log('No rows to write');
    return;
  }
  
  // Prepare data rows - V3 with budget + comment + spot name
  const dataRows = rows.map(row => [
    row.tier1Strategy,      // A
    row.subStrategy,        // B
    row.campaignName,       // C
    row.campaignId,         // D
    row.format,             // E
    row.country,            // F
    row.deviceOS,           // G
    row.spotName,           // H - Actual spot name (e.g., "Pornhub PC - Preroll")
    row.currentBid,         // I
    '',                     // J - New CPM Bid (empty for user)
    '',                     // K - Change % (formula)
    '',                     // L - Comment (empty for user)
    '',                     // M - T Bid Adjust (formula)
    '',                     // N - Date last bid Adjust (formula)
    row.dailyBudget,        // O - V3: Daily Budget
    '',                     // P - V3: New Budget (empty for user)
    row.tEcpm,              // Q
    row.yEcpm,              // R
    row.sdEcpm,             // S
    row.tSpend,             // T
    row.ySpend,             // U
    row.sdSpend,            // V (avg per active day)
    row.tCpa,               // W
    row.yCpa,               // X
    row.sdCpa,              // Y
    row.tConv,              // Z
    row.yConv,              // AA
    row.sdConv,             // AB (avg per active day)
    row.tCtr,               // AC
    row.yCtr,               // AD
    row.sdCtr,              // AE
    row.sdActiveDays,       // AF - Active days in 7D period
    row.spotId,             // AG
    row.bidId,              // AH
    row.geoId,              // AI
    row.lastUpdated         // AJ
  ]);
  
  // Write TOTALS row at row 2 (data starts at row 3)
  const lastDataRow = dataRows.length + 2;  // Data goes from row 3 to this row
  const totalsRow = [
    'TOTALS',                                           // A - Strategy
    '',                                                 // B - Sub Strategy
    `${dataRows.length} bids`,                          // C - Count
    '',                                                 // D - Campaign ID
    '',                                                 // E - Format
    '',                                                 // F - Country
    '',                                                 // G - Device
    '',                                                 // H - Spot Name
    '',                                                 // I - Current Bid (average doesn't make sense here)
    '',                                                 // J - New Bid
    '',                                                 // K - Change %
    '',                                                 // L - Comment
    '',                                                 // M - T Bid Adjust
    '',                                                 // N - Date last bid
    `=SUM(O3:O${lastDataRow})`,                         // O - Daily Budget (SUM)
    '',                                                 // P - New Budget
    `=AVERAGE(Q3:Q${lastDataRow})`,                     // Q - T eCPM (AVG)
    `=AVERAGE(R3:R${lastDataRow})`,                     // R - Y eCPM (AVG)
    `=AVERAGE(S3:S${lastDataRow})`,                     // S - 7D eCPM (AVG)
    `=SUM(T3:T${lastDataRow})`,                         // T - T Spend (SUM)
    `=SUM(U3:U${lastDataRow})`,                         // U - Y Spend (SUM)
    `=SUM(V3:V${lastDataRow})`,                         // V - 7D Spend (SUM)
    `=IFERROR(T2/Z2,0)`,                                // W - T CPA (calculated: T Spend / T Conv)
    `=IFERROR(U2/AA2,0)`,                               // X - Y CPA (calculated: Y Spend / Y Conv)
    `=IFERROR((T2+U2)/(Z2+AA2),0)`,                     // Y - 7D CPA (calculated)
    `=SUM(Z3:Z${lastDataRow})`,                         // Z - T Conv (SUM)
    `=SUM(AA3:AA${lastDataRow})`,                       // AA - Y Conv (SUM)
    `=SUM(AB3:AB${lastDataRow})`,                       // AB - 7D Conv (SUM)
    `=AVERAGE(AC3:AC${lastDataRow})`,                   // AC - T CTR (AVG)
    `=AVERAGE(AD3:AD${lastDataRow})`,                   // AD - Y CTR (AVG)
    `=AVERAGE(AE3:AE${lastDataRow})`,                   // AE - 7D CTR (AVG)
    '',                                                 // AF - Active Days
    '',                                                 // AG - Spot ID
    '',                                                 // AH - Bid ID
    '',                                                 // AI - Geo ID
    ''                                                  // AJ - Last Updated
  ];
  sheet.getRange(2, 1, 1, headers.length).setValues([totalsRow]);
  
  // Format totals row
  sheet.getRange(2, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#e8f5e9');  // Light green for totals
  
  // Write data starting at row 3
  sheet.getRange(3, 1, dataRows.length, headers.length).setValues(dataRows);
  
  // OPTIMIZED: Add formulas using batch setFormulas instead of row-by-row
  // Build formula arrays for columns K, M, N (data starts at row 3)
  const formulasK = [];  // Change %
  const formulasM = [];  // T Bid Adjust
  const formulasN = [];  // Date last bid Adjust
  
  for (let i = 3; i <= dataRows.length + 2; i++) {
    // Column K: Change % = (New - Current) / Current * 100
    formulasK.push([`=IF(AND(I${i}>0,J${i}<>""),(J${i}-I${i})/I${i}*100,"")`]);
    
    // Column M: T Bid Adjust - was bid adjusted today?
    formulasM.push([`=IFERROR(IF(ROWS(FILTER('Bid Logs'!$I$2:$I$10000,(TEXT('Bid Logs'!$D$2:$D$10000,"0")=TEXT(AH${i},"0"))*('Bid Logs'!$M$2:$M$10000="SUCCESS")*(INT('Bid Logs'!$A$2:$A$10000)=TODAY())))>0,"Yes ("&TEXT(INDEX(FILTER('Bid Logs'!$I$2:$I$10000,(TEXT('Bid Logs'!$D$2:$D$10000,"0")=TEXT(AH${i},"0"))*('Bid Logs'!$M$2:$M$10000="SUCCESS")*(INT('Bid Logs'!$A$2:$A$10000)=TODAY())),1),"$#,##0.000")&")",""),"")`]);
    
    // Column N: Date last bid Adjust - before today
    formulasN.push([`=IFERROR(LET(data,SORTN(FILTER({'Bid Logs'!$A$2:$A$10000,'Bid Logs'!$J$2:$J$10000},(TEXT('Bid Logs'!$D$2:$D$10000,"0")=TEXT(AH${i},"0"))*('Bid Logs'!$M$2:$M$10000="SUCCESS")*(INT('Bid Logs'!$A$2:$A$10000)<TODAY())),1,0,1,FALSE),TEXT(INDEX(data,1,1),"yyyy-mm-dd")&" ($"&TEXT(INDEX(data,1,2),"#,##0.000")&")"),"Never")`]);
  }
  
  // Write all formulas in batch (data starts at row 3)
  sheet.getRange(3, 11, dataRows.length, 1).setFormulas(formulasK);
  sheet.getRange(3, 13, dataRows.length, 1).setFormulas(formulasM);
  sheet.getRange(3, 14, dataRows.length, 1).setFormulas(formulasN);
  
  // Add TrafficJunky hyperlinks to Campaign Name (column C) - data starts at row 3
  const formulasC = [];
  for (let i = 3; i <= dataRows.length + 2; i++) {
    // HYPERLINK formula that uses Campaign ID from column D to build the TJ URL
    formulasC.push([`=HYPERLINK("https://advertiser.trafficjunky.com/campaign/"&D${i}&"/tracking-spots-rules", "${dataRows[i-3][2].toString().replace(/"/g, '""')}")`]);
  }
  sheet.getRange(3, 3, dataRows.length, 1).setFormulas(formulasC);
  
  // Format columns - V3 with new column positions
  // Data starts at row 3, totals at row 2
  const numRows = dataRows.length;
  
  // Format totals row (row 2) - specific number formats
  sheet.getRange(2, 15, 1, 1).setNumberFormat('$#,##0.00');   // O - Daily Budget
  sheet.getRange(2, 17, 1, 3).setNumberFormat('$#,##0.000');  // Q-S - eCPM
  sheet.getRange(2, 20, 1, 3).setNumberFormat('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');  // T-V - Spend
  sheet.getRange(2, 23, 1, 3).setNumberFormat('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');  // W-Y - CPA
  sheet.getRange(2, 26, 1, 3).setNumberFormat('#,##0');       // Z-AB - Conv
  sheet.getRange(2, 29, 1, 3).setNumberFormat('0.00"%"');     // AC-AE - CTR
  
  // Format data rows (starting at row 3)
  // Spot Name (H) - text
  sheet.getRange(3, 8, numRows, 1).setNumberFormat('@');
  
  // Current eCPM Bid (I) - currency
  sheet.getRange(3, 9, numRows, 1).setNumberFormat('$#,##0.000');
  
  // New CPM Bid (J) - currency (background applied in batch)
  sheet.getRange(3, 10, numRows, 1).setNumberFormat('$#,##0.000');
  
  // Change % (K) - percentage
  sheet.getRange(3, 11, numRows, 1).setNumberFormat('0.00"%"');
  
  // Comment (L) - text (background applied in batch)
  sheet.getRange(3, 12, numRows, 1).setNumberFormat('@');
  
  // T Bid Adjust (M) - conditional formatting
  const mRange = sheet.getRange(3, 13, numRows, 1);
  
  // Date last bid Adjust (N) - text format (contains formula result)
  sheet.getRange(3, 14, numRows, 1).setNumberFormat('@');
  
  // V3: Budget columns (O-P)
  // Daily Budget (O) - currency
  sheet.getRange(3, 15, numRows, 1).setNumberFormat('$#,##0.00');
  // New Budget (P) - currency (background applied in batch)
  sheet.getRange(3, 16, numRows, 1).setNumberFormat('$#,##0.00');
  
  // eCPM columns (Q-S) - currency
  sheet.getRange(3, 17, numRows, 3).setNumberFormat('$#,##0.000');
  
  // Spend columns (T-V) - accounting format
  sheet.getRange(3, 20, numRows, 3).setNumberFormat('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');
  
  // CPA columns (W-Y) - accounting format
  sheet.getRange(3, 23, numRows, 3).setNumberFormat('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');
  
  // Conv columns (Z-AB)
  // T Conv (Z) and Y Conv (AA) - whole numbers
  sheet.getRange(3, 26, numRows, 2).setNumberFormat('#,##0');
  // 7D Conv (AB) - 1 decimal for averages
  sheet.getRange(3, 28, numRows, 1).setNumberFormat('#,##0.0');
  
  // CTR columns (AC-AE) - percentage
  sheet.getRange(3, 29, numRows, 3).setNumberFormat('0.00"%"');
  
  // 7D Active Days (AF) - whole number (background applied in batch)
  sheet.getRange(3, 32, numRows, 1).setNumberFormat('0');
  
  // ID columns (AG-AI) - plain text
  sheet.getRange(3, 33, numRows, 3).setNumberFormat('@');
  
  // Last Updated (AJ) - datetime
  sheet.getRange(3, 36, numRows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  
  // Conditional formatting for T Bid Adjust (M) - green when contains "Yes"
  const rules = sheet.getConditionalFormatRules();
  const yesRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Yes')
    .setBackground('#c8e6c9')
    .setFontColor('#2e7d32')
    .setRanges([mRange])
    .build();
  rules.push(yesRule);
  sheet.setConditionalFormatRules(rules);
  
  // Freeze header row AND totals row, plus columns A-D
  sheet.setFrozenRows(2);  // Freeze rows 1-2 (header + totals)
  sheet.setFrozenColumns(4);  // Freeze columns A-D
  
  // Hide columns A and B (Tier 1 Strategy, Sub Strategy)
  sheet.hideColumns(1, 2);  // Hide columns A-B
  
  // OPTIMIZED: Add alternating row colors using batch setBackgrounds
  // Build a 2D array of background colors for data rows (starting at row 3)
  const backgrounds = [];
  for (let i = 0; i < numRows; i++) {
    const rowNum = i + 3; // Actual row number (1-indexed, starting from row 3)
    const rowColors = [];
    for (let j = 0; j < headers.length; j++) {
      const colNum = j + 1; // Actual column number (1-indexed)
      
      // Special column backgrounds take priority
      if (colNum === 10 || colNum === 12 || colNum === 16) {
        // New CPM Bid (J=10), Comment (L=12), New Budget (P=16) - yellow
        rowColors.push('#fff9c4');
      } else if (colNum === 32) {
        // 7D Active Days (AF=32) - light green
        rowColors.push('#e8f5e9');
      } else if (rowNum % 2 === 1) {
        // Odd rows (3, 5, 7...) - light blue zebra
        rowColors.push('#e3f2fd');
      } else {
        // Even rows (4, 6, 8...) - white
        rowColors.push('#ffffff');
      }
    }
    backgrounds.push(rowColors);
  }
  
  // Apply all backgrounds in one batch operation (data starts at row 3)
  sheet.getRange(3, 1, numRows, headers.length).setBackgrounds(backgrounds);
  
  // Remove existing filter if present, then add new filter
  // Filter includes header (row 1), totals (row 2), and data (row 3+)
  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }
  const dataRange = sheet.getRange(1, 1, numRows + 2, headers.length);  // +2 for header + totals
  dataRange.createFilter();
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  
  // Make Campaign Name wider
  sheet.setColumnWidth(3, 300);
  
  // Activate sheet
  ss.setActiveSheet(sheet);
  
  Logger.log(`Wrote ${dataRows.length} data rows + totals row to ${CP_SHEET_NAME}`);
}

// ============================================================================
// BID MANAGEMENT FUNCTIONS
// ============================================================================

/**
 * Copy current bids to New CPM Bid column
 */
function cpCopyBidsToNew() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CP_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Error', `Sheet "${CP_SHEET_NAME}" not found. Please refresh data first.`, ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('Error', 'No data found. Please refresh data first.', ui.ButtonSet.OK);
    return;
  }
  
  // Copy Current eCPM Bid (I) to New CPM Bid (J) - data starts at row 3 (row 2 is totals)
  const currentBids = sheet.getRange(3, 9, lastRow - 2, 1).getValues();
  sheet.getRange(3, 10, lastRow - 2, 1).setValues(currentBids);
  
  ui.alert('Done', `Copied ${lastRow - 2} bid values to "New CPM Bid" column.`, ui.ButtonSet.OK);
}

/**
 * Calculate and display bid change summary
 */
function cpCalculateBidChanges() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CP_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Error', `Sheet "${CP_SHEET_NAME}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('Error', 'No data found.', ui.ButtonSet.OK);
    return;
  }
  
  // Get data - V3: 36 columns (data starts at row 3, row 2 is totals)
  const data = sheet.getRange(3, 1, lastRow - 2, 36).getValues();
  
  let totalChanges = 0;
  let increased = 0;
  let decreased = 0;
  let unchanged = 0;
  const changedCampaigns = new Set();
  
  for (const row of data) {
    const currentBid = cpToNumeric(row[8], 0);  // I - index 8
    const newBid = cpToNumeric(row[9], 0);      // J - index 9
    const campaignId = row[3];                   // D - index 3
    
    if (newBid > 0) {
      if (newBid > currentBid) {
        increased++;
        changedCampaigns.add(campaignId);
      } else if (newBid < currentBid) {
        decreased++;
        changedCampaigns.add(campaignId);
      } else {
        unchanged++;
      }
      totalChanges++;
    }
  }
  
  const summary = [
    '📊 BID CHANGE SUMMARY',
    '',
    `Total entries with new bids: ${totalChanges}`,
    '',
    `📈 Increased: ${increased}`,
    `📉 Decreased: ${decreased}`,
    `➡️ Unchanged: ${unchanged}`,
    '',
    `Campaigns affected: ${changedCampaigns.size}`
  ].join('\n');
  
  ui.alert('Bid Change Summary', summary, ui.ButtonSet.OK);
}

/**
 * Update bids in TrafficJunky
 */
function cpUpdateBids() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CP_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Error', `Sheet "${CP_SHEET_NAME}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('Error', 'No data found.', ui.ButtonSet.OK);
    return;
  }
  
  // Get all data (36 columns now - V3, data starts at row 3)
  const data = sheet.getRange(3, 1, lastRow - 2, 36).getValues();
  
  // Find bids to update
  const bidsToUpdate = [];
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const currentBid = cpToNumeric(row[8], 0);   // I - Current eCPM Bid
    const newBid = cpToNumeric(row[9], 0);       // J - New CPM Bid
    const comment = String(row[11] || '');       // L - Comment
    const bidId = String(row[33]);               // AH - Bid ID
    const campaignName = row[2];                 // C - Campaign Name
    const campaignId = row[3];                   // D - Campaign ID
    const spotId = row[32];                      // AG - Spot ID
    const spotName = row[7];                     // H - Spot Name (actual, e.g., "Pornhub PC - Preroll")
    const deviceOS = row[6];                     // G - Device + iOS
    const country = row[5];                      // F - Country
    
    // Skip if no bid ID or no new bid or same bid
    if (!bidId || newBid === 0 || newBid === currentBid) continue;
    
    bidsToUpdate.push({
      rowIndex: i + 2,
      campaignId: String(campaignId),
      campaignName: campaignName,
      bidId: bidId,
      spotId: String(spotId),
      spotName: spotName,        // Actual spot name like "Pornhub PC - Preroll"
      deviceOS: deviceOS,        // Device + iOS
      country: country,
      currentBid: currentBid,
      newBid: newBid,
      change: ((newBid - currentBid) / currentBid * 100).toFixed(2),
      comment: comment           // User's comment explaining the change
    });
  }
  
  if (bidsToUpdate.length === 0) {
    ui.alert('No Changes', 'No bids to update. Fill in "New CPM Bid" column with different values.', ui.ButtonSet.OK);
    return;
  }
  
  // Confirmation dialog
  let confirmMsg = `⚠️ CONFIRM BID UPDATES ⚠️\n\n`;
  confirmMsg += `You are about to update ${bidsToUpdate.length} bid(s) in TrafficJunky:\n\n`;
  
  bidsToUpdate.slice(0, 10).forEach((bid, i) => {
    const dir = bid.newBid > bid.currentBid ? '📈' : '📉';
    confirmMsg += `${i + 1}. ${bid.spotName} (${bid.country})\n`;
    confirmMsg += `   $${bid.currentBid.toFixed(3)} → $${bid.newBid.toFixed(3)} (${bid.change}%) ${dir}\n`;
  });
  
  if (bidsToUpdate.length > 10) {
    confirmMsg += `\n... and ${bidsToUpdate.length - 10} more\n`;
  }
  
  confirmMsg += `\nThis will make REAL changes to your TrafficJunky account.\nAre you sure?`;
  
  const confirm = ui.alert('Confirm Bid Updates', confirmMsg, ui.ButtonSet.YES_NO);
  
  if (confirm !== ui.Button.YES) {
    ui.alert('Cancelled', 'No bids were updated.', ui.ButtonSet.OK);
    return;
  }
  
  // Process updates
  Logger.log(`Updating ${bidsToUpdate.length} bids...`);
  
  let successCount = 0;
  let failCount = 0;
  const logEntries = [];
  
  for (const bid of bidsToUpdate) {
    const timestamp = new Date();
    
    try {
      const url = `${CP_API_BASE_URL}/bids/${bid.bidId}/set.json?api_key=${CP_API_KEY}`;
      
      const response = UrlFetchApp.fetch(url, {
        method: 'put',
        contentType: 'application/json',
        payload: JSON.stringify({ bid: bid.newBid.toString() }),
        muteHttpExceptions: true
      });
      
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode === 200) {
        const result = JSON.parse(responseText);
        successCount++;
        
        // Update sheet - V3 column positions
        sheet.getRange(bid.rowIndex, 9).setValue(bid.newBid);  // Update Current eCPM Bid (I)
        sheet.getRange(bid.rowIndex, 10).setValue('');          // Clear New CPM Bid (J)
        sheet.getRange(bid.rowIndex, 12).setValue('');          // Clear Comment (L)
        
        // Log entry with spot name and comment
        logEntries.push({
          timestamp: timestamp,
          campaignId: bid.campaignId,
          campaignName: bid.campaignName,
          bidId: bid.bidId,
          spotId: bid.spotId,
          spotName: bid.spotName,       // Actual spot name like "Pornhub PC - Preroll"
          device: bid.deviceOS,         // Device + iOS
          country: bid.country,
          oldCpm: bid.currentBid,
          newCpm: result.bid || bid.newBid,
          changePercent: bid.change,
          comment: bid.comment,         // User's comment
          status: 'SUCCESS',
          error: ''
        });
        
        Logger.log(`✅ Updated bid ${bid.bidId}: $${bid.currentBid} → $${bid.newBid}`);
      } else {
        failCount++;
        logEntries.push({
          timestamp: timestamp,
          campaignId: bid.campaignId,
          campaignName: bid.campaignName,
          bidId: bid.bidId,
          spotId: bid.spotId,
          spotName: bid.spotName,
          device: bid.deviceOS,
          country: bid.country,
          oldCpm: bid.currentBid,
          newCpm: bid.newBid,
          changePercent: bid.change,
          comment: bid.comment,
          status: 'FAILED',
          error: responseText.substring(0, 200)
        });
        Logger.log(`❌ Failed bid ${bid.bidId}: ${responseCode}`);
      }
      
      Utilities.sleep(200);
      
    } catch (e) {
      failCount++;
      logEntries.push({
        timestamp: timestamp,
        campaignId: bid.campaignId,
        campaignName: bid.campaignName,
        bidId: bid.bidId,
        spotId: bid.spotId,
        spotName: bid.spotName,
        device: bid.deviceOS,
        country: bid.country,
        oldCpm: bid.currentBid,
        newCpm: bid.newBid,
        changePercent: bid.change,
        comment: bid.comment,
        status: 'ERROR',
        error: e.toString().substring(0, 200)
      });
      Logger.log(`❌ Error bid ${bid.bidId}: ${e}`);
    }
  }
  
  // Write to Bid Logs
  if (logEntries.length > 0) {
    cpWriteBidLogs(logEntries);
  }
  
  // Show results
  let resultMsg = `✅ Bid Update Complete!\n\n`;
  resultMsg += `Successful: ${successCount}\n`;
  resultMsg += `Failed: ${failCount}\n\n`;
  resultMsg += `The "Current eCPM Bid" column has been updated.`;
  
  ui.alert('Update Results', resultMsg, ui.ButtonSet.OK);
}

/**
 * Write bid logs (reuses BidManager's Bid Logs sheet)
 * V3: Now includes Comment column
 */
function cpWriteBidLogs(logEntries) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('Bid Logs');
  
  // Create if doesn't exist
  if (!logSheet) {
    logSheet = ss.insertSheet('Bid Logs');
    
    const headers = [
      'Timestamp', 'Campaign ID', 'Campaign Name', 'Bid ID', 'Spot ID',
      'Spot Name', 'Device', 'Country', 'Old CPM', 'New CPM',
      'Change %', 'Comment', 'Status', 'Error'
    ];
    
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    logSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('white');
    logSheet.setFrozenRows(1);
  }
  
  // Prepare rows - V3 includes Comment
  const logRows = logEntries.map(e => [
    e.timestamp, e.campaignId, e.campaignName, e.bidId, e.spotId,
    e.spotName, e.device, e.country, e.oldCpm, e.newCpm,
    e.changePercent, e.comment || '', e.status, e.error
  ]);
  
  // Append
  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow + 1, 1, logRows.length, 14).setValues(logRows);
  
  // Format
  const newStart = lastRow + 1;
  logSheet.getRange(newStart, 1, logRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  logSheet.getRange(newStart, 9, logRows.length, 2).setNumberFormat('$#,##0.000');
  logSheet.getRange(newStart, 11, logRows.length, 1).setNumberFormat('0.00"%"');
  
  Logger.log(`Wrote ${logRows.length} entries to Bid Logs`);
}

/**
 * Clear Control Panel data
 */
function cpClearData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CP_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Info', 'Control Panel sheet not found.', ui.ButtonSet.OK);
    return;
  }
  
  const confirm = ui.alert('Confirm', 'Clear all Control Panel data?', ui.ButtonSet.YES_NO);
  
  if (confirm === ui.Button.YES) {
    sheet.clear();
    ui.alert('Done', 'Control Panel cleared.', ui.ButtonSet.OK);
  }
}

// ============================================================================
// BUDGET MANAGEMENT FUNCTIONS (V3)
// ============================================================================

/**
 * Copy current daily budgets to New Budget column
 */
function cpCopyBudgetsToNew() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CP_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Error', `Sheet "${CP_SHEET_NAME}" not found. Please refresh data first.`, ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('Error', 'No data found. Please refresh data first.', ui.ButtonSet.OK);
    return;
  }
  
  // Copy Daily Budget (O) to New Budget (P) - data starts at row 3
  const currentBudgets = sheet.getRange(3, 15, lastRow - 2, 1).getValues();
  sheet.getRange(3, 16, lastRow - 2, 1).setValues(currentBudgets);
  
  ui.alert('Done', `Copied ${lastRow - 2} budget values to "New Budget" column.`, ui.ButtonSet.OK);
}

/**
 * Calculate and display budget change summary
 */
function cpCalculateBudgetChanges() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CP_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Error', `Sheet "${CP_SHEET_NAME}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('Error', 'No data found.', ui.ButtonSet.OK);
    return;
  }
  
  // Get all data (36 columns now, data starts at row 3)
  const data = sheet.getRange(3, 1, lastRow - 2, 36).getValues();
  
  // Track unique campaigns and their budget changes
  const campaignChanges = {};  // campaignId -> { current, new }
  
  for (const row of data) {
    const campaignId = String(row[3]);           // D - index 3
    const currentBudget = cpToNumeric(row[14], 0);  // O - index 14
    const newBudget = cpToNumeric(row[15], 0);      // P - index 15
    
    // Only track if there's a new budget value
    if (newBudget > 0 && !campaignChanges[campaignId]) {
      campaignChanges[campaignId] = {
        current: currentBudget,
        new: newBudget
      };
    }
  }
  
  // Calculate summary
  let increased = 0;
  let decreased = 0;
  let unchanged = 0;
  
  for (const campId in campaignChanges) {
    const change = campaignChanges[campId];
    if (change.new > change.current) {
      increased++;
    } else if (change.new < change.current) {
      decreased++;
    } else {
      unchanged++;
    }
  }
  
  const totalChanges = Object.keys(campaignChanges).length;
  
  const summary = [
    '💵 BUDGET CHANGE SUMMARY',
    '',
    `Campaigns with new budgets: ${totalChanges}`,
    '',
    `📈 Increased: ${increased}`,
    `📉 Decreased: ${decreased}`,
    `➡️ Unchanged: ${unchanged}`,
    '',
    'Note: Budget changes apply at the campaign level.'
  ].join('\n');
  
  ui.alert('Budget Change Summary', summary, ui.ButtonSet.OK);
}

/**
 * Update campaign budgets in TrafficJunky
 */
function cpUpdateBudgets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CP_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Error', `Sheet "${CP_SHEET_NAME}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('Error', 'No data found.', ui.ButtonSet.OK);
    return;
  }
  
  // Get all data (36 columns, data starts at row 3)
  const data = sheet.getRange(3, 1, lastRow - 2, 36).getValues();
  
  // Find unique campaigns with budget changes
  const budgetsToUpdate = {};  // campaignId -> { campaignName, currentBudget, newBudget, comment, rowIndices }
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const campaignId = String(row[3]);              // D - Campaign ID
    const campaignName = row[2];                    // C - Campaign Name
    const currentBudget = cpToNumeric(row[14], 0);  // O - Daily Budget
    const newBudget = cpToNumeric(row[15], 0);      // P - New Budget
    const comment = String(row[11] || '');          // L - Comment
    
    // Skip if no new budget or same budget
    if (newBudget === 0 || newBudget === currentBudget) continue;
    
    // Only add once per campaign (use first comment found)
    if (!budgetsToUpdate[campaignId]) {
      budgetsToUpdate[campaignId] = {
        campaignName: campaignName,
        currentBudget: currentBudget,
        newBudget: newBudget,
        comment: comment,
        rowIndices: []
      };
    }
    budgetsToUpdate[campaignId].rowIndices.push(i + 2);
  }
  
  const campaignsToUpdate = Object.keys(budgetsToUpdate);
  
  if (campaignsToUpdate.length === 0) {
    ui.alert('No Changes', 'No budgets to update. Fill in "New Budget" column with different values.', ui.ButtonSet.OK);
    return;
  }
  
  // Confirmation dialog
  let confirmMsg = `⚠️ CONFIRM BUDGET UPDATES ⚠️\n\n`;
  confirmMsg += `You are about to update budgets for ${campaignsToUpdate.length} campaign(s) in TrafficJunky:\n\n`;
  
  campaignsToUpdate.slice(0, 10).forEach((campId, i) => {
    const budget = budgetsToUpdate[campId];
    const dir = budget.newBudget > budget.currentBudget ? '📈' : '📉';
    const change = ((budget.newBudget - budget.currentBudget) / budget.currentBudget * 100).toFixed(1);
    confirmMsg += `${i + 1}. ${budget.campaignName}\n`;
    confirmMsg += `   $${budget.currentBudget.toFixed(2)} → $${budget.newBudget.toFixed(2)} (${change}%) ${dir}\n`;
  });
  
  if (campaignsToUpdate.length > 10) {
    confirmMsg += `\n... and ${campaignsToUpdate.length - 10} more\n`;
  }
  
  confirmMsg += `\nThis will make REAL changes to your TrafficJunky account.\nAre you sure?`;
  
  const confirm = ui.alert('Confirm Budget Updates', confirmMsg, ui.ButtonSet.YES_NO);
  
  if (confirm !== ui.Button.YES) {
    ui.alert('Cancelled', 'No budgets were updated.', ui.ButtonSet.OK);
    return;
  }
  
  // Process updates
  Logger.log(`Updating budgets for ${campaignsToUpdate.length} campaigns...`);
  
  let successCount = 0;
  let failCount = 0;
  const logEntries = [];
  
  for (const campaignId of campaignsToUpdate) {
    const budget = budgetsToUpdate[campaignId];
    const timestamp = new Date();
    
    try {
      const result = cpUpdateCampaignBudget(campaignId, budget.newBudget);
      
      if (result.success) {
        successCount++;
        
        // Update all rows for this campaign - V3 column positions
        for (const rowIndex of budget.rowIndices) {
          sheet.getRange(rowIndex, 15).setValue(budget.newBudget);  // Update Daily Budget (O)
          sheet.getRange(rowIndex, 16).setValue('');                 // Clear New Budget (P)
          sheet.getRange(rowIndex, 12).setValue('');                 // Clear Comment (L)
        }
        
        // Log entry with comment
        logEntries.push({
          timestamp: timestamp,
          campaignId: campaignId,
          campaignName: budget.campaignName,
          oldBudget: budget.currentBudget,
          newBudget: budget.newBudget,
          changePercent: ((budget.newBudget - budget.currentBudget) / budget.currentBudget * 100).toFixed(2),
          comment: budget.comment,
          status: 'SUCCESS',
          error: ''
        });
        
        Logger.log(`✅ Updated budget for ${campaignId}: $${budget.currentBudget} → $${budget.newBudget}`);
      } else {
        failCount++;
        logEntries.push({
          timestamp: timestamp,
          campaignId: campaignId,
          campaignName: budget.campaignName,
          oldBudget: budget.currentBudget,
          newBudget: budget.newBudget,
          changePercent: ((budget.newBudget - budget.currentBudget) / budget.currentBudget * 100).toFixed(2),
          comment: budget.comment,
          status: 'FAILED',
          error: result.responseText.substring(0, 200)
        });
        Logger.log(`❌ Failed budget for ${campaignId}: ${result.responseCode}`);
      }
      
      Utilities.sleep(200);
      
    } catch (e) {
      failCount++;
      logEntries.push({
        timestamp: timestamp,
        campaignId: campaignId,
        campaignName: budget.campaignName,
        oldBudget: budget.currentBudget,
        newBudget: budget.newBudget,
        changePercent: ((budget.newBudget - budget.currentBudget) / budget.currentBudget * 100).toFixed(2),
        comment: budget.comment,
        status: 'ERROR',
        error: e.toString().substring(0, 200)
      });
      Logger.log(`❌ Error budget for ${campaignId}: ${e}`);
    }
  }
  
  // Write to Budget Logs
  if (logEntries.length > 0) {
    cpWriteBudgetLogs(logEntries);
  }
  
  // Show results
  let resultMsg = `✅ Budget Update Complete!\n\n`;
  resultMsg += `Successful: ${successCount}\n`;
  resultMsg += `Failed: ${failCount}\n\n`;
  resultMsg += `The "Daily Budget" column has been updated.`;
  
  ui.alert('Update Results', resultMsg, ui.ButtonSet.OK);
}

/**
 * Write budget logs to Budget Logs sheet
 * V3: Now includes Comment column
 */
function cpWriteBudgetLogs(logEntries) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('Budget Logs');
  
  // Create if doesn't exist
  if (!logSheet) {
    logSheet = ss.insertSheet('Budget Logs');
    
    const headers = [
      'Timestamp', 'Campaign ID', 'Campaign Name', 
      'Old Budget', 'New Budget', 'Change %', 
      'Comment', 'Status', 'Error'
    ];
    
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    logSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#34a853')
      .setFontColor('white');
    logSheet.setFrozenRows(1);
  }
  
  // Prepare rows - V3 includes Comment
  const logRows = logEntries.map(e => [
    e.timestamp, e.campaignId, e.campaignName,
    e.oldBudget, e.newBudget, e.changePercent,
    e.comment || '', e.status, e.error
  ]);
  
  // Append
  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow + 1, 1, logRows.length, 9).setValues(logRows);
  
  // Format
  const newStart = lastRow + 1;
  logSheet.getRange(newStart, 1, logRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  logSheet.getRange(newStart, 4, logRows.length, 2).setNumberFormat('$#,##0.00');
  logSheet.getRange(newStart, 6, logRows.length, 1).setNumberFormat('0.00"%"');
  
  Logger.log(`Wrote ${logRows.length} entries to Budget Logs`);
}

// ============================================================================
// PIVOT VIEW FUNCTIONS (V5)
// ============================================================================

/**
 * Build or refresh the Pivot View sheet using NATIVE Google Sheets Pivot Table
 * 
 * Structure:
 * - LEFT SIDE: Native pivot table with Strategy > Sub-Strategy grouping (read-only, auto-subtotals)
 * - RIGHT SIDE: Bid Editor table with all individual bids for editing
 * 
 * The native pivot table handles grouping/subtotals automatically.
 * Edit bids in the Bid Editor section, then use "Update from Pivot" to apply changes.
 */
function cpBuildPivotView() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if Control Panel exists
  const cpSheet = ss.getSheetByName(CP_SHEET_NAME);
  if (!cpSheet || cpSheet.getLastRow() < 3) {
    ui.alert('Error', 'Please refresh Control Panel data first.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // Get data from Control Panel (skip totals row at row 2, data starts at row 3)
    const lastRow = cpSheet.getLastRow();
    const data = cpSheet.getRange(3, 1, lastRow - 2, 36).getValues();
    
    if (data.length === 0) {
      ui.alert('Error', 'No data in Control Panel.', ui.ButtonSet.OK);
      return;
    }
    
    // Get or create pivot sheet
    let pivotSheet = ss.getSheetByName(CP_PIVOT_SHEET_NAME);
    if (!pivotSheet) {
      pivotSheet = ss.insertSheet(CP_PIVOT_SHEET_NAME);
      Logger.log('Created Pivot View sheet');
    } else {
      // Clear existing data and remove all row groups
      pivotSheet.clear();
      try {
        // Remove all existing row groups by collapsing depth
        const maxRows = pivotSheet.getMaxRows();
        if (maxRows > 1) {
          const range = pivotSheet.getRange(2, 1, maxRows - 1, 1);
          // Reset group depth to 0
          while (true) {
            try {
              range.shiftRowGroupDepth(-1);
            } catch (e) {
              break; // No more groups to remove
            }
          }
        }
      } catch (e) {
        // Ignore errors when no groups exist
      }
      Logger.log('Cleared existing Pivot View sheet');
    }
    
    // Build pivot data structure
    // Sort by: Strategy (A), Sub Strategy (B), Campaign Name (C), Device (G)
    const pivotData = data.map((row, idx) => ({
      rowNum: idx + 2,  // Original row in Control Panel
      strategy: row[0] || '(No Strategy)',      // A - Tier 1 Strategy
      subStrategy: row[1] || '(No Sub)',        // B - Sub Strategy
      campaignName: row[2] || '',               // C - Campaign Name
      campaignId: row[3],                       // D - Campaign ID
      format: row[4],                           // E - Format (auto-extracted)
      country: row[5],                          // F - Country
      deviceOS: row[6],                         // G - Device + iOS
      spotName: row[7],                         // H - Spot Name
      currentBid: row[8],                       // I - Current eCPM Bid
      dailyBudget: row[14],                     // O - Daily Budget
      tEcpm: row[16],                           // Q - T eCPM
      yEcpm: row[17],                           // R - Y eCPM
      sdEcpm: row[18],                          // S - 7D eCPM
      tSpend: row[19],                          // T - T Spend
      ySpend: row[20],                          // U - Y Spend
      sdSpend: row[21],                         // V - 7D Spend
      tCpa: row[22],                            // W - T CPA
      yCpa: row[23],                            // X - Y CPA
      sdCpa: row[24],                           // Y - 7D CPA
      tConv: row[25],                           // Z - T Conv
      yConv: row[26],                           // AA - Y Conv
      sdConv: row[27],                          // AB - 7D Conv
      tCtr: row[28],                            // AC - T CTR
      yCtr: row[29],                            // AD - Y CTR
      sdCtr: row[30],                           // AE - 7D CTR
      activeDays: row[31],                      // AF - 7D Active Days
      spotId: row[32],                          // AG - Spot ID
      bidId: row[33],                           // AH - Bid ID
    }));
    
    // Sort by Strategy > Sub Strategy > Campaign Name > Device
    pivotData.sort((a, b) => {
      if (a.strategy !== b.strategy) return a.strategy.localeCompare(b.strategy);
      if (a.subStrategy !== b.subStrategy) return a.subStrategy.localeCompare(b.subStrategy);
      if (a.campaignName !== b.campaignName) return a.campaignName.localeCompare(b.campaignName);
      return a.deviceOS.localeCompare(b.deviceOS);
    });
    
    // Define headers for pivot view
    // Grouped data columns + Edit columns at the end
    const headers = [
      'Strategy',         // A - Group header
      'Sub Strategy',     // B - Group header
      'Campaign Name',    // C - Links to TrafficJunky
      'Format',           // D - Auto-extracted from Spot Name
      'Device + iOS',     // E
      'Spot Name',        // F
      'Country',          // G
      'Current Bid',      // H
      'T Spend',          // I
      'Y Spend',          // J
      '7D Spend',         // K
      'T CPA',            // L
      'Y CPA',            // M
      '7D CPA',           // N
      'T Conv',           // O
      'Y Conv',           // P
      '7D Conv',          // Q
      'T eCPM',           // R
      'Y eCPM',           // S
      '7D eCPM',          // T
      'T CTR',            // U
      'Y CTR',            // V
      '7D CTR',           // W
      'Active',           // X - 7D Active Days
      'Daily Budget',     // Y
      '📊',               // Z - Checkbox to go to Dashboard
      'New CPM',          // AA - EDITABLE
      'New Budget',       // AB - EDITABLE
      'Comment',          // AC - EDITABLE
      '',                 // AD - Spacer
      'Bid ID',           // AE - For reference (can be hidden)
      'Campaign ID',      // AF - For reference (can be hidden)
      'Spot ID'           // AG - For reference (can be hidden)
    ];
    
    // Write headers
    pivotSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    pivotSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1a73e8')
      .setFontColor('white')
      .setHorizontalAlignment('center');
    
    // Prepare data rows WITH SUBTOTALS for each group
    // Structure: Data rows + Sub-Strategy subtotal + Strategy subtotal
    const rows = [];
    const strategyGroups = [];  // { startRow, endRow, name } for row grouping
    const subStrategyGroups = [];
    const strategyHeaderRows = [];  // Track rows where strategy appears (for bold formatting)
    const subStrategyHeaderRows = [];
    const subStrategySubtotalRows = [];  // Track subtotal row numbers for formatting
    const strategySubtotalRows = [];
    const dataRowIndices = [];  // Track which rows are actual data (not subtotals) for hyperlinks/checkboxes
    
    // Helper to create a subtotal row
    // Uses SUMIF/AVERAGEIF to only include DATA rows (where AF column has Campaign ID, not empty)
    // This prevents double-counting when strategy subtotals include sub-strategy subtotal rows in their range
    // Budget also excludes 999999 (unlimited) values
    const createSubtotalRow = (label, sublabel, startRow, endRow, isStrategy) => {
      // AF column contains Campaign ID - empty for subtotal rows, so we use "<>" to only sum data rows
      return {
        isSubtotal: true,
        isStrategy: isStrategy,
        row: [
          isStrategy ? `📊 ${label}` : '',           // A - Strategy subtotal label
          isStrategy ? '' : `  └ ${sublabel}`,       // B - Sub-Strategy subtotal label
          '',                                         // C
          '',                                         // D
          '',                                         // E
          '',                                         // F
          '',                                         // G
          `=AVERAGEIF(AF${startRow}:AF${endRow},"<>",H${startRow}:H${endRow})`,  // H - Current Bid (AVG of data rows only)
          `=SUMIF(AF${startRow}:AF${endRow},"<>",I${startRow}:I${endRow})`,      // I - T Spend (data rows only)
          `=SUMIF(AF${startRow}:AF${endRow},"<>",J${startRow}:J${endRow})`,      // J - Y Spend
          `=SUMIF(AF${startRow}:AF${endRow},"<>",K${startRow}:K${endRow})`,      // K - 7D Spend
          `=IFERROR(I{ROW}/O{ROW},0)`,               // L - T CPA (calculated from this row's sums)
          `=IFERROR(J{ROW}/P{ROW},0)`,               // M - Y CPA
          `=IFERROR(K{ROW}/Q{ROW},0)`,               // N - 7D CPA
          `=SUMIF(AF${startRow}:AF${endRow},"<>",O${startRow}:O${endRow})`,      // O - T Conv
          `=SUMIF(AF${startRow}:AF${endRow},"<>",P${startRow}:P${endRow})`,      // P - Y Conv
          `=SUMIF(AF${startRow}:AF${endRow},"<>",Q${startRow}:Q${endRow})`,      // Q - 7D Conv
          `=AVERAGEIF(AF${startRow}:AF${endRow},"<>",R${startRow}:R${endRow})`,  // R - T eCPM
          `=AVERAGEIF(AF${startRow}:AF${endRow},"<>",S${startRow}:S${endRow})`,  // S - Y eCPM
          `=AVERAGEIF(AF${startRow}:AF${endRow},"<>",T${startRow}:T${endRow})`,  // T - 7D eCPM
          `=AVERAGEIF(AF${startRow}:AF${endRow},"<>",U${startRow}:U${endRow})`,  // U - T CTR
          `=AVERAGEIF(AF${startRow}:AF${endRow},"<>",V${startRow}:V${endRow})`,  // V - Y CTR
          `=AVERAGEIF(AF${startRow}:AF${endRow},"<>",W${startRow}:W${endRow})`,  // W - 7D CTR
          `=AVERAGEIF(AF${startRow}:AF${endRow},"<>",X${startRow}:X${endRow})`,  // X - Active Days
          `=SUMIFS(Y${startRow}:Y${endRow},AF${startRow}:AF${endRow},"<>",Y${startRow}:Y${endRow},"<999999")`, // Y - Budget (data rows, excl unlimited)
          '',                                         // Z - Checkbox
          '',                                         // AA - New CPM
          '',                                         // AB - New Budget
          '',                                         // AC - Comment
          '',                                         // AD - Spacer
          '',                                         // AE - Bid ID
          '',                                         // AF - Campaign ID
          ''                                          // AG - Spot ID
        ]
      };
    };
    
    let currentStrategy = '';
    let currentSubStrategy = '';
    let strategyDataStartRow = 3;  // First data row of current strategy (row 3 = first data after header+totals)
    let subStrategyDataStartRow = 3;  // First data row of current sub-strategy
    let currentSheetRow = 3;  // Tracks actual sheet row as we build (accounts for subtotal rows)
    
    for (let i = 0; i < pivotData.length; i++) {
      const item = pivotData[i];
      const isNewStrategy = item.strategy !== currentStrategy;
      const isNewSubStrategy = isNewStrategy || item.subStrategy !== currentSubStrategy;
      
      // When strategy or sub-strategy changes, insert subtotal rows for the previous group
      if (i > 0 && isNewSubStrategy) {
        // Insert sub-strategy subtotal for the previous sub-strategy
        const subSubtotal = createSubtotalRow(currentStrategy, currentSubStrategy, subStrategyDataStartRow, currentSheetRow - 1, false);
        // Fix the self-referencing formulas
        subSubtotal.row[11] = `=IFERROR(I${currentSheetRow}/O${currentSheetRow},0)`;  // L - T CPA
        subSubtotal.row[12] = `=IFERROR(J${currentSheetRow}/P${currentSheetRow},0)`;  // M - Y CPA
        subSubtotal.row[13] = `=IFERROR(K${currentSheetRow}/Q${currentSheetRow},0)`;  // N - 7D CPA
        rows.push(subSubtotal);
        subStrategySubtotalRows.push(currentSheetRow);
        currentSheetRow++;
      }
      
      if (i > 0 && isNewStrategy) {
        // Insert strategy subtotal for the previous strategy
        const stratSubtotal = createSubtotalRow(currentStrategy, '', strategyDataStartRow, currentSheetRow - 1, true);
        // Fix the self-referencing formulas
        stratSubtotal.row[11] = `=IFERROR(I${currentSheetRow}/O${currentSheetRow},0)`;  // L - T CPA
        stratSubtotal.row[12] = `=IFERROR(J${currentSheetRow}/P${currentSheetRow},0)`;  // M - Y CPA
        stratSubtotal.row[13] = `=IFERROR(K${currentSheetRow}/Q${currentSheetRow},0)`;  // N - 7D CPA
        rows.push(stratSubtotal);
        strategySubtotalRows.push(currentSheetRow);
        currentSheetRow++;
      }
      
      // Update tracking for new groups
      if (isNewStrategy) {
        currentStrategy = item.strategy;
        strategyDataStartRow = currentSheetRow;
        strategyHeaderRows.push(currentSheetRow);
      }
      if (isNewSubStrategy) {
        currentSubStrategy = item.subStrategy;
        subStrategyDataStartRow = currentSheetRow;
        subStrategyHeaderRows.push(currentSheetRow);
      }
      
      // Add the data row
      const showStrategy = isNewStrategy;
      const showSubStrategy = isNewSubStrategy;
      
      rows.push({
        isSubtotal: false,
        dataIndex: i,  // Index into pivotData for hyperlink generation
        row: [
          showStrategy ? item.strategy : '',    // A
          showSubStrategy ? item.subStrategy : '', // B
          item.campaignName,                    // C - Will be replaced with TJ hyperlink
          item.format,                          // D
          item.deviceOS,                        // E
          item.spotName,                        // F
          item.country,                         // G
          item.currentBid,                      // H
          item.tSpend,                          // I
          item.ySpend,                          // J
          item.sdSpend,                         // K - 7D Spend
          item.tCpa,                            // L
          item.yCpa,                            // M
          item.sdCpa,                           // N - 7D CPA
          item.tConv,                           // O
          item.yConv,                           // P
          item.sdConv,                          // Q - 7D Conv
          item.tEcpm,                           // R
          item.yEcpm,                           // S
          item.sdEcpm,                          // T - 7D eCPM
          item.tCtr,                            // U
          item.yCtr,                            // V
          item.sdCtr,                           // W - 7D CTR
          item.activeDays,                      // X - Active Days
          item.dailyBudget,                     // Y
          false,                                // Z - Checkbox for Dashboard navigation
          '',  // AA - New CPM (editable)
          '',  // AB - New Budget (editable)
          '',  // AC - Comment (editable)
          '',  // AD - Spacer
          item.bidId,                           // AE
          item.campaignId,                      // AF
          item.spotId                           // AG
        ]
      });
      dataRowIndices.push({ sheetRow: currentSheetRow, dataIndex: i });
      currentSheetRow++;
    }
    
    // Close final groups - add subtotals for the last strategy/sub-strategy
    if (pivotData.length > 0) {
      // Final sub-strategy subtotal
      const subSubtotal = createSubtotalRow(currentStrategy, currentSubStrategy, subStrategyDataStartRow, currentSheetRow - 1, false);
      subSubtotal.row[11] = `=IFERROR(I${currentSheetRow}/O${currentSheetRow},0)`;
      subSubtotal.row[12] = `=IFERROR(J${currentSheetRow}/P${currentSheetRow},0)`;
      subSubtotal.row[13] = `=IFERROR(K${currentSheetRow}/Q${currentSheetRow},0)`;
      rows.push(subSubtotal);
      subStrategySubtotalRows.push(currentSheetRow);
      currentSheetRow++;
      
      // Final strategy subtotal
      const stratSubtotal = createSubtotalRow(currentStrategy, '', strategyDataStartRow, currentSheetRow - 1, true);
      stratSubtotal.row[11] = `=IFERROR(I${currentSheetRow}/O${currentSheetRow},0)`;
      stratSubtotal.row[12] = `=IFERROR(J${currentSheetRow}/P${currentSheetRow},0)`;
      stratSubtotal.row[13] = `=IFERROR(K${currentSheetRow}/Q${currentSheetRow},0)`;
      rows.push(stratSubtotal);
      strategySubtotalRows.push(currentSheetRow);
      currentSheetRow++;
    }
    
    // Build row grouping info based on the new structure with subtotals
    // We'll create groups after writing data since row numbers are now known
    
    // Extract actual row data from the row objects
    const rowData = rows.map(r => r.row);
    const dataRowCount = dataRowIndices.length;  // Count of actual data rows (not subtotals)
    const lastDataRow = rows.length + 2;  // Last row of data in sheet
    
    // Write TOTALS row at row 2 - uses SUMIF to only sum data rows (exclude subtotals with empty AF)
    // Grand total sums ALL strategy subtotals (which already exclude sub-subtotals)
    const strategySubtotalRowsStr = strategySubtotalRows.join(',');
    const pivotTotalsRow = [
      'TOTALS',                                         // A - Strategy
      '',                                               // B - Sub Strategy  
      `${dataRowCount} bids`,                           // C - Count of actual bids
      '',                                               // D - Format
      '',                                               // E - Device
      '',                                               // F - Spot Name
      '',                                               // G - Country
      `=AVERAGEIF(AF3:AF${lastDataRow},"<>",H3:H${lastDataRow})`,  // H - Current Bid (AVG of data rows)
      `=SUMIF(AF3:AF${lastDataRow},"<>",I3:I${lastDataRow})`,      // I - T Spend (SUM of data rows)
      `=SUMIF(AF3:AF${lastDataRow},"<>",J3:J${lastDataRow})`,      // J - Y Spend
      `=SUMIF(AF3:AF${lastDataRow},"<>",K3:K${lastDataRow})`,      // K - 7D Spend
      `=IFERROR(I2/O2,0)`,                              // L - T CPA (calculated: T Spend / T Conv)
      `=IFERROR(J2/P2,0)`,                              // M - Y CPA (calculated: Y Spend / Y Conv)
      `=IFERROR(K2/Q2,0)`,                              // N - 7D CPA (calculated: 7D Spend / 7D Conv)
      `=SUMIF(AF3:AF${lastDataRow},"<>",O3:O${lastDataRow})`,      // O - T Conv
      `=SUMIF(AF3:AF${lastDataRow},"<>",P3:P${lastDataRow})`,      // P - Y Conv
      `=SUMIF(AF3:AF${lastDataRow},"<>",Q3:Q${lastDataRow})`,      // Q - 7D Conv
      `=AVERAGEIF(AF3:AF${lastDataRow},"<>",R3:R${lastDataRow})`,  // R - T eCPM (AVG)
      `=AVERAGEIF(AF3:AF${lastDataRow},"<>",S3:S${lastDataRow})`,  // S - Y eCPM
      `=AVERAGEIF(AF3:AF${lastDataRow},"<>",T3:T${lastDataRow})`,  // T - 7D eCPM
      `=AVERAGEIF(AF3:AF${lastDataRow},"<>",U3:U${lastDataRow})`,  // U - T CTR
      `=AVERAGEIF(AF3:AF${lastDataRow},"<>",V3:V${lastDataRow})`,  // V - Y CTR
      `=AVERAGEIF(AF3:AF${lastDataRow},"<>",W3:W${lastDataRow})`,  // W - 7D CTR
      `=AVERAGEIF(AF3:AF${lastDataRow},"<>",X3:X${lastDataRow})`,  // X - Active Days
      `=SUMIFS(Y3:Y${lastDataRow},AF3:AF${lastDataRow},"<>",Y3:Y${lastDataRow},"<999999")`,  // Y - Budget (exclude unlimited)
      '',                                               // Z - Checkbox (empty for totals)
      '',                                               // AA - New CPM
      '',                                               // AB - New Budget
      '',                                               // AC - Comment
      '',                                               // AD - Spacer
      '',                                               // AE - Bid ID
      '',                                               // AF - Campaign ID
      ''                                                // AG - Spot ID
    ];
    pivotSheet.getRange(2, 1, 1, headers.length).setValues([pivotTotalsRow]);
    
    // Format totals row
    pivotSheet.getRange(2, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#c8e6c9');  // Green for grand totals
    
    // Write data starting at row 3
    if (rowData.length > 0) {
      pivotSheet.getRange(3, 1, rowData.length, headers.length).setValues(rowData);
    }
    
    Logger.log('Adding hyperlink formulas and checkboxes for data rows only...');
    
    // Add TrafficJunky hyperlinks ONLY to data rows (not subtotals)
    // Also add checkboxes only to data rows
    for (const rowInfo of dataRowIndices) {
      const sheetRow = rowInfo.sheetRow;
      const item = pivotData[rowInfo.dataIndex];
      
      // Column C: Campaign Name → TrafficJunky link (Campaign ID in AF)
      const escapedName = item.campaignName.toString().replace(/"/g, '""');
      pivotSheet.getRange(sheetRow, 3).setFormula(`=HYPERLINK("https://advertiser.trafficjunky.com/campaign/"&AF${sheetRow}&"/tracking-spots-rules", "${escapedName}")`);
    }
    
    // Column Z (26): Add checkbox data validation ONLY to data rows
    const checkboxRule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .setAllowInvalid(false)
      .build();
    for (const rowInfo of dataRowIndices) {
      pivotSheet.getRange(rowInfo.sheetRow, 26).setDataValidation(checkboxRule);
    }
    
    // Format subtotal rows
    Logger.log(`Formatting ${subStrategySubtotalRows.length} sub-strategy subtotals and ${strategySubtotalRows.length} strategy subtotals`);
    
    if (subStrategySubtotalRows.length > 0) {
      const subRanges = subStrategySubtotalRows.map(r => `A${r}:AG${r}`);
      pivotSheet.getRangeList(subRanges)
        .setFontWeight('bold')
        .setFontStyle('italic')
        .setBackground('#e3f2fd');  // Light blue for sub-strategy subtotals
    }
    
    if (strategySubtotalRows.length > 0) {
      const stratRanges = strategySubtotalRows.map(r => `A${r}:AG${r}`);
      pivotSheet.getRangeList(stratRanges)
        .setFontWeight('bold')
        .setBackground('#bbdefb');  // Darker blue for strategy subtotals
    }
    
    Logger.log('Applying formatting (batch operations)...');
    const numRows = rowData.length;
    
    // Format totals row (row 2)
    pivotSheet.getRange(2, 8, 1, 1).setNumberFormat('$#,##0.000');   // H - Current Bid
    pivotSheet.getRange(2, 9, 1, 3).setNumberFormat('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');  // I-K Spend
    pivotSheet.getRange(2, 12, 1, 3).setNumberFormat('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'); // L-N CPA
    pivotSheet.getRange(2, 15, 1, 3).setNumberFormat('#,##0');       // O-Q Conv
    pivotSheet.getRange(2, 18, 1, 3).setNumberFormat('$#,##0.000');  // R-T eCPM
    pivotSheet.getRange(2, 21, 1, 3).setNumberFormat('0.00"%"');     // U-W CTR
    pivotSheet.getRange(2, 24, 1, 1).setNumberFormat('0.0');         // X - Active Days
    pivotSheet.getRange(2, 25, 1, 1).setNumberFormat('$#,##0.00');   // Y - Budget
    
    // BATCH ALL NUMBER FORMATTING for all rows (starting at row 3)
    // Current Bid (H) - currency
    pivotSheet.getRange(3, 8, numRows, 1).setNumberFormat('$#,##0.000');
    // Spend columns (I-K) - accounting
    pivotSheet.getRange(3, 9, numRows, 3).setNumberFormat('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');
    // CPA columns (L-N) - accounting
    pivotSheet.getRange(3, 12, numRows, 3).setNumberFormat('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');
    // Conv columns (O-Q) - number
    pivotSheet.getRange(3, 15, numRows, 3).setNumberFormat('#,##0');
    // eCPM columns (R-T) - currency
    pivotSheet.getRange(3, 18, numRows, 3).setNumberFormat('$#,##0.000');
    // CTR columns (U-W) - percentage
    pivotSheet.getRange(3, 21, numRows, 3).setNumberFormat('0.00"%"');
    // Active Days (X) - number
    pivotSheet.getRange(3, 24, numRows, 1).setNumberFormat('0');
    // Daily Budget (Y) - currency
    pivotSheet.getRange(3, 25, numRows, 1).setNumberFormat('$#,##0.00');
    // NEW CPM (AA = 27) - currency
    pivotSheet.getRange(3, 27, numRows, 1).setNumberFormat('$#,##0.000');
    // New Budget (AB = 28) - currency
    pivotSheet.getRange(3, 28, numRows, 1).setNumberFormat('$#,##0.00');
    // ID columns (AE-AG = 31-33) - plain text
    pivotSheet.getRange(3, 31, numRows, 3).setNumberFormat('@').setFontColor('#999999');
    
    // BATCH BACKGROUND COLORS for data rows only (skip subtotal rows - they have their own formatting)
    const editableRanges = [];
    const dataRowBgRanges = [];
    
    // Collect editable column ranges for data rows only (subtotal rows shouldn't be editable)
    for (const rowInfo of dataRowIndices) {
      const r = rowInfo.sheetRow;
      // Editable columns (AA, AB, AC = 27-29) - yellow for data rows
      editableRanges.push(`AA${r}:AC${r}`);
      // Alternating row colors for data rows
      if (r % 2 === 1) {
        dataRowBgRanges.push(`A${r}:Z${r}`);
        dataRowBgRanges.push(`AD${r}:AG${r}`);
      }
    }
    
    // Apply editable column background (yellow) to data rows only
    if (editableRanges.length > 0) {
      pivotSheet.getRangeList(editableRanges).setBackground('#fff9c4');
    }
    
    // Apply alternating row colors to data rows only
    if (dataRowBgRanges.length > 0) {
      pivotSheet.getRangeList(dataRowBgRanges).setBackground('#fafafa');
    }
    
    // Format strategy/sub-strategy header rows (first data row of each group)
    if (strategyHeaderRows.length > 0) {
      const strategyRanges = strategyHeaderRows.map(r => `A${r}`);
      pivotSheet.getRangeList(strategyRanges)
        .setFontWeight('bold');
    }
    
    if (subStrategyHeaderRows.length > 0) {
      const subStrategyRanges = subStrategyHeaderRows.map(r => `B${r}`);
      pivotSheet.getRangeList(subStrategyRanges)
        .setFontWeight('bold')
        .setFontStyle('italic');
    }
    
    // CREATE ROW GROUPS for collapsible pivot-table-like behavior
    // Structure: When collapsed, subtotal rows remain visible
    // - Collapsing sub-strategy: hides data rows, shows sub-strategy subtotal
    // - Collapsing strategy: hides data rows AND sub-strategy subtotals, shows strategy subtotal
    Logger.log('Creating row groups for collapsible sections...');
    
    // Scan the rows array to identify group boundaries
    // Each sub-strategy subtotal marks the END of a sub-strategy group
    // Each strategy subtotal marks the END of a strategy group
    const subStratGroupRanges = [];  // { startRow, endRow } - data rows to group under sub-strategy subtotal
    const stratGroupRanges = [];     // { startRow, endRow } - all rows (data + sub-subtotals) under strategy subtotal
    
    let currentSubStratDataStart = -1;
    let currentStratStart = -1;
    
    for (let i = 0; i < rows.length; i++) {
      const sheetRow = i + 3;  // Sheet row (1=header, 2=grand total, 3+=data)
      const rowObj = rows[i];
      
      if (rowObj.isSubtotal) {
        if (rowObj.isStrategy) {
          // Strategy subtotal - close the strategy group
          // Group includes everything from start to the row BEFORE this subtotal
          if (currentStratStart > 0 && sheetRow - 1 >= currentStratStart) {
            stratGroupRanges.push({
              startRow: currentStratStart,
              endRow: sheetRow - 1
            });
          }
          currentStratStart = -1;
        } else {
          // Sub-strategy subtotal - close the sub-strategy data group
          // Group includes data rows up to (not including) this subtotal
          if (currentSubStratDataStart > 0 && sheetRow - 1 >= currentSubStratDataStart) {
            subStratGroupRanges.push({
              startRow: currentSubStratDataStart,
              endRow: sheetRow - 1
            });
          }
          currentSubStratDataStart = -1;
        }
      } else {
        // Data row - start new groups if needed
        if (currentSubStratDataStart < 0) {
          currentSubStratDataStart = sheetRow;
        }
        if (currentStratStart < 0) {
          currentStratStart = sheetRow;
        }
      }
    }
    
    // Create STRATEGY groups FIRST (outer level - will be depth 1)
    Logger.log(`Creating ${stratGroupRanges.length} strategy groups (outer level)`);
    for (const group of stratGroupRanges) {
      try {
        const range = pivotSheet.getRange(group.startRow, 1, group.endRow - group.startRow + 1, 1);
        range.shiftRowGroupDepth(1);
        Logger.log(`  Strategy group: rows ${group.startRow}-${group.endRow}`);
      } catch (e) {
        Logger.log(`  Error creating strategy group (rows ${group.startRow}-${group.endRow}): ${e}`);
      }
    }
    
    // Create SUB-STRATEGY groups SECOND (inner level - will be depth 2 where they overlap)
    Logger.log(`Creating ${subStratGroupRanges.length} sub-strategy groups (inner level)`);
    for (const group of subStratGroupRanges) {
      try {
        const range = pivotSheet.getRange(group.startRow, 1, group.endRow - group.startRow + 1, 1);
        range.shiftRowGroupDepth(1);
        Logger.log(`  Sub-strategy group: rows ${group.startRow}-${group.endRow}`);
      } catch (e) {
        Logger.log(`  Error creating sub-strategy group (rows ${group.startRow}-${group.endRow}): ${e}`);
      }
    }
    
    // Freeze header row + totals row, and first 2 columns
    pivotSheet.setFrozenRows(2);  // Freeze rows 1-2 (header + totals)
    pivotSheet.setFrozenColumns(2);
    
    // Auto-resize columns
    pivotSheet.autoResizeColumns(1, headers.length);
    
    // Set specific widths
    pivotSheet.setColumnWidth(1, 150);   // A - Strategy
    pivotSheet.setColumnWidth(2, 150);   // B - Sub Strategy
    pivotSheet.setColumnWidth(3, 250);   // C - Campaign Name
    pivotSheet.setColumnWidth(4, 80);    // D - Format
    pivotSheet.setColumnWidth(6, 180);   // F - Spot Name
    pivotSheet.setColumnWidth(24, 50);   // X - Active Days
    pivotSheet.setColumnWidth(26, 40);   // Z - Dashboard checkbox
    pivotSheet.setColumnWidth(30, 20);   // AD - Spacer
    
    // Hide ID columns (AE-AG = 31-33) for cleaner view
    pivotSheet.hideColumns(31, 3);
    
    // Remove existing filter if present, then add new filter
    const existingFilter = pivotSheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }
    pivotSheet.getRange(1, 1, numRows + 2, headers.length).createFilter();  // +2 for header + totals
    
    // Activate pivot sheet
    ss.setActiveSheet(pivotSheet);
    
    ui.alert('Success', 
      `Pivot View created with:\n` +
      `• ${dataRowCount} bid rows\n` +
      `• ${subStrategySubtotalRows.length} sub-strategy subtotals\n` +
      `• ${strategySubtotalRows.length} strategy subtotals\n` +
      `• 1 grand total row\n\n` +
      'Use the +/- buttons on the LEFT to collapse/expand groups like a pivot table.\n' +
      'Subtotals: SUM for Spend/Conv/Budget (excl. unlimited), AVG for Bid/eCPM/CTR.',
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    ui.alert('Error', `Failed to build Pivot View: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Copy current bids to New CPM column in Pivot View
 */
function cpCopyBidsPivot() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CP_PIVOT_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Error', `Sheet "${CP_PIVOT_SHEET_NAME}" not found. Please build Pivot View first.`, ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('Error', 'No data found in Pivot View.', ui.ButtonSet.OK);
    return;
  }
  
  // Copy Current Bid (H=8) to New CPM (AA=27) - data starts at row 3 (row 2 is totals)
  const currentBids = sheet.getRange(3, 8, lastRow - 2, 1).getValues();
  sheet.getRange(3, 27, lastRow - 2, 1).setValues(currentBids);
  
  ui.alert('Done', `Copied ${lastRow - 2} bid values to "New CPM" column.`, ui.ButtonSet.OK);
}

/**
 * Copy current budgets to New Budget column in Pivot View
 */
function cpCopyBudgetsPivot() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CP_PIVOT_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Error', `Sheet "${CP_PIVOT_SHEET_NAME}" not found. Please build Pivot View first.`, ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('Error', 'No data found in Pivot View.', ui.ButtonSet.OK);
    return;
  }
  
  // Copy Daily Budget (Y=25) to New Budget (AB=28) - data starts at row 3 (row 2 is totals)
  const currentBudgets = sheet.getRange(3, 25, lastRow - 2, 1).getValues();
  sheet.getRange(3, 28, lastRow - 2, 1).setValues(currentBudgets);
  
  ui.alert('Done', `Copied ${lastRow - 2} budget values to "New Budget" column.`, ui.ButtonSet.OK);
}

/**
 * Update bids and budgets from Pivot View
 * Reads Bid ID and Campaign ID from the same row as the edit values
 */
function cpUpdateFromPivot() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CP_PIVOT_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Error', `Sheet "${CP_PIVOT_SHEET_NAME}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('Error', 'No data found.', ui.ButtonSet.OK);
    return;
  }
  
  // Get all data from pivot view (33 columns with all stats, skip totals row at row 2)
  const data = sheet.getRange(3, 1, lastRow - 2, 33).getValues();
  
  // Collect bid and budget updates
  const bidUpdates = [];
  const budgetUpdates = {};  // campaignId -> { newBudget, comment, rowIndices }
  
  // Column indices (0-based):
  // H=7 Current Bid, Y=24 Daily Budget, AA=26 New CPM, AB=27 New Budget, AC=28 Comment
  // AE=30 Bid ID, AF=31 Campaign ID, AG=32 Spot ID
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const currentBid = cpToNumeric(row[7], 0);    // H - Current Bid (index 7)
    const newCpm = cpToNumeric(row[26], 0);       // AA - New CPM (index 26)
    const newBudget = cpToNumeric(row[27], 0);    // AB - New Budget (index 27)
    const comment = String(row[28] || '');        // AC - Comment (index 28)
    const bidId = String(row[30] || '');          // AE - Bid ID (index 30)
    const campaignId = String(row[31] || '');     // AF - Campaign ID (index 31)
    const spotId = String(row[32] || '');         // AG - Spot ID (index 32)
    const campaignName = row[2];                  // C - Campaign Name (index 2)
    const spotName = row[5];                      // F - Spot Name (index 5)
    const deviceOS = row[4];                      // E - Device + iOS (index 4)
    const country = row[6];                       // G - Country (index 6)
    const dailyBudget = cpToNumeric(row[24], 0);  // Y - Daily Budget (index 24)
    
    // Check for bid update
    if (bidId && newCpm > 0 && newCpm !== currentBid) {
      bidUpdates.push({
        rowIndex: i + 2,
        bidId: bidId,
        campaignId: campaignId,
        campaignName: campaignName,
        spotId: spotId,
        spotName: spotName,
        deviceOS: deviceOS,
        country: country,
        currentBid: currentBid,
        newBid: newCpm,
        change: ((newCpm - currentBid) / currentBid * 100).toFixed(2),
        comment: comment
      });
    }
    
    // Check for budget update (aggregate by campaign)
    if (campaignId && newBudget > 0 && newBudget !== dailyBudget) {
      if (!budgetUpdates[campaignId]) {
        budgetUpdates[campaignId] = {
          campaignName: campaignName,
          currentBudget: dailyBudget,
          newBudget: newBudget,
          comment: comment,
          rowIndices: []
        };
      }
      budgetUpdates[campaignId].rowIndices.push(i + 2);
    }
  }
  
  const budgetCampaigns = Object.keys(budgetUpdates);
  
  if (bidUpdates.length === 0 && budgetCampaigns.length === 0) {
    ui.alert('No Changes', 'No bids or budgets to update. Fill in "New CPM" or "New Budget" columns with different values.', ui.ButtonSet.OK);
    return;
  }
  
  // Confirmation dialog
  let confirmMsg = `⚠️ CONFIRM UPDATES FROM PIVOT ⚠️\n\n`;
  
  if (bidUpdates.length > 0) {
    confirmMsg += `BID UPDATES: ${bidUpdates.length}\n`;
    bidUpdates.slice(0, 5).forEach((bid, i) => {
      const dir = bid.newBid > bid.currentBid ? '📈' : '📉';
      confirmMsg += `  ${i + 1}. ${bid.spotName}: $${bid.currentBid.toFixed(3)} → $${bid.newBid.toFixed(3)} ${dir}\n`;
    });
    if (bidUpdates.length > 5) {
      confirmMsg += `  ... and ${bidUpdates.length - 5} more\n`;
    }
    confirmMsg += '\n';
  }
  
  if (budgetCampaigns.length > 0) {
    confirmMsg += `BUDGET UPDATES: ${budgetCampaigns.length} campaign(s)\n`;
    budgetCampaigns.slice(0, 5).forEach((campId, i) => {
      const budget = budgetUpdates[campId];
      const dir = budget.newBudget > budget.currentBudget ? '📈' : '📉';
      confirmMsg += `  ${i + 1}. ${budget.campaignName}: $${budget.currentBudget.toFixed(2)} → $${budget.newBudget.toFixed(2)} ${dir}\n`;
    });
    if (budgetCampaigns.length > 5) {
      confirmMsg += `  ... and ${budgetCampaigns.length - 5} more\n`;
    }
  }
  
  confirmMsg += `\nThis will make REAL changes to your TrafficJunky account.\nAre you sure?`;
  
  const confirm = ui.alert('Confirm Updates', confirmMsg, ui.ButtonSet.YES_NO);
  
  if (confirm !== ui.Button.YES) {
    ui.alert('Cancelled', 'No changes were made.', ui.ButtonSet.OK);
    return;
  }
  
  // Process bid updates
  let bidSuccess = 0;
  let bidFail = 0;
  const bidLogEntries = [];
  
  for (const bid of bidUpdates) {
    const timestamp = new Date();
    
    try {
      const url = `${CP_API_BASE_URL}/bids/${bid.bidId}/set.json?api_key=${CP_API_KEY}`;
      
      const response = UrlFetchApp.fetch(url, {
        method: 'put',
        contentType: 'application/json',
        payload: JSON.stringify({ bid: bid.newBid.toString() }),
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() === 200) {
        const result = JSON.parse(response.getContentText());
        bidSuccess++;
        
        // Update pivot sheet (column positions with Format column)
        sheet.getRange(bid.rowIndex, 8).setValue(bid.newBid);   // Update Current Bid (H)
        sheet.getRange(bid.rowIndex, 17).setValue('');          // Clear New CPM (Q)
        sheet.getRange(bid.rowIndex, 19).setValue('');          // Clear Comment (S)
        
        bidLogEntries.push({
          timestamp: timestamp,
          campaignId: bid.campaignId,
          campaignName: bid.campaignName,
          bidId: bid.bidId,
          spotId: bid.spotId,
          spotName: bid.spotName,
          device: bid.deviceOS,
          country: bid.country,
          oldCpm: bid.currentBid,
          newCpm: result.bid || bid.newBid,
          changePercent: bid.change,
          comment: bid.comment,
          status: 'SUCCESS',
          error: ''
        });
        
        Logger.log(`✅ Updated bid ${bid.bidId}: $${bid.currentBid} → $${bid.newBid}`);
      } else {
        bidFail++;
        bidLogEntries.push({
          timestamp: timestamp,
          campaignId: bid.campaignId,
          campaignName: bid.campaignName,
          bidId: bid.bidId,
          spotId: bid.spotId,
          spotName: bid.spotName,
          device: bid.deviceOS,
          country: bid.country,
          oldCpm: bid.currentBid,
          newCpm: bid.newBid,
          changePercent: bid.change,
          comment: bid.comment,
          status: 'FAILED',
          error: response.getContentText().substring(0, 200)
        });
        Logger.log(`❌ Failed bid ${bid.bidId}`);
      }
      
      Utilities.sleep(200);
      
    } catch (e) {
      bidFail++;
      bidLogEntries.push({
        timestamp: timestamp,
        campaignId: bid.campaignId,
        campaignName: bid.campaignName,
        bidId: bid.bidId,
        spotId: bid.spotId,
        spotName: bid.spotName,
        device: bid.deviceOS,
        country: bid.country,
        oldCpm: bid.currentBid,
        newCpm: bid.newBid,
        changePercent: bid.change,
        comment: bid.comment,
        status: 'ERROR',
        error: e.toString().substring(0, 200)
      });
      Logger.log(`❌ Error bid ${bid.bidId}: ${e}`);
    }
  }
  
  // Process budget updates
  let budgetSuccess = 0;
  let budgetFail = 0;
  const budgetLogEntries = [];
  
  for (const campaignId of budgetCampaigns) {
    const budget = budgetUpdates[campaignId];
    const timestamp = new Date();
    
    try {
      const result = cpUpdateCampaignBudget(campaignId, budget.newBudget);
      
      if (result.success) {
        budgetSuccess++;
        
        // Update pivot sheet rows for this campaign (column positions with Format column)
        for (const rowIndex of budget.rowIndices) {
          sheet.getRange(rowIndex, 15).setValue(budget.newBudget);  // Update Daily Budget (O)
          sheet.getRange(rowIndex, 18).setValue('');                 // Clear New Budget (R)
          sheet.getRange(rowIndex, 19).setValue('');                 // Clear Comment (S)
        }
        
        budgetLogEntries.push({
          timestamp: timestamp,
          campaignId: campaignId,
          campaignName: budget.campaignName,
          oldBudget: budget.currentBudget,
          newBudget: budget.newBudget,
          changePercent: ((budget.newBudget - budget.currentBudget) / budget.currentBudget * 100).toFixed(2),
          comment: budget.comment,
          status: 'SUCCESS',
          error: ''
        });
        
        Logger.log(`✅ Updated budget for ${campaignId}`);
      } else {
        budgetFail++;
        budgetLogEntries.push({
          timestamp: timestamp,
          campaignId: campaignId,
          campaignName: budget.campaignName,
          oldBudget: budget.currentBudget,
          newBudget: budget.newBudget,
          changePercent: ((budget.newBudget - budget.currentBudget) / budget.currentBudget * 100).toFixed(2),
          comment: budget.comment,
          status: 'FAILED',
          error: result.responseText.substring(0, 200)
        });
        Logger.log(`❌ Failed budget for ${campaignId}`);
      }
      
      Utilities.sleep(200);
      
    } catch (e) {
      budgetFail++;
      budgetLogEntries.push({
        timestamp: timestamp,
        campaignId: campaignId,
        campaignName: budget.campaignName,
        oldBudget: budget.currentBudget,
        newBudget: budget.newBudget,
        changePercent: ((budget.newBudget - budget.currentBudget) / budget.currentBudget * 100).toFixed(2),
        comment: budget.comment,
        status: 'ERROR',
        error: e.toString().substring(0, 200)
      });
      Logger.log(`❌ Error budget for ${campaignId}: ${e}`);
    }
  }
  
  // Write logs
  if (bidLogEntries.length > 0) {
    cpWriteBidLogs(bidLogEntries);
  }
  if (budgetLogEntries.length > 0) {
    cpWriteBudgetLogs(budgetLogEntries);
  }
  
  // Show results
  let resultMsg = `✅ Update Complete!\n\n`;
  
  if (bidUpdates.length > 0) {
    resultMsg += `BIDS: ${bidSuccess} successful, ${bidFail} failed\n`;
  }
  if (budgetCampaigns.length > 0) {
    resultMsg += `BUDGETS: ${budgetSuccess} successful, ${budgetFail} failed\n`;
  }
  
  ui.alert('Update Results', resultMsg, ui.ButtonSet.OK);
}
