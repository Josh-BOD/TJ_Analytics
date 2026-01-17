/**
 * TJ Control Panel V3 - Google Apps Script
 * 
 * A dashboard for bid optimization with multi-period stats (Today, Yesterday, 7-Day).
 * Uses BID-LEVEL stats for all periods - granular per-bid T/Y/7D comparisons.
 * 
 * V3 ADDITIONS:
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
    .addSubMenu(ui.createMenu('📊 Daily Dashboard')
      .addItem('📥 Pull Daily Stats (7 Days)', 'cpPullDailyStats')
      .addItem('📈 Build/Refresh Dashboard', 'cpBuildDashboard'))
    .addSeparator()
    .addItem('🗑️ Clear Data', 'cpClearData')
    .addToUi();
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

// ============================================================================
// API FETCH FUNCTIONS
// ============================================================================

/**
 * Fetch current bids from /api/bids/{campaignId}.json
 * Returns object keyed by bid_id
 */
function cpFetchCurrentBids(campaignIds) {
  const allBids = {};
  
  for (const campaignId of campaignIds) {
    Logger.log(`Fetching bids for campaign ${campaignId}...`);
    
    // First get campaign name
    let campaignName = '';
    try {
      const campaignUrl = `${CP_API_BASE_URL}/campaigns/${campaignId}.json?api_key=${CP_API_KEY}`;
      const campaignResp = UrlFetchApp.fetch(campaignUrl, { muteHttpExceptions: true });
      if (campaignResp.getResponseCode() === 200) {
        const campaignData = JSON.parse(campaignResp.getContentText());
        campaignName = campaignData.campaign_name || '';
      }
    } catch (e) {
      Logger.log(`Could not fetch campaign name: ${e}`);
    }
    
    // Fetch bids
    const bidsUrl = `${CP_API_BASE_URL}/bids/${campaignId}.json?api_key=${CP_API_KEY}`;
    try {
      const resp = UrlFetchApp.fetch(bidsUrl, { muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) {
        Logger.log(`Error fetching bids for ${campaignId}: ${resp.getResponseCode()}`);
        continue;
      }
      
      const data = JSON.parse(resp.getContentText());
      
      if (typeof data === 'object' && data !== null) {
        const bidIds = Object.keys(data);
        Logger.log(`Found ${bidIds.length} bids for campaign ${campaignId}`);
        
        for (const bidId of bidIds) {
          const bid = data[bidId];
          if (bid && typeof bid === 'object') {
            bid.campaign_id = campaignId;
            bid.campaign_name = campaignName;
            allBids[bidId] = bid;
          }
        }
      }
    } catch (e) {
      Logger.log(`Error fetching bids: ${e}`);
    }
    
    // Small delay between campaigns
    Utilities.sleep(100);
  }
  
  Logger.log(`Total bids fetched: ${Object.keys(allBids).length}`);
  return allBids;
}

/**
 * Fetch BID-LEVEL stats for a single date
 * Uses /api/bids/{campaignId}.json with startDate/endDate params
 * Returns object keyed by bid_id
 */
function cpFetchBidStats(campaignIds, startDate, endDate, periodLabel) {
  Logger.log(`Fetching ${periodLabel} BID stats (${startDate} to ${endDate})...`);
  
  const allStats = {};
  
  for (const campaignId of campaignIds) {
    const url = `${CP_API_BASE_URL}/bids/${campaignId}.json?api_key=${CP_API_KEY}&startDate=${startDate}&endDate=${endDate}`;
    
    try {
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      
      if (resp.getResponseCode() !== 200) {
        Logger.log(`Error fetching ${periodLabel} bids for ${campaignId}: ${resp.getResponseCode()}`);
        continue;
      }
      
      const data = JSON.parse(resp.getContentText());
      
      if (typeof data === 'object' && data !== null) {
        for (const bidId in data) {
          const bid = data[bidId];
          const stats = bid.stats || {};
          
          allStats[bidId] = {
            impressions: cpToNumeric(stats.impressions, 0),
            clicks: cpToNumeric(stats.clicks, 0),
            conversions: cpToNumeric(stats.conversions, 0),
            cost: cpToNumeric(stats.revenue, 0),  // API calls it 'revenue'
            ecpm: cpToNumeric(stats.ecpm, 0),
            ctr: cpToNumeric(stats.ctr, 0)
          };
        }
      }
      
    } catch (e) {
      Logger.log(`Error fetching ${periodLabel} bids for ${campaignId}: ${e}`);
    }
    
    // Small delay between campaigns
    Utilities.sleep(100);
  }
  
  Logger.log(`Got ${periodLabel} stats for ${Object.keys(allStats).length} bids`);
  return allStats;
}

/**
 * Fetch BID-LEVEL stats for each day in a 7-day period
 * Returns object keyed by bid_id with totals AND active day count
 */
function cpFetch7DayBidStats(campaignIds) {
  Logger.log('Fetching 7-day bid stats (day by day)...');
  
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
  
  // Collect stats per bid per day
  const bidDailyStats = {};  // bidId -> array of daily stats
  
  for (const dateStr of dates) {
    Logger.log(`  Fetching ${dateStr}...`);
    
    for (const campaignId of campaignIds) {
      const url = `${CP_API_BASE_URL}/bids/${campaignId}.json?api_key=${CP_API_KEY}&startDate=${dateStr}&endDate=${dateStr}`;
      
      try {
        const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        
        if (resp.getResponseCode() !== 200) continue;
        
        const data = JSON.parse(resp.getContentText());
        
        if (typeof data === 'object' && data !== null) {
          for (const bidId in data) {
            const bid = data[bidId];
            const stats = bid.stats || {};
            
            if (!bidDailyStats[bidId]) {
              bidDailyStats[bidId] = [];
            }
            
            bidDailyStats[bidId].push({
              impressions: cpToNumeric(stats.impressions, 0),
              clicks: cpToNumeric(stats.clicks, 0),
              conversions: cpToNumeric(stats.conversions, 0),
              cost: cpToNumeric(stats.revenue, 0),
              ecpm: cpToNumeric(stats.ecpm, 0),
              ctr: cpToNumeric(stats.ctr, 0)
            });
          }
        }
        
      } catch (e) {
        Logger.log(`Error: ${e}`);
      }
    }
    
    Utilities.sleep(100);  // Small delay between days
  }
  
  // Calculate totals and active days per bid
  const result = {};
  
  for (const bidId in bidDailyStats) {
    const dailyStats = bidDailyStats[bidId];
    
    // Count active days (days with impressions > 0)
    let activeDays = 0;
    let totalImpressions = 0;
    let totalClicks = 0;
    let totalConversions = 0;
    let totalCost = 0;
    let totalEcpmSum = 0;
    let totalCtrSum = 0;
    
    for (const day of dailyStats) {
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
    
    // Calculate averages based on active days (not 7)
    const divisor = activeDays > 0 ? activeDays : 1;
    
    result[bidId] = {
      impressions: totalImpressions,
      clicks: totalClicks,
      conversions: totalConversions,
      cost: totalCost,
      // Averages per active day
      avgImpressions: totalImpressions / divisor,
      avgClicks: totalClicks / divisor,
      avgConversions: totalConversions / divisor,
      avgCost: totalCost / divisor,
      // Weighted averages for eCPM and CTR
      ecpm: activeDays > 0 ? totalEcpmSum / activeDays : 0,
      ctr: activeDays > 0 ? totalCtrSum / activeDays : 0,
      // Active days count
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
 * Returns object keyed by campaign_id with budget info
 */
function cpFetchCampaignBudgets(campaignIds) {
  Logger.log(`Fetching budgets for ${campaignIds.length} campaigns...`);
  
  const budgets = {};
  
  for (const campaignId of campaignIds) {
    const url = `${CP_API_BASE_URL}/campaigns/${campaignId}.json?api_key=${CP_API_KEY}`;
    
    try {
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      
      if (resp.getResponseCode() === 200) {
        const data = JSON.parse(resp.getContentText());
        
        budgets[campaignId] = {
          campaignId: campaignId,
          campaignName: data.campaign_name || '',
          dailyBudget: cpToNumeric(data.campaign_daily_budget, 0),
          budgetLeft: cpToNumeric(data.daily_budget_left, 0),
          status: data.status || ''
        };
        
        Logger.log(`  ${campaignId}: Budget=$${budgets[campaignId].dailyBudget}, Left=$${budgets[campaignId].budgetLeft}`);
      } else {
        Logger.log(`  Error fetching ${campaignId}: ${resp.getResponseCode()}`);
      }
    } catch (e) {
      Logger.log(`  Error fetching ${campaignId}: ${e}`);
    }
    
    Utilities.sleep(100);
  }
  
  Logger.log(`Fetched budgets for ${Object.keys(budgets).length} campaigns`);
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
    
    Logger.log(`Processing ${campaignIds.length} campaigns: ${campaignIds.join(', ')}`);
    
    // Get date ranges
    const dateRanges = cpGetDateRanges();
    Logger.log(`Date ranges: Today=${dateRanges.today.start}, Yesterday=${dateRanges.yesterday.start}, 7D=${dateRanges.sevenDay.start} to ${dateRanges.sevenDay.end}`);
    
    // Step 1: Fetch current bids (bid-level, no date filter - gets current bid values)
    Logger.log('Fetching current bids...');
    const currentBids = cpFetchCurrentBids(campaignIds);
    
    // Step 2: Fetch BID-LEVEL stats for Today and Yesterday
    Logger.log('Fetching bid stats for Today...');
    const todayStats = cpFetchBidStats(campaignIds, dateRanges.today.start, dateRanges.today.end, 'Today');
    Utilities.sleep(200);
    
    Logger.log('Fetching bid stats for Yesterday...');
    const yesterdayStats = cpFetchBidStats(campaignIds, dateRanges.yesterday.start, dateRanges.yesterday.end, 'Yesterday');
    Utilities.sleep(200);
    
    // Step 3: Fetch 7-day stats day-by-day for accurate active day calculation
    Logger.log('Fetching 7-day stats (day by day for active day calculation)...');
    const sevenDayStats = cpFetch7DayBidStats(campaignIds);
    
    // Step 4: Fetch campaign budgets (V3)
    Logger.log('Fetching campaign budgets...');
    const campaignBudgets = cpFetchCampaignBudgets(campaignIds);
    
    // Step 5: Build Legend lookup
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
    
    // Step 6: Merge data into rows
    const rows = cpMergeBidData(currentBids, todayStats, yesterdayStats, sevenDayStats, legendLookup, campaignBudgets);
    
    // Step 7: Write to sheet
    cpWriteToSheet(rows);
    
    ui.alert('Success', 
      `Refreshed Control Panel with ${rows.length} bid entries from ${campaignIds.length} campaign(s).\n\n` +
      `Date ranges:\n` +
      `• Today: ${dateRanges.today.start}\n` +
      `• Yesterday: ${dateRanges.yesterday.start}\n` +
      `• 7-Day: ${dateRanges.sevenDay.start} to ${dateRanges.sevenDay.end}`,
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    ui.alert('Error', `Failed to refresh: ${error.toString()}`, ui.ButtonSet.OK);
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
    
    rows.push({
      tier1Strategy: legend.strategy || '',
      subStrategy: legend.subStrategy || '',
      campaignName: bid.campaign_name || '',
      campaignId: campaignId,
      format: legend.format || '',
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
  
  // Write data
  sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  
  // Add formulas (Bid ID is now in column AH)
  for (let i = 2; i <= dataRows.length + 1; i++) {
    // Column K: Change % = (New - Current) / Current * 100
    // Current Bid is I, New Bid is J
    sheet.getRange(i, 11).setFormula(`=IF(AND(I${i}>0,J${i}<>""),(J${i}-I${i})/I${i}*100,"")`);
    
    // Column M: T Bid Adjust - was bid adjusted today? Show "Yes ($previous_bid)" if so
    // Bid ID is now in AH, Old CPM is column I in Bid Logs, Status is column M
    sheet.getRange(i, 13).setFormula(`=IFERROR(IF(ROWS(FILTER('Bid Logs'!$I$2:$I$10000,(TEXT('Bid Logs'!$D$2:$D$10000,"0")=TEXT(AH${i},"0"))*('Bid Logs'!$M$2:$M$10000="SUCCESS")*(INT('Bid Logs'!$A$2:$A$10000)=TODAY())))>0,"Yes ("&TEXT(INDEX(FILTER('Bid Logs'!$I$2:$I$10000,(TEXT('Bid Logs'!$D$2:$D$10000,"0")=TEXT(AH${i},"0"))*('Bid Logs'!$M$2:$M$10000="SUCCESS")*(INT('Bid Logs'!$A$2:$A$10000)=TODAY())),1),"$#,##0.000")&")",""),"")`);
    
    // Column N: Date last bid Adjust - before today, with bid value
    // Format: "2026-01-13 ($2.600)" - New CPM is column J in Bid Logs
    sheet.getRange(i, 14).setFormula(`=IFERROR(LET(data,SORTN(FILTER({'Bid Logs'!$A$2:$A$10000,'Bid Logs'!$J$2:$J$10000},(TEXT('Bid Logs'!$D$2:$D$10000,"0")=TEXT(AH${i},"0"))*('Bid Logs'!$M$2:$M$10000="SUCCESS")*(INT('Bid Logs'!$A$2:$A$10000)<TODAY())),1,0,1,FALSE),TEXT(INDEX(data,1,1),"yyyy-mm-dd")&" ($"&TEXT(INDEX(data,1,2),"#,##0.000")&")"),"Never")`);
  }
  
  // Format columns - V3 with new column positions
  const numRows = dataRows.length;
  
  // Strategy columns (A-B) - light blue background
  sheet.getRange(2, 1, numRows, 2).setBackground('#e3f2fd');
  
  // Spot Name (H) - text
  sheet.getRange(2, 8, numRows, 1).setNumberFormat('@');
  
  // Current eCPM Bid (I) - currency
  sheet.getRange(2, 9, numRows, 1).setNumberFormat('$#,##0.000');
  
  // New CPM Bid (J) - currency, yellow background (editable)
  sheet.getRange(2, 10, numRows, 1)
    .setNumberFormat('$#,##0.000')
    .setBackground('#fff9c4');
  
  // Change % (K) - percentage
  sheet.getRange(2, 11, numRows, 1).setNumberFormat('0.00"%"');
  
  // Comment (L) - text, yellow background (editable)
  sheet.getRange(2, 12, numRows, 1)
    .setNumberFormat('@')
    .setBackground('#fff9c4');
  
  // T Bid Adjust (M) - conditional formatting
  const mRange = sheet.getRange(2, 13, numRows, 1);
  
  // Date last bid Adjust (N) - text format (contains formula result)
  sheet.getRange(2, 14, numRows, 1).setNumberFormat('@');
  
  // V3: Budget columns (O-P)
  // Daily Budget (O) - currency
  sheet.getRange(2, 15, numRows, 1).setNumberFormat('$#,##0.00');
  // New Budget (P) - currency, yellow background (editable)
  sheet.getRange(2, 16, numRows, 1)
    .setNumberFormat('$#,##0.00')
    .setBackground('#fff9c4');
  
  // eCPM columns (Q-S) - currency
  sheet.getRange(2, 17, numRows, 3).setNumberFormat('$#,##0.000');
  
  // Spend columns (T-V) - accounting format
  sheet.getRange(2, 20, numRows, 3).setNumberFormat('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');
  
  // CPA columns (W-Y) - accounting format
  sheet.getRange(2, 23, numRows, 3).setNumberFormat('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');
  
  // Conv columns (Z-AB)
  // T Conv (Z) and Y Conv (AA) - whole numbers
  sheet.getRange(2, 26, numRows, 2).setNumberFormat('#,##0');
  // 7D Conv (AB) - 1 decimal for averages
  sheet.getRange(2, 28, numRows, 1).setNumberFormat('#,##0.0');
  
  // CTR columns (AC-AE) - percentage
  sheet.getRange(2, 29, numRows, 3).setNumberFormat('0.00"%"');
  
  // 7D Active Days (AF) - whole number, highlight for visibility
  sheet.getRange(2, 32, numRows, 1)
    .setNumberFormat('0')
    .setBackground('#e8f5e9');  // Light green
  
  // ID columns (AG-AI) - plain text
  sheet.getRange(2, 33, numRows, 3).setNumberFormat('@');
  
  // Last Updated (AJ) - datetime
  sheet.getRange(2, 36, numRows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  
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
  
  // Freeze header row and columns A-D
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4);  // Freeze columns A-D
  
  // Hide columns A and B (Tier 1 Strategy, Sub Strategy)
  sheet.hideColumns(1, 2);  // Hide columns A-B
  
  // Add alternating row colors (zebra striping) for easier reading
  for (let i = 2; i <= numRows + 1; i++) {
    if (i % 2 === 0) {
      // Even rows - light blue
      sheet.getRange(i, 1, 1, headers.length).setBackground('#e3f2fd');
    } else {
      // Odd rows - white
      sheet.getRange(i, 1, 1, headers.length).setBackground('#ffffff');
    }
  }
  
  // Re-apply special column backgrounds on top of zebra striping
  for (let i = 2; i <= numRows + 1; i++) {
    // New CPM Bid (J) - yellow (editable)
    sheet.getRange(i, 10).setBackground('#fff9c4');
    // Comment (L) - yellow (editable)
    sheet.getRange(i, 12).setBackground('#fff9c4');
    // New Budget (P) - yellow (editable) - V3
    sheet.getRange(i, 16).setBackground('#fff9c4');
    // 7D Active Days (AF) - light green
    sheet.getRange(i, 32).setBackground('#e8f5e9');
  }
  
  // Remove existing filter if present, then add new filter
  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }
  const dataRange = sheet.getRange(1, 1, numRows + 1, headers.length);
  dataRange.createFilter();
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  
  // Make Campaign Name wider
  sheet.setColumnWidth(3, 300);
  
  // Activate sheet
  ss.setActiveSheet(sheet);
  
  Logger.log(`Wrote ${dataRows.length} rows to ${CP_SHEET_NAME}`);
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
  if (lastRow < 2) {
    ui.alert('Error', 'No data found. Please refresh data first.', ui.ButtonSet.OK);
    return;
  }
  
  // Copy Current eCPM Bid (I) to New CPM Bid (J)
  const currentBids = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  sheet.getRange(2, 10, lastRow - 1, 1).setValues(currentBids);
  
  ui.alert('Done', `Copied ${lastRow - 1} bid values to "New CPM Bid" column.`, ui.ButtonSet.OK);
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
  if (lastRow < 2) {
    ui.alert('Error', 'No data found.', ui.ButtonSet.OK);
    return;
  }
  
  // Get data - V3: 36 columns
  const data = sheet.getRange(2, 1, lastRow - 1, 36).getValues();
  
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
  if (lastRow < 2) {
    ui.alert('Error', 'No data found.', ui.ButtonSet.OK);
    return;
  }
  
  // Get all data (36 columns now - V3)
  const data = sheet.getRange(2, 1, lastRow - 1, 36).getValues();
  
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
  if (lastRow < 2) {
    ui.alert('Error', 'No data found. Please refresh data first.', ui.ButtonSet.OK);
    return;
  }
  
  // Copy Daily Budget (O) to New Budget (P)
  const currentBudgets = sheet.getRange(2, 15, lastRow - 1, 1).getValues();
  sheet.getRange(2, 16, lastRow - 1, 1).setValues(currentBudgets);
  
  ui.alert('Done', `Copied ${lastRow - 1} budget values to "New Budget" column.`, ui.ButtonSet.OK);
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
  if (lastRow < 2) {
    ui.alert('Error', 'No data found.', ui.ButtonSet.OK);
    return;
  }
  
  // Get all data (36 columns now)
  const data = sheet.getRange(2, 1, lastRow - 1, 36).getValues();
  
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
  if (lastRow < 2) {
    ui.alert('Error', 'No data found.', ui.ButtonSet.OK);
    return;
  }
  
  // Get all data (36 columns)
  const data = sheet.getRange(2, 1, lastRow - 1, 36).getValues();
  
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
