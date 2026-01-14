/**
 * TJ Control Panel V1 - Google Apps Script
 * 
 * A dashboard for bid optimization with multi-period stats (Today, Yesterday, 7-Day).
 * Uses BID-LEVEL stats for all periods - granular per-bid T/Y/7D comparisons.
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
    .addItem('📋 Copy Bids to New Column', 'cpCopyBidsToNew')
    .addItem('📈 Calculate Bid Changes', 'cpCalculateBidChanges')
    .addItem('🚀 UPDATE BIDS IN TJ', 'cpUpdateBids')
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
    
    // Step 3: Build Legend lookup
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
    
    // Step 4: Merge data into rows
    const rows = cpMergeBidData(currentBids, todayStats, yesterdayStats, sevenDayStats, legendLookup);
    
    // Step 5: Write to sheet
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
 */
function cpMergeBidData(currentBids, todayStats, yesterdayStats, sevenDayStats, legendLookup) {
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
      currentBid: cpToNumeric(bid.bid, 0),
      newBid: '',  // User editable
      // Columns J-L are formulas
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
      sdActiveDays: sdActiveDays,  // NEW: Active days in 7D period
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
 */
function cpWriteToSheet(rows) {
  const sheet = cpGetOrCreateSheet(CP_SHEET_NAME);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear existing data
  sheet.clear();
  
  // Define headers (32 columns: A-AF)
  const headers = [
    'Tier 1 Strategy',      // A
    'Sub Strategy',         // B
    'Campaign Name',        // C
    'Campaign ID',          // D
    'Format',               // E
    'Country',              // F
    'Device + iOS',         // G
    'Current eCPM Bid',     // H
    'New CPM Bid',          // I
    'Change %',             // J - Formula
    'T Bid Adjust',         // K - Formula
    'Date last bid Adjust', // L - Formula
    'T eCPM',               // M
    'Y eCPM',               // N
    '7D eCPM',              // O
    'T Spend',              // P
    'Y Spend',              // Q
    '7D Spend',             // R (avg per active day)
    'T CPA',                // S
    'Y CPA',                // T
    '7D CPA',               // U
    'T Conv',               // V
    'Y Conv',               // W
    '7D Conv',              // X (avg per active day)
    'T CTR',                // Y
    'Y CTR',                // Z
    '7D CTR',               // AA
    '7D Active Days',       // AB - NEW: Number of active days in 7D period
    'Spot ID',              // AC
    'Bid ID',               // AD
    'Geo ID',               // AE
    'Last Updated'          // AF
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
  
  // Prepare data rows
  const dataRows = rows.map(row => [
    row.tier1Strategy,      // A
    row.subStrategy,        // B
    row.campaignName,       // C
    row.campaignId,         // D
    row.format,             // E
    row.country,            // F
    row.deviceOS,           // G
    row.currentBid,         // H
    '',                     // I - New CPM Bid (empty for user)
    '',                     // J - Change % (formula)
    '',                     // K - T Bid Adjust (formula)
    '',                     // L - Date last bid Adjust (formula)
    row.tEcpm,              // M
    row.yEcpm,              // N
    row.sdEcpm,             // O
    row.tSpend,             // P
    row.ySpend,             // Q
    row.sdSpend,            // R (avg per active day)
    row.tCpa,               // S
    row.yCpa,               // T
    row.sdCpa,              // U
    row.tConv,              // V
    row.yConv,              // W
    row.sdConv,             // X (avg per active day)
    row.tCtr,               // Y
    row.yCtr,               // Z
    row.sdCtr,              // AA
    row.sdActiveDays,       // AB - Active days in 7D period
    row.spotId,             // AC
    row.bidId,              // AD
    row.geoId,              // AE
    row.lastUpdated         // AF
  ]);
  
  // Write data
  sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  
  // Add formulas (Bid ID is now in column AD)
  for (let i = 2; i <= dataRows.length + 1; i++) {
    // Column J: Change % = (New - Current) / Current * 100
    sheet.getRange(i, 10).setFormula(`=IF(AND(H${i}>0,I${i}<>""),(I${i}-H${i})/H${i}*100,"")`);
    
    // Column K: T Bid Adjust - was bid adjusted today? Show "Yes ($previous_bid)" if so
    // Use FILTER to get the Old CPM (column I) from today's adjustment
    sheet.getRange(i, 11).setFormula(`=IFERROR(IF(ROWS(FILTER('Bid Logs'!$I$2:$I$10000,(TEXT('Bid Logs'!$D$2:$D$10000,"0")=TEXT(AD${i},"0"))*('Bid Logs'!$L$2:$L$10000="SUCCESS")*(INT('Bid Logs'!$A$2:$A$10000)=TODAY())))>0,"Yes ("&TEXT(INDEX(FILTER('Bid Logs'!$I$2:$I$10000,(TEXT('Bid Logs'!$D$2:$D$10000,"0")=TEXT(AD${i},"0"))*('Bid Logs'!$L$2:$L$10000="SUCCESS")*(INT('Bid Logs'!$A$2:$A$10000)=TODAY())),1),"$#,##0.000")&")",""),"")`);
    
    // Column L: Date last bid Adjust - before today
    // Use bounded ranges for FILTER as well
    sheet.getRange(i, 12).setFormula(`=IFERROR(TEXT(MAX(FILTER('Bid Logs'!$A$2:$A$10000,(TEXT('Bid Logs'!$D$2:$D$10000,"0")=TEXT(AD${i},"0"))*('Bid Logs'!$L$2:$L$10000="SUCCESS")*(INT('Bid Logs'!$A$2:$A$10000)<TODAY()))),"yyyy-mm-dd"),"Never")`);
  }
  
  // Format columns
  const numRows = dataRows.length;
  
  // Strategy columns (A-B) - light blue background
  sheet.getRange(2, 1, numRows, 2).setBackground('#e3f2fd');
  
  // Current eCPM Bid (H) - currency
  sheet.getRange(2, 8, numRows, 1).setNumberFormat('$#,##0.000');
  
  // New CPM Bid (I) - currency, yellow background (editable)
  sheet.getRange(2, 9, numRows, 1)
    .setNumberFormat('$#,##0.000')
    .setBackground('#fff9c4');
  
  // Change % (J) - percentage
  sheet.getRange(2, 10, numRows, 1).setNumberFormat('0.00"%"');
  
  // T Bid Adjust (K) - conditional formatting
  const kRange = sheet.getRange(2, 11, numRows, 1);
  
  // Date last bid Adjust (L) - date format
  sheet.getRange(2, 12, numRows, 1).setNumberFormat('yyyy-mm-dd');
  
  // eCPM columns (M-O) - currency
  sheet.getRange(2, 13, numRows, 3).setNumberFormat('$#,##0.000');
  
  // Spend columns (P-R) - accounting format
  sheet.getRange(2, 16, numRows, 3).setNumberFormat('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');
  
  // CPA columns (S-U) - accounting format
  sheet.getRange(2, 19, numRows, 3).setNumberFormat('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');
  
  // Conv columns (V-X)
  // T Conv (V) and Y Conv (W) - whole numbers
  sheet.getRange(2, 22, numRows, 2).setNumberFormat('#,##0');
  // 7D Conv (X) - 1 decimal for averages
  sheet.getRange(2, 24, numRows, 1).setNumberFormat('#,##0.0');
  
  // CTR columns (Y-AA) - percentage
  sheet.getRange(2, 25, numRows, 3).setNumberFormat('0.00"%"');
  
  // 7D Active Days (AB) - whole number, highlight for visibility
  sheet.getRange(2, 28, numRows, 1)
    .setNumberFormat('0')
    .setBackground('#e8f5e9');  // Light green
  
  // ID columns (AC-AE) - plain text
  sheet.getRange(2, 29, numRows, 3).setNumberFormat('@');
  
  // Last Updated (AF) - datetime
  sheet.getRange(2, 32, numRows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  
  // Conditional formatting for T Bid Adjust - green when contains "Yes"
  const rules = sheet.getConditionalFormatRules();
  const yesRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Yes')
    .setBackground('#c8e6c9')
    .setFontColor('#2e7d32')
    .setRanges([kRange])
    .build();
  rules.push(yesRule);
  sheet.setConditionalFormatRules(rules);
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
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
    // New CPM Bid (I) - yellow (editable)
    sheet.getRange(i, 9).setBackground('#fff9c4');
    // 7D Active Days (AB) - light green
    sheet.getRange(i, 28).setBackground('#e8f5e9');
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
  
  // Copy Current eCPM Bid (H) to New CPM Bid (I)
  const currentBids = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
  sheet.getRange(2, 9, lastRow - 1, 1).setValues(currentBids);
  
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
  
  // Get columns H (Current), I (New), D (Campaign ID)
  const data = sheet.getRange(2, 1, lastRow - 1, 32).getValues();
  
  let totalChanges = 0;
  let increased = 0;
  let decreased = 0;
  let unchanged = 0;
  const changedCampaigns = new Set();
  
  for (const row of data) {
    const currentBid = cpToNumeric(row[7], 0);  // H - index 7
    const newBid = cpToNumeric(row[8], 0);      // I - index 8
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
  
  // Get all data (32 columns now)
  const data = sheet.getRange(2, 1, lastRow - 1, 32).getValues();
  
  // Find bids to update
  const bidsToUpdate = [];
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const currentBid = cpToNumeric(row[7], 0);   // H - Current eCPM Bid
    const newBid = cpToNumeric(row[8], 0);       // I - New CPM Bid
    const bidId = String(row[29]);               // AD - Bid ID (was AC, now AD after adding Active Days)
    const campaignName = row[2];                 // C - Campaign Name
    const campaignId = row[3];                   // D - Campaign ID
    const spotId = row[28];                      // AC - Spot ID (was AB, now AC)
    const spotName = row[6];                     // G - Device + iOS
    const country = row[5];                      // F - Country
    
    // Skip if no bid ID or no new bid or same bid
    if (!bidId || newBid === 0 || newBid === currentBid) continue;
    
    bidsToUpdate.push({
      rowIndex: i + 2,
      campaignId: String(campaignId),
      campaignName: campaignName,
      bidId: bidId,
      spotId: String(spotId),
      spotName: spotName,
      country: country,
      currentBid: currentBid,
      newBid: newBid,
      change: ((newBid - currentBid) / currentBid * 100).toFixed(2)
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
        
        // Update sheet
        sheet.getRange(bid.rowIndex, 8).setValue(bid.newBid);  // Update Current eCPM Bid
        sheet.getRange(bid.rowIndex, 9).setValue('');          // Clear New CPM Bid
        
        // Log entry
        logEntries.push({
          timestamp: timestamp,
          campaignId: bid.campaignId,
          campaignName: bid.campaignName,
          bidId: bid.bidId,
          spotId: bid.spotId,
          spotName: bid.spotName,
          device: '',
          country: bid.country,
          oldCpm: bid.currentBid,
          newCpm: result.bid || bid.newBid,
          changePercent: bid.change,
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
          device: '',
          country: bid.country,
          oldCpm: bid.currentBid,
          newCpm: bid.newBid,
          changePercent: bid.change,
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
        device: '',
        country: bid.country,
        oldCpm: bid.currentBid,
        newCpm: bid.newBid,
        changePercent: bid.change,
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
      'Change %', 'Status', 'Error'
    ];
    
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    logSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('white');
    logSheet.setFrozenRows(1);
  }
  
  // Prepare rows
  const logRows = logEntries.map(e => [
    e.timestamp, e.campaignId, e.campaignName, e.bidId, e.spotId,
    e.spotName, e.device, e.country, e.oldCpm, e.newCpm,
    e.changePercent, e.status, e.error
  ]);
  
  // Append
  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow + 1, 1, logRows.length, 13).setValues(logRows);
  
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
