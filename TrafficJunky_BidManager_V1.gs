/**
 * TrafficJunky Bid Manager for Google Sheets
 * 
 * This script pulls campaign bids from TrafficJunky API and displays them
 * in an editable format for bid management.
 * 
 * SETUP INSTRUCTIONS:
 * 1. Create a new Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code and paste this entire script
 * 4. Save the script (Ctrl+S / Cmd+S)
 * 5. Close the Apps Script editor
 * 6. Refresh your Google Sheet
 * 7. You should see a "Bid Manager" menu appear
 * 
 * USAGE:
 * - Click "Bid Manager" > "🔄 Pull All Bids" to fetch current bids
 * - Edit the "Current Bid" column as needed
 * - Use "📋 Copy Bids to New Column" to create a "New Bid" column for your adjustments
 */

// ============================================================================
// CONFIGURATION - Update these values as needed
// ============================================================================

const API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039";
const API_BASE_URL = "https://api.trafficjunky.com/api";
const API_URL_STATS = "https://api.trafficjunky.com/api/campaigns/stats.json"; // For campaign stats
const API_URL_CAMPAIGNS_BIDS = "https://api.trafficjunky.com/api/campaigns/bids/stats.json"; // For all campaigns with bids
const BIDS_SHEET_NAME = "Campaign Bids";
const API_TIMEZONE = "America/New_York";

// Campaign IDs are now pulled dynamically from the Legend sheet (Column A)
// The CAMPAIGN_IDS constant below is only used as fallback if Legend sheet doesn't exist
const CAMPAIGN_IDS_FALLBACK = ["1013232471"];

// ============================================================================
// MENU SETUP
// ============================================================================

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Bid Manager')
    .addItem('🔄 Pull All Bids', 'pullAllBids')
    .addItem('📊 Pull Active Campaign Bids Only', 'pullActiveBids')
    .addItem('📈 Pull Campaign Stats', 'pullCampaignStats')
    .addSeparator()
    .addItem('📋 Copy Bids to New Column', 'copyBidsToNewColumn')
    .addItem('📈 Calculate Bid Changes', 'calculateBidChanges')
    .addItem('🚀 UPDATE BIDS IN TRAFFICJUNKY', 'updateBidsInTJ')
    .addSeparator()
    .addItem('🔍 Debug API Response', 'debugAPIResponse')
    .addItem('🧪 Explore API Endpoints', 'exploreAPIEndpoints')
    .addItem('📋 List All Campaign IDs', 'listAllCampaignIds')
    .addItem('🔎 Search Campaign by Name', 'searchCampaignByName')
    .addSeparator()
    .addItem('🗑️ Clear Bid Data', 'clearBidData')
    .addItem('🗑️ Clear Bid Logs', 'clearBidLogs')
    .addToUi();
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Get current date/time in EST timezone
 */
function getESTDate() {
  return new Date(new Date().toLocaleString("en-US", {timeZone: API_TIMEZONE}));
}

/**
 * Format date to DD/MM/YYYY for API
 */
function formatDateForAPI(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

/**
 * Convert value to numeric safely
 */
function toNumeric(value, defaultValue = 0) {
  if (value === null || value === undefined || value === '') return defaultValue;
  const num = Number(value);
  return isNaN(num) ? defaultValue : num;
}

/**
 * Get or create sheet by name
 */
function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`Created new sheet: ${sheetName}`);
  }
  
  return sheet;
}

// ============================================================================
// MAIN FUNCTIONS
// ============================================================================

/**
 * Pull all campaign bids from TrafficJunky API
 */
function pullAllBids() {
  pullBids(false);
}

/**
 * Pull only active campaign bids
 */
function pullActiveBids() {
  pullBids(true);
}

/**
 * Main function to pull bids from API
 * Uses /api/bids/{campaignId}.json which returns full bid details including:
 * - spot_id, spot_name (Source info)
 * - bid (Your CPM)
 * - geos (Country targeting)
 * - stats (impressions, clicks, conversions, etc.)
 */
function pullBids(activeOnly = false) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Get campaign IDs from Legend sheet (Column A)
    let campaignIds = [];
    const legendSheet = ss.getSheetByName('Legend');
    
    if (legendSheet) {
      const legendLastRow = legendSheet.getLastRow();
      if (legendLastRow >= 2) {
        const legendCampaignIds = legendSheet.getRange(2, 1, legendLastRow - 1, 1).getValues();
        campaignIds = legendCampaignIds
          .map(row => String(row[0]).trim())
          .filter(id => id && id !== '');
      }
    }
    
    // Fallback to hardcoded IDs if Legend sheet is empty or missing
    if (campaignIds.length === 0) {
      campaignIds = CAMPAIGN_IDS_FALLBACK;
      Logger.log('Legend sheet empty or not found - using fallback campaign IDs');
    }
    
    Logger.log(`Starting bid pull for ${campaignIds.length} campaigns from Legend sheet: ${campaignIds.join(', ')}`);
    
    if (campaignIds.length === 0) {
      ui.alert('Error', 'No campaign IDs found in Legend sheet (Column A).', ui.ButtonSet.OK);
      return;
    }
    
    const bidRows = [];
    
    // =========================================================================
    // Fetch bids for each campaign using /api/bids/{campaignId}.json
    // This endpoint returns full bid details with spot info and stats
    // NOTE: For campaign-level Avg eCPM, use "Pull Campaign Stats" menu option
    //       to populate the "Campaign Stats" sheet, then VLOOKUP from there.
    // =========================================================================
    
    for (const campaignId of campaignIds) {
      Logger.log(`\n=== Fetching bids for campaign ${campaignId} ===`);
      
      // First get campaign name from /api/campaigns/{campaignId}.json
      const campaignUrl = `${API_BASE_URL}/campaigns/${campaignId}.json?api_key=${API_KEY}`;
      let campaignName = '';
      let campaignStatus = 'active';
      let campaignType = '';  // Format (e.g., "banner", "video", etc.)
      
      try {
        const campaignResponse = UrlFetchApp.fetch(campaignUrl, { muteHttpExceptions: true });
        if (campaignResponse.getResponseCode() === 200) {
          const campaignData = JSON.parse(campaignResponse.getContentText());
          
          // Log all available fields to help debug
          Logger.log(`Campaign data fields: ${Object.keys(campaignData).join(', ')}`);
          
          campaignName = campaignData.campaign_name || '';
          campaignStatus = campaignData.status || 'active';
          
          // Try multiple possible field names for campaign type/format
          campaignType = campaignData.campaign_type || 
                         campaignData.campaignType || 
                         campaignData.type || 
                         campaignData.format || 
                         campaignData.creative_type ||
                         campaignData.creativeType ||
                         campaignData.ad_format ||
                         '';
          
          Logger.log(`Campaign name: ${campaignName}, Status: ${campaignStatus}, Type: ${campaignType}`);
        }
      } catch (e) {
        Logger.log(`Could not fetch campaign details: ${e}`);
      }
      
      // Now get the bids with full details
      const url = `${API_BASE_URL}/bids/${campaignId}.json?api_key=${API_KEY}`;
      Logger.log(`URL: ${url.replace(API_KEY, 'HIDDEN')}`);
      
      const response = UrlFetchApp.fetch(url, {
        'method': 'get',
        'contentType': 'application/json',
        'muteHttpExceptions': true
      });
      
      const responseCode = response.getResponseCode();
      Logger.log(`Response code: ${responseCode}`);
      
      if (responseCode !== 200) {
        Logger.log(`Error response: ${response.getContentText()}`);
        ui.alert('API Error', `Campaign ${campaignId}: API returned ${responseCode}\n\n${response.getContentText()}`, ui.ButtonSet.OK);
        continue;
      }
      
      const data = JSON.parse(response.getContentText());
      
      // Response is an object with bid_ids as keys
      // { "1201505001": { bid_id, bid, spot_id, spot_name, geos, stats, ... }, ... }
      
      if (typeof data !== 'object' || data === null) {
        Logger.log(`Unexpected response type: ${typeof data}`);
        continue;
      }
      
      const bidIds = Object.keys(data);
      Logger.log(`Got ${bidIds.length} bids`);
      
      // Log first bid structure with full details
      if (bidIds.length > 0) {
        const firstBid = data[bidIds[0]];
        Logger.log(`First bid ALL FIELDS: ${Object.keys(firstBid).join(', ')}`);
        
        // Log the full first bid object to see all data
        Logger.log(`First bid FULL DATA: ${JSON.stringify(firstBid, null, 2)}`);
        
        // Log geos structure if available
        if (firstBid.geos) {
          const geoKeys = Object.keys(firstBid.geos);
          if (geoKeys.length > 0) {
            Logger.log(`First geo ALL FIELDS: ${Object.keys(firstBid.geos[geoKeys[0]]).join(', ')}`);
          }
        }
      }
      
      // Process each bid
      for (const bidId of bidIds) {
        const bid = data[bidId];
        if (!bid || typeof bid !== 'object') continue;
        
        // Skip if activeOnly and bid is paused
        if (activeOnly && (bid.isPaused === 1 || bid.isActive === false)) {
          continue;
        }
        
        // Extract device from spot_name (e.g., "Pornhub Mobile - Preroll" -> "Mobile")
        const spotName = bid.spot_name || '';
        let device = '';
        if (spotName.includes('Mobile')) device = 'Mobile';
        else if (spotName.includes('Tablet')) device = 'Tablet';
        else if (spotName.includes('PC')) device = 'Desktop';
        
        // Extract OS from campaign name (e.g., "US_EN_PREROLL_CPM_PH_Key-Hentai_AND_M_JB" -> "Android")
        let os = '';
        const campaignNameUpper = campaignName.toUpperCase();
        if (campaignNameUpper.includes('_AND_') || campaignNameUpper.includes('_AND') || 
            campaignNameUpper.startsWith('AND_') || campaignNameUpper.includes('-AND_') ||
            campaignNameUpper.includes('-AND-')) {
          os = 'Android';
        } else if (campaignNameUpper.includes('_IOS_') || campaignNameUpper.includes('_IOS') || 
                   campaignNameUpper.startsWith('IOS_') || campaignNameUpper.includes('-IOS_') ||
                   campaignNameUpper.includes('-IOS-')) {
          os = 'iOS';
        }
        
        // Extract geo info from geos object - collect ALL countries
        let countries = [];
        let geoIds = [];
        const geos = bid.geos || {};
        const geoKeys = Object.keys(geos);
        
        // Log the geos for debugging
        Logger.log(`Bid ${bid.bid_id} has ${geoKeys.length} geos: ${JSON.stringify(geos)}`);
        
        for (const geoKey of geoKeys) {
          const geo = geos[geoKey];
          if (geo.countryCode && !countries.includes(geo.countryCode)) {
            countries.push(geo.countryCode);
          }
          geoIds.push(geo.geoId || geoKey);
        }
        
        Logger.log(`Bid ${bid.bid_id} countries found: ${countries.join(', ')}`)
        
        // Join all countries with comma, or show count if too many
        let countryDisplay = '';
        if (countries.length <= 5) {
          countryDisplay = countries.join(', ');
        } else {
          countryDisplay = `${countries.slice(0, 3).join(', ')} (+${countries.length - 3} more)`;
        }
        
        // Show geo ID count if multiple
        let geoIdDisplay = '';
        if (geoIds.length === 1) {
          geoIdDisplay = String(geoIds[0]);
        } else if (geoIds.length > 1) {
          geoIdDisplay = `${geoIds.length} geos`;
        }
        
        // Check for time targeting
        let timeTargeting = '';
        if (bid.time_targets && Object.keys(bid.time_targets).length > 0) {
          timeTargeting = 'Yes';
        } else if (bid.timeTargets && Object.keys(bid.timeTargets).length > 0) {
          timeTargeting = 'Yes';
        } else if (bid.time_bids && Object.keys(bid.time_bids).length > 0) {
          timeTargeting = 'Yes';
        } else if (bid.timeBids && Object.keys(bid.timeBids).length > 0) {
          timeTargeting = 'Yes';
        } else {
          timeTargeting = 'No';
        }
        
        // Get stats
        const stats = bid.stats || {};
        
        bidRows.push({
          campaignId: campaignId,
          campaignName: campaignName,
          campaignType: campaignType,  // Format (banner, video, etc.)
          bidId: bid.bid_id || bidId,
          spotId: bid.spot_id || '',
          spotName: spotName,
          device: device,
          os: os,                        // OS from campaign name (Android/iOS)
          countries: countryDisplay,   // All countries (comma-separated)
          geoId: geoIdDisplay,         // Geo ID count or single ID
          timeTargeting: timeTargeting, // Has time targeting?
          currentBid: toNumeric(bid.bid, 0),
          isActive: bid.isActive ? 'Yes' : 'No',
          isPaused: bid.isPaused ? 'Yes' : 'No',
          // Stats from the stats object (per-bid)
          impressions: toNumeric(stats.impressions, 0),
          clicks: toNumeric(stats.clicks, 0),
          conversions: toNumeric(stats.conversions, 0),
          cost: toNumeric(stats.revenue, 0),  // API calls it "revenue" but it's your cost/spend
          // Per-bid eCPM (for campaign avg, use "Pull Campaign Stats" + VLOOKUP)
          ecpm: toNumeric(stats.ecpm, 0),
          ecpc: toNumeric(stats.ecpc, 0),
          ctr: toNumeric(stats.ctr, 0)
        });
      }
      
      // Small delay between campaigns
      if (campaignIds.length > 1) {
        Utilities.sleep(200);
      }
    }
    
    Logger.log(`\n✅ Complete: ${bidRows.length} total bid rows`);
    
    // Write to sheet
    writeBidsToSheet(bidRows);
    
    ui.alert('Success', 
      `Pulled ${bidRows.length} bid entries for ${campaignIds.length} campaign(s) from Legend sheet.\n\n` +
      `Campaigns: ${campaignIds.join(', ')}\n\n` +
      `Check the "${BIDS_SHEET_NAME}" sheet.`, 
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    ui.alert('Error', `Failed to pull bids: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Write bid data to the sheet
 * @param {Array} bidRows - Array of bid objects
 */
function writeBidsToSheet(bidRows) {
  const sheet = getOrCreateSheet(BIDS_SHEET_NAME);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear existing data
  sheet.clear();
  
  // Define headers - reordered per user request
  const headers = [
    'Strategy',         // A - From Legend sheet (Column C)
    'Sub Strategy',     // B - From Legend sheet (Column D)
    'Keyword',          // C - From Legend sheet (Column E)
    'Format',           // D - From Legend sheet (Column F)
    'Country',          // E
    'Geo ID',           // F - Unique geo target ID (different IDs = different sub-regions)
    'Spot Name',        // G
    'Device',           // H - From spot_name (Mobile/Tablet/Desktop)
    'OS',               // I - From campaign name (Android/iOS)
    'Time Target',      // J - Has time targeting?
    'Your CPM',         // K (Current Bid)
    'New CPM',          // L (Editable)
    'Change %',         // M
    'Avg eCPM',         // N - Campaign Avg eCPM (VLOOKUP from Campaign Stats sheet)
    'Cost',             // O
    'Conversions',      // P
    'CPA',              // Q (Calculated: Cost / Conversions)
    'Impressions',      // R
    'Clicks',           // S
    'CTR',              // T
    'Bid Status',       // U (Active/Paused)
    'Last Updated',     // V
    'Campaign Name',    // W
    'Campaign ID',      // X
    'Spot ID',          // Y
    'Bid ID'            // Z
  ];
  
  // Write headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white')
    .setHorizontalAlignment('center');
  
  // Build lookup map from Legend sheet
  // Legend sheet: Column A = Campaign ID, Column C = Strategy, Column D = Sub Strategy, Column E = Keyword, Column F = Type/Format
  const legendLookup = {};
  const legendSheet = ss.getSheetByName('Legend');
  if (legendSheet) {
    const legendLastRow = legendSheet.getLastRow();
    if (legendLastRow >= 2) {
      // Get columns A through F (Campaign ID, ?, Strategy, Sub Strategy, Keyword, Type)
      const legendData = legendSheet.getRange(2, 1, legendLastRow - 1, 6).getValues();
      for (const legendRow of legendData) {
        const campaignId = String(legendRow[0]).trim();  // Column A - Campaign ID
        if (campaignId) {
          legendLookup[campaignId] = {
            strategy: legendRow[2] || '',      // Column C - Strategy
            subStrategy: legendRow[3] || '',   // Column D - Sub Strategy
            keyword: legendRow[4] || '',       // Column E - Keyword
            format: legendRow[5] || ''         // Column F - Type/Format
          };
        }
      }
      Logger.log(`Loaded ${Object.keys(legendLookup).length} entries from Legend sheet`);
    }
  } else {
    Logger.log('Legend sheet not found - Strategy/Sub Strategy/Keyword/Format columns will be empty');
  }
  
  // Prepare data rows - matching new column structure
  const now = new Date();
  const dataRows = bidRows.map(row => {
    // Determine bid status
    let bidStatus = 'Active';
    if (row.isPaused === 'Yes' || row.isPaused === true || row.isPaused === 1) {
      bidStatus = 'Paused';
    } else if (row.isActive === 'No' || row.isActive === false) {
      bidStatus = 'Inactive';
    }
    
    // Calculate CPA (Cost per Acquisition)
    const conversions = row.conversions || 0;
    const cost = row.cost || 0;
    const cpa = conversions > 0 ? cost / conversions : 0;
    
    // Look up Strategy/Sub Strategy/Keyword from Legend sheet
    const legend = legendLookup[String(row.campaignId)] || {};
    
    return [
      legend.strategy || '',                    // A - Strategy (from Legend)
      legend.subStrategy || '',                 // B - Sub Strategy (from Legend)
      legend.keyword || '',                     // C - Keyword (from Legend)
      legend.format || row.campaignType || '',  // D - Format (from Legend F, fallback to API)
      row.countries || '',                      // E - Country (all countries comma-separated)
      row.geoId || '',                          // F - Geo ID (count or single ID)
      row.spotName || '',                       // G - Spot Name
      row.device || '',                         // H - Device (from spot_name)
      row.os || '',                             // I - OS (from campaign name: Android/iOS)
      row.timeTargeting || '',                  // J - Time Targeting
      row.currentBid || 0,                      // K - Your CPM
      '',                                       // L - New CPM (empty for user to fill)
      '',                                       // M - Change % (will be formula)
      '',                                       // N - eCPM (VLOOKUP formula added below)
      cost,                                     // O - Cost
      row.conversions || 0,                     // P - Conversions
      cpa,                                      // Q - CPA (calculated)
      row.impressions || 0,                     // R - Impressions
      row.clicks || 0,                          // S - Clicks
      row.ctr || 0,                             // T - CTR
      bidStatus,                                // U - Bid Status
      now,                                      // V - Last Updated
      row.campaignName || '',                   // W - Campaign Name
      row.campaignId,                           // X - Campaign ID
      row.spotId || '',                         // Y - Spot ID
      row.bidId || ''                           // Z - Bid ID
    ];
  });
  
  if (dataRows.length > 0) {
    // Write data
    sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    
    // Add formula for Change % column (M/13) - compares New CPM (L/12) to Your CPM (K/11)
    for (let i = 2; i <= dataRows.length + 1; i++) {
      sheet.getRange(i, 13).setFormula(`=IF(AND(K${i}>0,L${i}<>""),((L${i}-K${i})/K${i})*100,"")`);
    }
    
    // Add VLOOKUP formula for eCPM column (N/14) - looks up Campaign ID (X) in Campaign Stats sheet
    // Returns Avg eCPM (column D) from Campaign Stats, or "Run Stats Pull" if not found
    for (let i = 2; i <= dataRows.length + 1; i++) {
      sheet.getRange(i, 14).setFormula(`=IFERROR(VLOOKUP(X${i},'Campaign Stats'!A:D,4,FALSE),"Run Stats Pull")`);
    }
    
    // Make Campaign Name (W/23) a clickable link to TJ campaign page
    for (let i = 0; i < dataRows.length; i++) {
      const rowNum = i + 2;  // Sheet row (1-based, after header)
      const campaignName = dataRows[i][22];  // Column W (index 22) - Campaign Name
      const campaignId = dataRows[i][23];    // Column X (index 23) - Campaign ID
      if (campaignName && campaignId) {
        const url = `https://advertiser.trafficjunky.com/campaign/${campaignId}/tracking-spots-rules`;
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(campaignName)
          .setLinkUrl(url)
          .build();
        sheet.getRange(rowNum, 23).setRichTextValue(richText);  // Column W
      }
    }
    
    // Format columns
    // Strategy/Sub Strategy/Keyword/Format (A-D) - from Legend sheet (light green to show populated)
    sheet.getRange(2, 1, dataRows.length, 4).setBackground('#e6f4ea');  // Light green
    
    // Your CPM (K/11) - currency
    sheet.getRange(2, 11, dataRows.length, 1).setNumberFormat('$#,##0.000');
    
    // New CPM (L/12) - currency (editable column - highlight it)
    sheet.getRange(2, 12, dataRows.length, 1)
      .setNumberFormat('$#,##0.000')
      .setBackground('#fff9c4'); // Light yellow to indicate editable
    
    // Change % (M/13) - percentage
    sheet.getRange(2, 13, dataRows.length, 1).setNumberFormat('0.00"%"');
    
    // eCPM (N/14) - currency
    sheet.getRange(2, 14, dataRows.length, 1).setNumberFormat('$#,##0.000');
    
    // Cost (O/15) - currency
    sheet.getRange(2, 15, dataRows.length, 1).setNumberFormat('$#,##0.00');
    
    // Conversions (P/16) - number
    sheet.getRange(2, 16, dataRows.length, 1).setNumberFormat('#,##0');
    
    // CPA (Q/17) - currency
    sheet.getRange(2, 17, dataRows.length, 1).setNumberFormat('$#,##0.00');
    
    // Impressions (R/18) - number
    sheet.getRange(2, 18, dataRows.length, 1).setNumberFormat('#,##0');
    
    // Clicks (S/19) - number
    sheet.getRange(2, 19, dataRows.length, 1).setNumberFormat('#,##0');
    
    // CTR (T/20) - percentage
    sheet.getRange(2, 20, dataRows.length, 1).setNumberFormat('0.00"%"');
    
    // Last Updated (V/22) - date/time
    sheet.getRange(2, 22, dataRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    
    // Campaign ID (X/24) - plain number (no formatting)
    sheet.getRange(2, 24, dataRows.length, 1).setNumberFormat('0');
    
    // Spot ID (Y/25) - plain number
    sheet.getRange(2, 25, dataRows.length, 1).setNumberFormat('0');
    
    // Bid ID (Z/26) - plain number
    sheet.getRange(2, 26, dataRows.length, 1).setNumberFormat('0');
    
    // Bid Status column (U/21) - conditional formatting
    const statusRange = sheet.getRange(2, 21, dataRows.length, 1);
    
    const rules = sheet.getConditionalFormatRules();
    
    // Active -> green
    const activeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Active')
      .setBackground('#c6efce')
      .setFontColor('#006100')
      .setRanges([statusRange])
      .build();
    
    // Paused -> yellow  
    const pausedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Paused')
      .setBackground('#ffeb9c')
      .setFontColor('#9c5700')
      .setRanges([statusRange])
      .build();
    
    // Inactive -> red
    const inactiveRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Inactive')
      .setBackground('#ffc7ce')
      .setFontColor('#9c0006')
      .setRanges([statusRange])
      .build();
    
    rules.push(activeRule, pausedRule, inactiveRule);
    sheet.setConditionalFormatRules(rules);
  }
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);

  // Make Spot Name column wider
  sheet.setColumnWidth(7, 200);  // G - Spot Name

  // Make Campaign Name column wider
  sheet.setColumnWidth(23, 250);  // W - Campaign Name
  
  // Activate the sheet
  ss.setActiveSheet(sheet);
  
  Logger.log(`Wrote ${dataRows.length} rows to ${BIDS_SHEET_NAME}`);
}

/**
 * Copy current bids to the New Bid column
 */
function copyBidsToNewColumn() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BIDS_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Error', `Sheet "${BIDS_SHEET_NAME}" not found. Please pull bids first.`, ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('Error', 'No bid data found. Please pull bids first.', ui.ButtonSet.OK);
    return;
  }
  
  // Copy Your CPM (column K/11) to New CPM (column L/12)
  const currentBids = sheet.getRange(2, 11, lastRow - 1, 1).getValues();
  sheet.getRange(2, 12, lastRow - 1, 1).setValues(currentBids);
  
  ui.alert('Done', `Copied ${lastRow - 1} bid values to "New CPM" column. You can now adjust the values.`, ui.ButtonSet.OK);
}

/**
 * Calculate and display bid change summary
 */
function calculateBidChanges() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BIDS_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Error', `Sheet "${BIDS_SHEET_NAME}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('Error', 'No bid data found.', ui.ButtonSet.OK);
    return;
  }
  
  // Get bid data - columns A through X (24 columns to get Campaign ID in X at index 23)
  const data = sheet.getRange(2, 1, lastRow - 1, 24).getValues();
  
  let totalChanges = 0;
  let increasedCount = 0;
  let decreasedCount = 0;
  let unchangedCount = 0;
  let totalCurrentBid = 0;
  let totalNewBid = 0;
  
  const changedCampaigns = new Set();
  
  for (const row of data) {
    const currentBid = toNumeric(row[10], 0);  // Column K (index 10) - Your CPM
    const newBid = toNumeric(row[11], 0);     // Column L (index 11) - New CPM
    const campaignIdForCalc = row[23];        // Column X (index 23) - Campaign ID
    
    totalCurrentBid += currentBid;
    
    if (newBid > 0) {
      totalNewBid += newBid;
      
      if (newBid > currentBid) {
        increasedCount++;
        changedCampaigns.add(campaignIdForCalc);
      } else if (newBid < currentBid) {
        decreasedCount++;
        changedCampaigns.add(campaignIdForCalc);
      } else {
        unchangedCount++;
      }
      
      totalChanges++;
    }
  }
  
  const summary = [
    `📊 BID CHANGE SUMMARY`,
    ``,
    `Total entries with new bids: ${totalChanges}`,
    ``,
    `📈 Increased: ${increasedCount}`,
    `📉 Decreased: ${decreasedCount}`,
    `➡️ Unchanged: ${unchangedCount}`,
    ``,
    `Campaigns affected: ${changedCampaigns.size}`,
    ``,
    `Total current bid value: $${totalCurrentBid.toFixed(3)}`,
    `Total new bid value: $${totalNewBid.toFixed(3)}`,
    `Net change: $${(totalNewBid - totalCurrentBid).toFixed(3)}`
  ].join('\n');
  
  ui.alert('Bid Change Summary', summary, ui.ButtonSet.OK);
}

/**
 * Clear all bid data from the sheet
 */
function clearBidData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BIDS_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Info', `Sheet "${BIDS_SHEET_NAME}" not found. Nothing to clear.`, ui.ButtonSet.OK);
    return;
  }
  
  const confirm = ui.alert('Confirm', 'This will clear all bid data. Are you sure?', ui.ButtonSet.YES_NO);
  
  if (confirm === ui.Button.YES) {
    sheet.clear();
    ui.alert('Done', 'Bid data cleared.', ui.ButtonSet.OK);
  }
}

// ============================================================================
// DEBUG FUNCTION - To troubleshoot API issues
// ============================================================================

/**
 * Debug function to test the /api/bids/{campaignId}.json endpoint (with full details)
 */
function debugAPIResponse() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Get campaign IDs from Legend sheet
    let campaignIds = [];
    const legendSheet = ss.getSheetByName('Legend');
    if (legendSheet) {
      const legendLastRow = legendSheet.getLastRow();
      if (legendLastRow >= 2) {
        const legendCampaignIds = legendSheet.getRange(2, 1, legendLastRow - 1, 1).getValues();
        campaignIds = legendCampaignIds
          .map(row => String(row[0]).trim())
          .filter(id => id && id !== '');
      }
    }
    if (campaignIds.length === 0) {
      campaignIds = CAMPAIGN_IDS_FALLBACK;
    }
    
    Logger.log(`=== DEBUG API RESPONSE ===`);
    Logger.log(`Testing /api/bids/{campaignId}.json endpoint (FULL DETAILS)`);
    Logger.log(`Campaign IDs from Legend: ${campaignIds.join(', ')}`);
    
    let summaryText = '';
    
    for (const campaignId of campaignIds) {
      Logger.log(`\n========================================`);
      Logger.log(`Testing campaign: ${campaignId}`);
      Logger.log(`========================================`);
      
      // Use /api/bids/{campaignId}.json (NOT /active.json) for full details
      const url = `${API_BASE_URL}/bids/${campaignId}.json?api_key=${API_KEY}`;
      Logger.log(`URL: ${url.replace(API_KEY, 'HIDDEN')}`);
      
      const response = UrlFetchApp.fetch(url, {
        'method': 'get',
        'contentType': 'application/json',
        'muteHttpExceptions': true
      });
      
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      Logger.log(`Response code: ${responseCode}`);
      Logger.log(`Response length: ${responseText.length} chars`);
      
      summaryText += `\nCampaign ${campaignId}: `;
      
      if (responseCode !== 200) {
        Logger.log(`ERROR: ${responseText}`);
        summaryText += `ERROR ${responseCode}`;
        continue;
      }
      
      const data = JSON.parse(responseText);
      
      // Log response structure
      Logger.log(`\n--- RESPONSE STRUCTURE ---`);
      Logger.log(`Type: ${typeof data}, isArray: ${Array.isArray(data)}`);
      
      if (Array.isArray(data)) {
        Logger.log(`Array length: ${data.length} bids`);
        summaryText += `✅ ${data.length} bids found`;
        
        if (data.length > 0) {
          Logger.log(`\n--- FIRST BID ALL FIELDS ---`);
          const firstBid = data[0];
          const fields = Object.keys(firstBid);
          Logger.log(`Fields (${fields.length}): ${fields.join(', ')}`);
          
          // Log each field with its value
          for (const field of fields) {
            const value = firstBid[field];
            if (typeof value === 'object' && value !== null) {
              Logger.log(`  ${field}: ${JSON.stringify(value)}`);
            } else {
              Logger.log(`  ${field}: ${value}`);
            }
          }
          
          // Show ALL bids
          Logger.log(`\n--- ALL BIDS ---`);
          data.forEach((bid, i) => {
            Logger.log(`\nBid ${i + 1}:`);
            Logger.log(`  bid_id: ${bid.bid_id || bid.id || 'N/A'}`);
            Logger.log(`  spot_id: ${bid.spot_id || bid.spotId || 'N/A'}`);
            Logger.log(`  spot_name: ${bid.spot_name || bid.spotName || 'N/A'}`);
            Logger.log(`  bid amount: ${bid.bid || bid.cpm || 'N/A'}`);
            Logger.log(`  geos: ${JSON.stringify(bid.geos || bid.geo || 'N/A')}`);
            // Log any other interesting fields
            for (const key of Object.keys(bid)) {
              if (!['bid_id', 'id', 'spot_id', 'spotId', 'spot_name', 'spotName', 'bid', 'cpm', 'geos', 'geo'].includes(key)) {
                Logger.log(`  ${key}: ${JSON.stringify(bid[key])}`);
              }
            }
          });
        }
      } else if (typeof data === 'object' && data !== null) {
        const keys = Object.keys(data);
        Logger.log(`Object keys: ${keys.join(', ')}`);
        Logger.log(`Full response: ${JSON.stringify(data, null, 2).substring(0, 5000)}`);
        summaryText += `Object with ${keys.length} keys`;
        
        // If it's an object, might have bids nested
        if (data.bids) {
          Logger.log(`\n--- BIDS ARRAY ---`);
          Logger.log(JSON.stringify(data.bids, null, 2));
        }
      }
    }
    
    // Show summary
    const summary = [
      `=== DEBUG RESULTS ===`,
      ``,
      `Endpoint: /api/bids/{campaignId}.json`,
      ``,
      summaryText,
      ``,
      `Check Execution Log for FULL field details.`
    ].join('\n');
    
    ui.alert('Debug Results', summary, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Debug Error: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    ui.alert('Debug Error', error.toString(), ui.ButtonSet.OK);
  }
}

// ============================================================================
// UPDATE BIDS IN TRAFFICJUNKY
// ============================================================================

/**
 * Update bids in TrafficJunky using PUT /api/bids/{bidId}/set.json
 */
function updateBidsInTJ() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BIDS_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Error', `Sheet "${BIDS_SHEET_NAME}" not found. Please pull bids first.`, ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('Error', 'No bid data found. Please pull bids first.', ui.ButtonSet.OK);
    return;
  }
  
  // Get bid data - all columns (A through Z = 26 columns)
  const data = sheet.getRange(2, 1, lastRow - 1, 26).getValues();

  // Find bids that need to be updated (New CPM is filled and different from Your CPM)
  // Column positions after adding OS column:
  // E (4): Country, F (5): Geo ID, G (6): Spot Name, H (7): Device, I (8): OS
  // K (10): Your CPM, L (11): New CPM
  // W (22): Campaign Name, X (23): Campaign ID, Y (24): Spot ID, Z (25): Bid ID
  const bidsToUpdate = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const country = row[4];                    // Column E - Country (index 4)
    const spotName = row[6];                   // Column G - Spot Name (index 6)
    const device = row[7];                     // Column H - Device (index 7)
    const os = row[8];                         // Column I - OS (index 8)
    const currentBid = toNumeric(row[10], 0);  // Column K - Your CPM (index 10)
    const newBid = toNumeric(row[11], 0);      // Column L - New CPM (index 11)
    const campaignName = row[22];              // Column W - Campaign Name (index 22)
    const campaignId = row[23];                // Column X - Campaign ID (index 23)
    const spotId = row[24];                    // Column Y - Spot ID (index 24)
    const bidId = row[25];                     // Column Z - Bid ID (index 25)

    // Skip if no bid ID or New CPM is empty
    if (!bidId || newBid === 0) continue;

    // Skip if no change
    if (newBid === currentBid) continue;

    bidsToUpdate.push({
      rowIndex: i + 2,  // Sheet row (1-based, after header)
      campaignId: String(campaignId),
      campaignName: campaignName,
      bidId: String(bidId),
      spotId: String(spotId),
      spotName: spotName,
      device: device,
      country: country,
      currentBid: currentBid,
      newBid: newBid,
      change: ((newBid - currentBid) / currentBid * 100).toFixed(2)
    });
  }
  
  if (bidsToUpdate.length === 0) {
    ui.alert('No Changes', 'No bids to update. Fill in the "New CPM" column with different values than "Your CPM".', ui.ButtonSet.OK);
    return;
  }
  
  // Show confirmation dialog
  let confirmMsg = `⚠️ CONFIRM BID UPDATES ⚠️\n\n`;
  confirmMsg += `You are about to update ${bidsToUpdate.length} bid(s) in TrafficJunky:\n\n`;
  
  bidsToUpdate.slice(0, 10).forEach((bid, i) => {
    const direction = bid.newBid > bid.currentBid ? '📈' : '📉';
    confirmMsg += `${i + 1}. ${bid.spotName}\n`;
    confirmMsg += `   $${bid.currentBid.toFixed(3)} → $${bid.newBid.toFixed(3)} (${bid.change}%) ${direction}\n`;
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
  Logger.log(`\n=== UPDATING ${bidsToUpdate.length} BIDS IN TRAFFICJUNKY ===`);
  
  let successCount = 0;
  let failCount = 0;
  const results = [];
  const logEntries = [];  // For bid log sheet
  
  for (const bid of bidsToUpdate) {
    Logger.log(`\nUpdating bid ${bid.bidId}: $${bid.currentBid} → $${bid.newBid}`);
    
    const timestamp = new Date();
    
    try {
      const url = `${API_BASE_URL}/bids/${bid.bidId}/set.json?api_key=${API_KEY}`;
      
      const response = UrlFetchApp.fetch(url, {
        'method': 'put',
        'contentType': 'application/json',
        'payload': JSON.stringify({ bid: bid.newBid.toString() }),
        'muteHttpExceptions': true
      });
      
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      Logger.log(`Response: ${responseCode} - ${responseText}`);
      
      if (responseCode === 200) {
        const result = JSON.parse(responseText);
        successCount++;
        results.push({
          bidId: bid.bidId,
          spotName: bid.spotName,
          status: 'SUCCESS',
          newBid: result.bid
        });
        
        // Add to log entries
        logEntries.push({
          timestamp: timestamp,
          campaignId: bid.campaignId,
          campaignName: bid.campaignName,
          bidId: bid.bidId,
          spotId: bid.spotId,
          spotName: bid.spotName,
          device: bid.device,
          country: bid.country,
          oldCpm: bid.currentBid,
          newCpm: result.bid,
          changePercent: bid.change,
          status: 'SUCCESS',
          error: ''
        });
        
        // Update the "Your CPM" column in the sheet to reflect the new value
        sheet.getRange(bid.rowIndex, 11).setValue(bid.newBid);  // Column K (Your CPM)
        // Clear the "New CPM" column
        sheet.getRange(bid.rowIndex, 12).setValue('');  // Column L (New CPM)
        
      } else {
        failCount++;
        results.push({
          bidId: bid.bidId,
          spotName: bid.spotName,
          status: 'FAILED',
          error: responseText.substring(0, 100)
        });
        
        // Log failed attempts too
        logEntries.push({
          timestamp: timestamp,
          campaignId: bid.campaignId,
          campaignName: bid.campaignName,
          bidId: bid.bidId,
          spotId: bid.spotId,
          spotName: bid.spotName,
          device: bid.device,
          country: bid.country,
          oldCpm: bid.currentBid,
          newCpm: bid.newBid,
          changePercent: bid.change,
          status: 'FAILED',
          error: responseText.substring(0, 200)
        });
      }
      
      // Small delay between API calls to avoid rate limiting
      Utilities.sleep(200);
      
    } catch (error) {
      Logger.log(`Error: ${error.toString()}`);
      failCount++;
      results.push({
        bidId: bid.bidId,
        spotName: bid.spotName,
        status: 'ERROR',
        error: error.toString()
      });
      
      // Log errors too
      logEntries.push({
        timestamp: timestamp,
        campaignId: bid.campaignId,
        campaignName: bid.campaignName,
        bidId: bid.bidId,
        spotId: bid.spotId,
        spotName: bid.spotName,
        device: bid.device,
        country: bid.country,
        oldCpm: bid.currentBid,
        newCpm: bid.newBid,
        changePercent: bid.change,
        status: 'ERROR',
        error: error.toString().substring(0, 200)
      });
    }
  }
  
  // Write to Bid Logs sheet
  if (logEntries.length > 0) {
    writeBidLogs(logEntries);
  }
  
  // Show results
  let resultMsg = `✅ Bid Update Complete!\n\n`;
  resultMsg += `Successful: ${successCount}\n`;
  resultMsg += `Failed: ${failCount}\n\n`;
  
  if (successCount > 0) {
    resultMsg += `Updated bids:\n`;
    results.filter(r => r.status === 'SUCCESS').slice(0, 10).forEach(r => {
      resultMsg += `• ${r.spotName}: $${r.newBid}\n`;
    });
  }
  
  if (failCount > 0) {
    resultMsg += `\nFailed bids:\n`;
    results.filter(r => r.status !== 'SUCCESS').forEach(r => {
      resultMsg += `• ${r.spotName}: ${r.error}\n`;
    });
  }
  
  resultMsg += `\nThe "Your CPM" column has been updated with the new values.`;
  
  ui.alert('Update Results', resultMsg, ui.ButtonSet.OK);
  
  Logger.log(`\n=== UPDATE COMPLETE: ${successCount} success, ${failCount} failed ===`);
}

// ============================================================================
// BID LOGGING
// ============================================================================

const BID_LOGS_SHEET_NAME = "Bid Logs";

/**
 * Write bid update logs to the Bid Logs sheet
 * @param {Array} logEntries - Array of log entry objects
 */
function writeBidLogs(logEntries) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(BID_LOGS_SHEET_NAME);
  
  // Create sheet if it doesn't exist
  if (!logSheet) {
    logSheet = ss.insertSheet(BID_LOGS_SHEET_NAME);
    
    // Set up headers
    const headers = [
      'Timestamp',        // A
      'Campaign ID',      // B
      'Campaign Name',    // C
      'Bid ID',           // D
      'Spot ID',          // E
      'Spot Name',        // F
      'Device',           // G
      'Country',          // H
      'Old CPM',          // I
      'New CPM',          // J
      'Change %',         // K
      'Status',           // L
      'Error'             // M
    ];
    
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    logSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('white')
      .setHorizontalAlignment('center');
    
    // Freeze header row
    logSheet.setFrozenRows(1);
    
    Logger.log(`Created new sheet: ${BID_LOGS_SHEET_NAME}`);
  }
  
  // Prepare log rows
  const logRows = logEntries.map(entry => [
    entry.timestamp,
    entry.campaignId,
    entry.campaignName,
    entry.bidId,
    entry.spotId,
    entry.spotName,
    entry.device,
    entry.country,
    entry.oldCpm,
    entry.newCpm,
    entry.changePercent,
    entry.status,
    entry.error
  ]);
  
  // Find the last row with data
  const lastRow = logSheet.getLastRow();
  
  // Append new log entries
  if (logRows.length > 0) {
    logSheet.getRange(lastRow + 1, 1, logRows.length, logRows[0].length).setValues(logRows);
    
    // Format the new rows
    const newRowStart = lastRow + 1;
    
    // Timestamp format
    logSheet.getRange(newRowStart, 1, logRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    
    // CPM columns - currency format
    logSheet.getRange(newRowStart, 9, logRows.length, 1).setNumberFormat('$#,##0.000');  // Old CPM
    logSheet.getRange(newRowStart, 10, logRows.length, 1).setNumberFormat('$#,##0.000'); // New CPM
    
    // Change % format
    logSheet.getRange(newRowStart, 11, logRows.length, 1).setNumberFormat('0.00"%"');
    
    // Conditional formatting for status column
    const statusRange = logSheet.getRange(newRowStart, 12, logRows.length, 1);
    const rules = logSheet.getConditionalFormatRules();
    
    // SUCCESS = green
    const successRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('SUCCESS')
      .setBackground('#c6efce')
      .setFontColor('#006100')
      .setRanges([statusRange])
      .build();
    
    // FAILED = red
    const failedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('FAILED')
      .setBackground('#ffc7ce')
      .setFontColor('#9c0006')
      .setRanges([statusRange])
      .build();
    
    // ERROR = red
    const errorRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('ERROR')
      .setBackground('#ffc7ce')
      .setFontColor('#9c0006')
      .setRanges([statusRange])
      .build();
    
    rules.push(successRule, failedRule, errorRule);
    logSheet.setConditionalFormatRules(rules);
    
    Logger.log(`Wrote ${logRows.length} entries to ${BID_LOGS_SHEET_NAME}`);
  }
  
  // Auto-resize columns
  logSheet.autoResizeColumns(1, 13);
}

/**
 * Clear all bid logs
 */
function clearBidLogs() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(BID_LOGS_SHEET_NAME);
  
  if (!logSheet) {
    ui.alert('Info', 'No bid logs to clear.', ui.ButtonSet.OK);
    return;
  }
  
  const confirm = ui.alert('Confirm', 'This will delete ALL bid log history. Are you sure?', ui.ButtonSet.YES_NO);
  
  if (confirm === ui.Button.YES) {
    // Keep the header row, clear everything else
    const lastRow = logSheet.getLastRow();
    if (lastRow > 1) {
      logSheet.getRange(2, 1, lastRow - 1, 13).clearContent();
    }
    ui.alert('Done', 'Bid logs cleared.', ui.ButtonSet.OK);
  }
}

// ============================================================================
// API EXPLORATION - Find endpoints with full bid details
// ============================================================================

/**
 * Test multiple API endpoints to find one with full bid details
 * (Site, Source ID, Source Name, Position, Min CPM, Suggested CPM)
 */
function exploreAPIEndpoints() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get first campaign ID from Legend sheet
  let campaignId = CAMPAIGN_IDS_FALLBACK[0];
  const legendSheet = ss.getSheetByName('Legend');
  if (legendSheet && legendSheet.getLastRow() >= 2) {
    campaignId = String(legendSheet.getRange(2, 1).getValue()).trim() || campaignId;
  }
  
  // Get a bid_id from the active bids endpoint
  const activeBidsUrl = `${API_BASE_URL}/bids/${campaignId}/active.json?api_key=${API_KEY}`;
  const activeBidsResponse = UrlFetchApp.fetch(activeBidsUrl, { muteHttpExceptions: true });
  const activeBids = JSON.parse(activeBidsResponse.getContentText());
  const bidId = activeBids[0]?.bid_id || '1201505001';
  
  Logger.log(`=== EXPLORING API ENDPOINTS ===`);
  Logger.log(`Campaign ID: ${campaignId}`);
  Logger.log(`Sample Bid ID: ${bidId}`);
  
  // List of endpoints to test
  const endpointsToTest = [
    // Bid detail endpoints
    `/bids/${bidId}.json`,
    `/bids/${bidId}/details.json`,
    `/bid/${bidId}.json`,
    
    // Campaign bids with different formats
    `/campaigns/${campaignId}/bids.json`,
    `/campaigns/${campaignId}/bids/details.json`,
    `/campaigns/${campaignId}/placements.json`,
    `/campaigns/${campaignId}/spots.json`,
    `/campaigns/${campaignId}/sources.json`,
    
    // Placement/source endpoints
    `/placements.json`,
    `/placements/${campaignId}.json`,
    `/spots.json`,
    `/sources.json`,
    
    // Campaign details
    `/campaigns/${campaignId}.json`,
    `/campaigns/${campaignId}/details.json`,
    `/campaign/${campaignId}.json`,
    
    // Bid management endpoints
    `/bids/campaign/${campaignId}.json`,
    `/bids/list/${campaignId}.json`,
  ];
  
  const results = [];
  
  for (const endpoint of endpointsToTest) {
    const url = `${API_BASE_URL}${endpoint}?api_key=${API_KEY}`;
    Logger.log(`\n--- Testing: ${endpoint} ---`);
    
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const code = response.getResponseCode();
      const text = response.getContentText();
      
      Logger.log(`Status: ${code}`);
      
      if (code === 200) {
        try {
          const data = JSON.parse(text);
          let fieldCount = 0;
          let fields = [];
          
          if (Array.isArray(data) && data.length > 0) {
            fields = Object.keys(data[0]);
            fieldCount = fields.length;
          } else if (typeof data === 'object' && data !== null) {
            fields = Object.keys(data);
            fieldCount = fields.length;
          }
          
          Logger.log(`✅ SUCCESS - ${fieldCount} fields: ${fields.slice(0, 15).join(', ')}`);
          
          // Check if it has the fields we need
          const hasGoodFields = fields.some(f => 
            f.toLowerCase().includes('source') || 
            f.toLowerCase().includes('site') || 
            f.toLowerCase().includes('position') ||
            f.toLowerCase().includes('suggested') ||
            f.toLowerCase().includes('min')
          );
          
          if (hasGoodFields) {
            Logger.log(`⭐ HAS USEFUL FIELDS!`);
            Logger.log(`Full response: ${JSON.stringify(data, null, 2).substring(0, 2000)}`);
          }
          
          results.push({ endpoint, status: 'SUCCESS', fields: fieldCount, hasGoodFields });
        } catch (e) {
          Logger.log(`Response (not JSON): ${text.substring(0, 200)}`);
          results.push({ endpoint, status: 'NOT_JSON' });
        }
      } else if (code === 404) {
        Logger.log(`❌ Not found`);
        results.push({ endpoint, status: '404' });
      } else {
        Logger.log(`❌ Error ${code}: ${text.substring(0, 100)}`);
        results.push({ endpoint, status: `ERROR_${code}` });
      }
    } catch (e) {
      Logger.log(`❌ Exception: ${e.toString()}`);
      results.push({ endpoint, status: 'EXCEPTION' });
    }
    
    // Small delay
    Utilities.sleep(100);
  }
  
  // Summary
  Logger.log(`\n=== SUMMARY ===`);
  const successful = results.filter(r => r.status === 'SUCCESS');
  const withGoodFields = results.filter(r => r.hasGoodFields);
  
  Logger.log(`Endpoints tested: ${results.length}`);
  Logger.log(`Successful: ${successful.length}`);
  Logger.log(`With useful fields: ${withGoodFields.length}`);
  
  if (withGoodFields.length > 0) {
    Logger.log(`\n⭐ PROMISING ENDPOINTS:`);
    withGoodFields.forEach(r => Logger.log(`  ${r.endpoint} (${r.fields} fields)`));
  }
  
  let summaryText = `Tested ${results.length} endpoints\n\n`;
  summaryText += `✅ Successful: ${successful.length}\n`;
  summaryText += `⭐ With useful fields: ${withGoodFields.length}\n\n`;
  
  if (successful.length > 0) {
    summaryText += `Working endpoints:\n`;
    successful.forEach(r => {
      summaryText += `• ${r.endpoint} (${r.fields} fields)${r.hasGoodFields ? ' ⭐' : ''}\n`;
    });
  }
  
  summaryText += `\nCheck Execution Log for full details.`;
  
  ui.alert('API Exploration Results', summaryText, ui.ButtonSet.OK);
}

// ============================================================================
// CAMPAIGN SEARCH FUNCTIONS
// ============================================================================

/**
 * List all campaign IDs from the API to a new sheet
 */
function listAllCampaignIds() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Calculate date range - use last 30 days
    const endDate = getESTDate();
    endDate.setDate(endDate.getDate() - 1);
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 30);
    
    const formattedStartDate = formatDateForAPI(startDate);
    const formattedEndDate = formatDateForAPI(endDate);
    
    Logger.log(`Fetching all campaigns to list IDs...`);
    
    const listUrl = `${API_URL_CAMPAIGNS_BIDS}?api_key=${API_KEY}&startDate=${formattedStartDate}&endDate=${formattedEndDate}&limit=500&offset=1`;
    const bidsResponse = UrlFetchApp.fetch(listUrl, {
      'method': 'get',
      'contentType': 'application/json',
      'muteHttpExceptions': true
    });
    
    const bidsData = JSON.parse(bidsResponse.getContentText());
    
    let campaigns = [];
    if (Array.isArray(bidsData)) {
      campaigns = bidsData;
    } else if (typeof bidsData === 'object' && bidsData !== null) {
      campaigns = Object.values(bidsData);
    }
    
    // Create or get sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Campaign ID List');
    if (!sheet) {
      sheet = ss.insertSheet('Campaign ID List');
    }
    sheet.clear();
    
    // Headers
    const headers = ['Campaign ID', 'Campaign Name', 'Status', 'Has Bids', 'Num Bids'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    
    // Data
    const rows = campaigns.map(c => [
      c.campaignId || c.id || '',
      c.campaignName || '',
      c.status || '',
      c.bids && c.bids.length > 0 ? 'YES' : 'NO',
      c.bids ? c.bids.length : 0
    ]);
    
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    
    sheet.autoResizeColumns(1, headers.length);
    ss.setActiveSheet(sheet);
    
    ui.alert('Campaign List', 
      `Found ${campaigns.length} campaigns.\n\n` +
      `Check the "Campaign ID List" sheet.\n\n` +
      `Use Ctrl+F to search for your campaign ID: 1013232471`, 
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    ui.alert('Error', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Search for a campaign by name (partial match)
 */
function searchCampaignByName() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt('Search Campaign', 'Enter campaign name or partial name to search:', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() !== ui.Button.OK) return;
  
  const searchTerm = result.getResponseText().toLowerCase().trim();
  if (!searchTerm) {
    ui.alert('Error', 'Please enter a search term.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // Calculate date range
    const endDate = getESTDate();
    endDate.setDate(endDate.getDate() - 1);
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 30);
    
    const formattedStartDate = formatDateForAPI(startDate);
    const formattedEndDate = formatDateForAPI(endDate);
    
    const listUrl = `${API_URL_CAMPAIGNS_BIDS}?api_key=${API_KEY}&startDate=${formattedStartDate}&endDate=${formattedEndDate}&limit=500&offset=1`;
    const bidsResponse = UrlFetchApp.fetch(listUrl, {
      'method': 'get',
      'contentType': 'application/json',
      'muteHttpExceptions': true
    });
    
    const bidsData = JSON.parse(bidsResponse.getContentText());
    
    let campaigns = [];
    if (Array.isArray(bidsData)) {
      campaigns = bidsData;
    } else if (typeof bidsData === 'object' && bidsData !== null) {
      campaigns = Object.values(bidsData);
    }
    
    // Search
    const matches = campaigns.filter(c => {
      const name = (c.campaignName || '').toLowerCase();
      const id = String(c.campaignId || c.id || '');
      return name.includes(searchTerm) || id.includes(searchTerm);
    });
    
    if (matches.length === 0) {
      ui.alert('No Results', `No campaigns found matching "${searchTerm}"`, ui.ButtonSet.OK);
      return;
    }
    
    // Show results
    let resultText = `Found ${matches.length} campaign(s) matching "${searchTerm}":\n\n`;
    matches.slice(0, 10).forEach(c => {
      resultText += `ID: ${c.campaignId || c.id}\n`;
      resultText += `Name: ${c.campaignName}\n`;
      resultText += `Status: ${c.status}\n`;
      resultText += `Bids: ${c.bids ? c.bids.length : 0}\n\n`;
    });
    
    if (matches.length > 10) {
      resultText += `... and ${matches.length - 10} more`;
    }
    
    ui.alert('Search Results', resultText, ui.ButtonSet.OK);
    
    // Log full details
    Logger.log(`\n=== SEARCH RESULTS for "${searchTerm}" ===`);
    matches.forEach(c => {
      Logger.log(`Campaign ID: ${c.campaignId || c.id}`);
      Logger.log(`Name: ${c.campaignName}`);
      Logger.log(`Status: ${c.status}`);
      Logger.log(`Bids: ${c.bids ? c.bids.length : 0}`);
      Logger.log('---');
    });
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    ui.alert('Error', error.toString(), ui.ButtonSet.OK);
  }
}

// ============================================================================
// EXPORT FUNCTION - For getting bid adjustments as JSON
// ============================================================================

/**
 * Export bid changes as JSON (useful for automation)
 * @returns {string} JSON string of bid changes
 */
function exportBidChangesAsJSON() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BIDS_SHEET_NAME);
  
  if (!sheet) {
    return JSON.stringify({ error: 'Sheet not found' });
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return JSON.stringify({ error: 'No data' });
  }
  
  // Get all columns (A through Z = 26 columns)
  const data = sheet.getRange(2, 1, lastRow - 1, 26).getValues();
  const changes = [];

  for (const row of data) {
    const currentBid = toNumeric(row[10], 0);  // Column K (index 10) - Your CPM
    const newBid = toNumeric(row[11], 0);     // Column L (index 11) - New CPM

    // Only include rows where New Bid is set and different from Current Bid
    if (newBid > 0 && newBid !== currentBid) {
      changes.push({
        strategy: row[0],                        // Column A
        subStrategy: row[1],                     // Column B
        keyword: row[2],                         // Column C
        format: row[3],                          // Column D
        country: row[4],                         // Column E
        geoId: row[5],                           // Column F
        spotName: row[6],                        // Column G
        device: row[7],                          // Column H
        os: row[8],                              // Column I - OS (Android/iOS)
        timeTarget: row[9],                      // Column J
        currentCpm: currentBid,                  // Column K
        newCpm: newBid,                          // Column L
        changePercent: ((newBid - currentBid) / currentBid * 100).toFixed(2),
        campaignName: row[22],                   // Column W
        campaignId: row[23],                     // Column X
        spotId: row[24],                         // Column Y
        bidId: row[25]                           // Column Z
      });
    }
  }
  
  return JSON.stringify({
    exportDate: new Date().toISOString(),
    totalChanges: changes.length,
    changes: changes
  }, null, 2);
}

/**
 * Show export dialog with JSON data
 */
function showExportDialog() {
  const json = exportBidChangesAsJSON();
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          textarea { width: 100%; height: 400px; font-family: monospace; font-size: 12px; }
          button { margin-top: 10px; padding: 10px 20px; cursor: pointer; }
        </style>
      </head>
      <body>
        <h3>Bid Changes Export</h3>
        <textarea id="json" readonly>${json}</textarea>
        <br>
        <button onclick="copyToClipboard()">Copy to Clipboard</button>
        <script>
          function copyToClipboard() {
            const textarea = document.getElementById('json');
            textarea.select();
            document.execCommand('copy');
            alert('Copied to clipboard!');
          }
        </script>
      </body>
    </html>
  `)
  .setWidth(600)
  .setHeight(550);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Export Bid Changes');
}

// ============================================================================
// PULL CAMPAIGN STATS - Separate sheet for campaign-level metrics
// ============================================================================

/**
 * Pull campaign stats from /api/campaigns/stats.json into a "Campaign Stats" sheet
 * Use VLOOKUP in Campaign Bids sheet to get Avg eCPM from here
 */
function pullCampaignStats() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create Campaign Stats sheet
  let statsSheet = ss.getSheetByName('Campaign Stats');
  if (!statsSheet) {
    statsSheet = ss.insertSheet('Campaign Stats');
  }
  statsSheet.clear();
  
  // Get today's date in EST (TrafficJunky timezone)
  const now = new Date();
  const estDateStr = Utilities.formatDate(now, API_TIMEZONE, 'dd/MM/yyyy');
  
  // Use EST date for both start and end (today only in EST)
  const startStr = estDateStr;
  const endStr = estDateStr;
  
  Logger.log(`Fetching campaign stats from ${startStr} to ${endStr}`);
  
  // Fetch with pagination
  const allStats = [];
  let offset = 1;
  let hasMoreData = true;
  const limit = 500;
  
  while (hasMoreData) {
    const url = `${API_URL_STATS}?api_key=${API_KEY}&startDate=${startStr}&endDate=${endStr}&limit=${limit}&offset=${offset}`;
    Logger.log(`Fetching offset=${offset}`);
    
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      
      if (response.getResponseCode() !== 200) {
        Logger.log(`Error: ${response.getResponseCode()} - ${response.getContentText()}`);
        break;
      }
      
      const data = JSON.parse(response.getContentText());
      
      let batchCount = 0;
      for (const [campaignId, stats] of Object.entries(data)) {
        // Handle both object and array formats
        const s = Array.isArray(stats) ? stats[0] : stats;
        if (s) {
          allStats.push({
            campaignId: s.campaign_id || campaignId,
            campaignName: s.campaign_name || '',
            campaignType: s.campaign_type || '',
            avgEcpm: toNumeric(s.ecpm, 0),
            avgEcpc: toNumeric(s.ecpc, 0),
            avgCtr: toNumeric(s.ctr, 0),
            totalCost: toNumeric(s.cost, 0),
            totalImpressions: toNumeric(s.impressions, 0),
            totalClicks: toNumeric(s.clicks, 0),
            totalConversions: toNumeric(s.conversions, 0),
            adsCount: s.ads_count || 0,
            adsPaused: s.ads_paused || 0
          });
          batchCount++;
        }
      }
      
      Logger.log(`Got ${batchCount} campaigns (Total: ${allStats.length})`);
      
      if (batchCount < limit) {
        hasMoreData = false;
      } else {
        offset += limit;
      }
      
      // Safety limit
      if (offset > 10000) {
        hasMoreData = false;
      }
      
    } catch (e) {
      Logger.log(`Error fetching stats: ${e}`);
      hasMoreData = false;
    }
  }
  
  if (allStats.length === 0) {
    ui.alert('No Data', 'No campaign stats found. Try adjusting the date range.', ui.ButtonSet.OK);
    return;
  }
  
  // Write headers
  const headers = [
    'Campaign ID',      // A - Use this for VLOOKUP
    'Campaign Name',    // B
    'Type',             // C
    'Avg eCPM',         // D - VLOOKUP target!
    'Avg eCPC',         // E
    'Avg CTR',          // F
    'Total Cost',       // G
    'Impressions',      // H
    'Clicks',           // I
    'Conversions',      // J
    'Ads Count',        // K
    'Ads Paused'        // L
  ];
  
  statsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  statsSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('white');
  
  // Write data
  const rows = allStats.map(s => [
    s.campaignId,
    s.campaignName,
    s.campaignType,
    s.avgEcpm,
    s.avgEcpc,
    s.avgCtr,
    s.totalCost,
    s.totalImpressions,
    s.totalClicks,
    s.totalConversions,
    s.adsCount,
    s.adsPaused
  ]);
  
  if (rows.length > 0) {
    statsSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    
    // Format columns
    statsSheet.getRange(2, 1, rows.length, 1).setNumberFormat('0');           // Campaign ID
    statsSheet.getRange(2, 4, rows.length, 1).setNumberFormat('$#,##0.000');  // Avg eCPM
    statsSheet.getRange(2, 5, rows.length, 1).setNumberFormat('$#,##0.000');  // Avg eCPC
    statsSheet.getRange(2, 6, rows.length, 1).setNumberFormat('0.00"%"');     // Avg CTR
    statsSheet.getRange(2, 7, rows.length, 1).setNumberFormat('$#,##0.00');   // Total Cost
    statsSheet.getRange(2, 8, rows.length, 1).setNumberFormat('#,##0');       // Impressions
    statsSheet.getRange(2, 9, rows.length, 1).setNumberFormat('#,##0');       // Clicks
    statsSheet.getRange(2, 10, rows.length, 1).setNumberFormat('#,##0');      // Conversions
    
    statsSheet.autoResizeColumns(1, headers.length);
    statsSheet.setColumnWidth(2, 300);  // Campaign Name wider
  }
  
  Logger.log(`✅ Wrote ${rows.length} campaigns to Campaign Stats sheet`);
  
  ui.alert('Campaign Stats Loaded', 
    `Loaded ${rows.length} campaigns to "Campaign Stats" sheet.\n\n` +
    `Date range: ${startStr} to ${endStr}\n\n` +
    `The Campaign Bids sheet will automatically look up Avg eCPM from here.`,
    ui.ButtonSet.OK);
}
