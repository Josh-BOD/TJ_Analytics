/**
 * TJ Control Panel V8 - Google Apps Script
 * 
 * A dashboard for bid optimization with multi-period stats (Today, Yesterday, 7-Day).
 * Uses BID-LEVEL stats for all periods - granular per-bid T/Y/7D comparisons.
 * 
 * V8 CHANGES:
 * - Advanced Filtering: Filter by any dimension (Strategy, Sub-Strategy, Campaign, Spot, Country)
 * - Multiple Filters: Combine up to 3 filters with AND logic
 * - Filter Modes: "equals" for exact match, "contains" for partial match
 * - Comparison Date Range: Compare two time periods (e.g., this week vs last week)
 * - Period-over-period analysis: See delta/change between periods
 * 
 * V7 FEATURES:
 * - Multi-level Dashboard: View data at Strategy/Sub-Strategy/Spot/Campaign/Country levels
 * - Spot-level daily stats: Pull bid-level stats per day for granular analysis
 * - Multi-select comparison: Compare up to 5 items with overlaid chart
 * - Breakdown tables: See child items with vs Yesterday comparison
 * 
 * V6 FEATURES:
 * - Native Google Sheets Pivot Table for analytics (proper grouping/subtotals)
 * - Lookup formulas adjacent to pivot to retrieve Bid IDs for editing
 * - Dashboard checkbox navigation from Pivot View
 * 
 * V4/V5 FEATURES:
 * - Pivot View sheet: Hierarchical view by Strategy > Sub Strategy > Campaign
 * - Edit columns at end of pivot for bid/budget changes
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
const CP_DAILY_SPOT_SHEET_NAME = "Daily Stats (Spot)";  // V7: Spot-level daily stats
const CP_DASHBOARD_SHEET_NAME = "Dashboard";           // V8 Dashboard
const CP_DASHBOARD_V7_SHEET_NAME = "Dashboard-V7";     // V7 Dashboard (separate to avoid conflicts)
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
    .addSubMenu(ui.createMenu('📊 Pivot Table (V8)')
      .addItem('1️⃣ Prepare Pivot Data Source', 'cpPreparePivotSource')
      .addItem('2️⃣ Add Edit Columns to Pivot', 'cpAddEditColumnsToPivot')
      .addItem('🔄 Refresh Edit Columns', 'cpRefreshPivotEditColumns')
      .addSeparator()
      .addItem('📋 Copy Bids (Pivot)', 'cpCopyBidsPivot')
      .addItem('📋 Copy Budgets (Pivot)', 'cpCopyBudgetsPivot')
      .addItem('🚀 UPDATE FROM PIVOT', 'cpUpdateFromPivot'))
    .addSubMenu(ui.createMenu('📊 Dashboard (V8)')
      .addItem('📥 Pull Daily Stats (Spot Level)', 'cpPullDailyStatsSpotLevel')
      .addSeparator()
      .addItem('📈 Build Dashboard V8', 'cpBuildDashboardV8')
      .addItem('🔄 Apply Filters', 'cpApplyDashboardFilters')
      .addItem('🔃 Refresh Dashboard', 'cpRefreshDashboardV8')
      .addSeparator()
      .addItem('📈 Build Dashboard V7 (Legacy)', 'cpBuildDashboardV7'))
    .addSeparator()
    .addItem('🗑️ Clear Data', 'cpClearData')
    .addSeparator()
    .addItem('🐛 Debug: Test Parallel Fetch', 'cpDebugParallelFetch')
    .addToUi();
}

/**
 * onEdit trigger handler - V7
 * 
 * Handles:
 * 1. Dashboard View Level changes (B2) - updates selection dropdowns
 * 2. Dashboard Selection changes - refreshes breakdown table
 * 3. Pivot View checkbox clicks - navigates to Dashboard with that campaign
 */
function onEdit(e) {
  // Only process if we have event info
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // =========================================================================
  // DASHBOARD V8: View Level, Selection, or Metric Toggle changes
  // =========================================================================
  if (sheetName === CP_DASHBOARD_SHEET_NAME) {
    try {
      // View Level changed (B2)
      if (row === 2 && col === 2) {
        const newViewLevel = e.value;
        cpUpdateDashboardDropdowns(sheet, newViewLevel);
        cpRefreshBreakdownTable(sheet);
        return;
      }
      
      // Selection changed (B3-F3, columns 2-6 for 5 selections)
      if (row === 3 && col >= 2 && col <= 6) {
        cpRefreshBreakdownTable(sheet);
        // Rebuild charts to reflect new selections - use V8 function
        cpRebuildDashboardChartsV8(sheet);
        return;
      }
      
      // Metric toggle checkbox changed - V8 uses row 13, columns B-H = 2-8
      if (row === 13 && col >= 2 && col <= 8) {
        cpRebuildDashboardChartsV8(sheet);
        return;
      }
    } catch (error) {
      Logger.log('Dashboard V8 onEdit error: ' + error.toString());
    }
    return;
  }
  
  // =========================================================================
  // DASHBOARD V7 (Legacy): View Level, Selection, or Metric Toggle changes
  // =========================================================================
  if (sheetName === CP_DASHBOARD_V7_SHEET_NAME) {
    try {
      // View Level changed (B2)
      if (row === 2 && col === 2) {
        const newViewLevel = e.value;
        cpUpdateDashboardDropdowns(sheet, newViewLevel);
        cpRefreshBreakdownTable(sheet);
        return;
      }
      
      // Selection changed (B3-F3, columns 2-6 for 5 selections)
      if (row === 3 && col >= 2 && col <= 6) {
        cpRefreshBreakdownTable(sheet);
        // V7 uses the V7 chart builder
        cpRebuildDashboardCharts(sheet);
        return;
      }
      
      // Metric toggle checkbox changed - V7 uses row 6, columns B-G = 2-7
      if (row === 6 && col >= 2 && col <= 7) {
        cpRebuildDashboardCharts(sheet);
        return;
      }
    } catch (error) {
      Logger.log('Dashboard V7 onEdit error: ' + error.toString());
    }
    return;
  }
  
  // =========================================================================
  // PIVOT VIEW: Checkbox clicks for Dashboard navigation
  // =========================================================================
  if (sheetName === CP_PIVOT_SHEET_NAME) {
    // Only process if checkbox was checked (value = true)
    const value = e.value;
    if (value !== 'TRUE' && value !== true) return;
    
    // Skip header row
    if (row < 2) return;
    
    try {
      // Find the checkbox column by looking for the 📊 header
      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      
      let checkboxCol = -1;
      let campIdCol = -1;
      let campaignNameCol = -1;
      
      for (let i = 0; i < headers.length; i++) {
        if (headers[i] === '📊') checkboxCol = i + 1;
        if (headers[i] === 'Campaign ID') campIdCol = i + 1;
        if (String(headers[i]).toLowerCase().includes('campaign') && 
            String(headers[i]).toLowerCase().includes('name')) campaignNameCol = i + 1;
      }
      
      // Only process if the edited cell is in the checkbox column
      if (col !== checkboxCol) return;
      
      // Get Campaign ID and Campaign Name from this row
      const campaignId = campIdCol > 0 ? sheet.getRange(row, campIdCol).getValue() : '';
      const campaignName = campaignNameCol > 0 ? sheet.getRange(row, campaignNameCol).getDisplayValue() : '';
      
      // If no Campaign ID (subtotal row), just uncheck and return
      if (!campaignId) {
        e.range.setValue(false);
        return;
      }
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      // Get Dashboard sheet
      const dashSheet = ss.getSheetByName(CP_DASHBOARD_SHEET_NAME);
      if (!dashSheet) {
        SpreadsheetApp.getUi().alert('Dashboard sheet not found. Please run "Build Dashboard V7" first.');
        e.range.setValue(false);
        return;
      }
      
      // Set View Level to Campaign and Selection 1 to the campaign name
      dashSheet.getRange('B2').setValue('Campaign');
      dashSheet.getRange('B3').setValue(campaignName);
      dashSheet.getRange('C3').setValue('');  // Clear other selections
      dashSheet.getRange('D3').setValue('');
      
      // Uncheck the checkbox
      e.range.setValue(false);
      
      // Navigate to Dashboard sheet
      ss.setActiveSheet(dashSheet);
      dashSheet.getRange('B3').activate();
      
    } catch (error) {
      Logger.log('Pivot onEdit error: ' + error.toString());
      // Uncheck the checkbox on error
      try {
        e.range.setValue(false);
      } catch (e2) {
        // Ignore
      }
    }
  }
}

/**
 * Update Dashboard dropdown options based on View Level
 */
function cpUpdateDashboardDropdowns(dashSheet, viewLevel) {
  const ss = dashSheet.getParent();
  
  // Get options from hidden sheet
  const optionsSheet = ss.getSheetByName('_DashboardOptions');
  if (!optionsSheet) {
    Logger.log('Options sheet not found - rebuild dashboard');
    return;
  }
  
  // Get all options
  const lastRow = optionsSheet.getLastRow();
  if (lastRow < 2) return;
  
  const allData = optionsSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  
  // Build options based on view level
  let options = [];
  switch (viewLevel) {
    case 'Tier 1 Strategy':
      options = allData.map(r => r[0]).filter(v => v);
      break;
    case 'Sub-Strategy':
      options = allData.map(r => r[1]).filter(v => v);
      break;
    case 'Spot Name':
      options = allData.map(r => r[2]).filter(v => v);
      break;
    case 'Campaign':
      options = allData.map(r => r[3]).filter(v => v);
      break;
    case 'Country':
      options = allData.map(r => r[4]).filter(v => v);
      break;
    case 'All':
    default:
      // Combine all
      for (let col = 0; col < 5; col++) {
        for (const row of allData) {
          if (row[col]) options.push(row[col]);
        }
      }
      options = [...new Set(options)].sort();
  }
  
  if (options.length === 0) {
    options = ['(No options available)'];
  }
  
  // Add "(Clear)" option at the start to allow clearing selection
  options = ['(Clear)', ...options];
  
  // Update dropdown validation for B3-F3 (5 selections)
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(true)
    .build();
  
  // Apply to all 5 selection cells
  for (let col = 2; col <= 6; col++) {  // B=2 through F=6
    dashSheet.getRange(3, col).setDataValidation(rule);
  }
  
  // Set first option in B3, clear others
  dashSheet.getRange('B3').setValue(options[1] || '');  // Skip "(Clear)", use first real option
  dashSheet.getRange('C3').setValue('');
  dashSheet.getRange('D3').setValue('');
  dashSheet.getRange('E3').setValue('');
  dashSheet.getRange('F3').setValue('');
  
  Logger.log(`Updated dropdowns for ${viewLevel}: ${options.length} options`);
}

/**
 * Refresh the breakdown table based on current selection
 */
function cpRefreshBreakdownTable(dashSheet) {
  const ss = dashSheet.getParent();
  const spotSheet = ss.getSheetByName(CP_DAILY_SPOT_SHEET_NAME);
  if (!spotSheet) return;
  
  const viewLevel = dashSheet.getRange('B2').getValue();
  const selection1 = dashSheet.getRange('B3').getValue();
  
  // Handle empty or "(Clear)" selection
  if (!selection1 || selection1 === '(Clear)') return;
  
  const breakdownStartRow = 36;  // Updated for V7 expanded layout + selection labels row
  
  // Clear existing breakdown data (keep headers)
  const lastBreakdownRow = dashSheet.getLastRow();
  if (lastBreakdownRow > breakdownStartRow + 1) {
    dashSheet.getRange(breakdownStartRow + 2, 1, lastBreakdownRow - breakdownStartRow - 1, 9).clear();
  }
  
  // Get spot data
  const spotLastRow = spotSheet.getLastRow();
  if (spotLastRow < 2) return;
  
  const spotData = spotSheet.getRange(2, 1, spotLastRow - 1, 16).getValues();
  
  // Determine which column to match and which column to show as breakdown
  // B=Strategy, C=Sub-Strategy, E=CampaignName, G=SpotName, H=Country
  let matchCol, breakdownCol, breakdownLabel;
  
  switch (viewLevel) {
    case 'Tier 1 Strategy':
      matchCol = 1;  // B - Strategy
      breakdownCol = 2;  // C - Sub-Strategy
      breakdownLabel = 'Sub-Strategy';
      break;
    case 'Sub-Strategy':
      matchCol = 2;  // C - Sub-Strategy
      breakdownCol = 4;  // E - Campaign Name
      breakdownLabel = 'Campaign';
      break;
    case 'Spot Name':
      matchCol = 6;  // G - Spot Name
      breakdownCol = 4;  // E - Campaign Name
      breakdownLabel = 'Campaign (using this spot)';
      break;
    case 'Campaign':
      matchCol = 4;  // E - Campaign Name (or "Campaign {ID}" fallback)
      breakdownCol = 6;  // G - Spot Name
      breakdownLabel = 'Spot';
      break;
    case 'Country':
      matchCol = 7;  // H - Country
      breakdownCol = 4;  // E - Campaign Name
      breakdownLabel = 'Campaign (targeting this country)';
      break;
    default:
      return;
  }
  
  // Aggregate data by breakdown item
  const breakdown = {};
  const yesterdayData = {};
  
  // Get yesterday's date for comparison
  const dates = [...new Set(spotData.map(r => r[0]))].sort();
  const latestDate = dates[dates.length - 1];
  const yesterdayDate = dates.length > 1 ? dates[dates.length - 2] : null;
  
  for (const row of spotData) {
    const matchValue = String(row[matchCol] || '');
    const breakdownValue = String(row[breakdownCol] || '');
    const date = row[0];
    
    // Check if this row matches our selection
    // For country, need to check if selection is in the comma-separated list
    let matches = false;
    if (viewLevel === 'Country') {
      const countries = matchValue.split(',').map(c => c.trim());
      matches = countries.includes(selection1);
    } else {
      matches = matchValue === selection1;
    }
    
    if (!matches || !breakdownValue) continue;
    
    // Aggregate
    if (!breakdown[breakdownValue]) {
      breakdown[breakdownValue] = {
        impressions: 0, clicks: 0, conversions: 0, spend: 0
      };
    }
    
    breakdown[breakdownValue].impressions += Number(row[9]) || 0;   // J
    breakdown[breakdownValue].clicks += Number(row[10]) || 0;       // K
    breakdown[breakdownValue].conversions += Number(row[11]) || 0;  // L
    breakdown[breakdownValue].spend += Number(row[12]) || 0;        // M
    
    // Track yesterday's data separately
    if (yesterdayDate && date === yesterdayDate) {
      if (!yesterdayData[breakdownValue]) {
        yesterdayData[breakdownValue] = { spend: 0 };
      }
      yesterdayData[breakdownValue].spend += Number(row[12]) || 0;
    }
  }
  
  // Convert to array and sort by spend
  const breakdownItems = Object.entries(breakdown)
    .map(([item, data]) => ({
      item,
      impressions: data.impressions,
      clicks: data.clicks,
      conversions: data.conversions,
      spend: data.spend,
      cpa: data.conversions > 0 ? data.spend / data.conversions : 0,
      ecpm: data.impressions > 0 ? (data.spend / data.impressions) * 1000 : 0,
      ctr: data.impressions > 0 ? (data.clicks / data.impressions) * 100 : 0,
      vsYesterday: yesterdayData[item] ? 
        ((data.spend - yesterdayData[item].spend * (dates.length)) / (yesterdayData[item].spend * dates.length) * 100) : 0
    }))
    .sort((a, b) => b.spend - a.spend)
    .slice(0, 20);  // Top 20
  
  // Update header
  dashSheet.getRange(breakdownStartRow, 1).setValue(`BREAKDOWN: ${breakdownLabel} for "${selection1}"`);
  
  // Write breakdown data
  if (breakdownItems.length > 0) {
    const rows = breakdownItems.map(item => [
      item.item,
      item.impressions,
      item.clicks,
      item.conversions,
      item.spend,
      item.cpa,
      item.ecpm,
      item.ctr,
      item.vsYesterday
    ]);
    
    dashSheet.getRange(breakdownStartRow + 2, 1, rows.length, 9).setValues(rows);
    
    // Format
    dashSheet.getRange(breakdownStartRow + 2, 2, rows.length, 1).setNumberFormat('#,##0');
    dashSheet.getRange(breakdownStartRow + 2, 3, rows.length, 1).setNumberFormat('#,##0');
    dashSheet.getRange(breakdownStartRow + 2, 4, rows.length, 1).setNumberFormat('#,##0');
    dashSheet.getRange(breakdownStartRow + 2, 5, rows.length, 2).setNumberFormat('$#,##0.00');
    dashSheet.getRange(breakdownStartRow + 2, 7, rows.length, 1).setNumberFormat('$#,##0.000');
    dashSheet.getRange(breakdownStartRow + 2, 8, rows.length, 1).setNumberFormat('0.00"%"');
    dashSheet.getRange(breakdownStartRow + 2, 9, rows.length, 1).setNumberFormat('+0.0"%";-0.0"%";0"%"');
    
    // Conditional formatting for vs Yesterday
    const vsYesterdayRange = dashSheet.getRange(breakdownStartRow + 2, 9, rows.length, 1);
    const positiveRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground('#c8e6c9')
      .setFontColor('#2e7d32')
      .setRanges([vsYesterdayRange])
      .build();
    const negativeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#ffcdd2')
      .setFontColor('#c62828')
      .setRanges([vsYesterdayRange])
      .build();
    
    const rules = dashSheet.getConditionalFormatRules();
    rules.push(positiveRule, negativeRule);
    dashSheet.setConditionalFormatRules(rules);
  } else {
    dashSheet.getRange(breakdownStartRow + 2, 1).setValue('No data found for this selection');
    dashSheet.getRange(breakdownStartRow + 2, 1).setFontStyle('italic').setFontColor('#666666');
  }
  
  Logger.log(`Breakdown table refreshed: ${breakdownItems.length} items`);
}

/**
 * Rebuild dashboard charts (called on selection or metric toggle change)
 */
function cpRebuildDashboardCharts(dashSheet) {
  // Get the data range to determine chart rows
  const lastDataRow = dashSheet.getLastRow();
  // Data starts at row 14, so count rows from there
  let chartRows = 0;
  for (let r = 14; r <= Math.min(lastDataRow, 30); r++) {
    const dateVal = dashSheet.getRange(r, 1).getValue();
    if (dateVal) chartRows++;
  }
  if (chartRows === 0) chartRows = 7;  // Default
  chartRows += 2;  // Add header rows
  
  cpBuildDashboardCharts(dashSheet, chartRows);
  Logger.log('Charts rebuilt');
}

/**
 * Rebuild dashboard charts for V8 layout (called on selection or metric toggle change)
 * V8 data starts at row 25 (headers at row 24)
 */
function cpRebuildDashboardChartsV8(dashSheet) {
  // Get the data range to determine chart rows
  const lastDataRow = dashSheet.getLastRow();
  // V8: Data starts at row 25, so count rows from there
  let chartRows = 0;
  for (let r = 25; r <= Math.min(lastDataRow, 42); r++) {
    const dateVal = dashSheet.getRange(r, 1).getValue();
    if (dateVal) chartRows++;
  }
  if (chartRows === 0) chartRows = 7;  // Default
  chartRows += 1;  // Add header row
  
  cpBuildDashboardChartsV8(dashSheet, chartRows);
  Logger.log('V8 Charts rebuilt');
}

/**
 * Build dashboard charts based on metric toggles
 * Chart 1: Spend, Conversions, CPA (left axis: Spend/CPA, right axis: Conversions)
 * Chart 2: Impressions, Clicks, CTR (left axis: Impr/Clicks, right axis: CTR)
 * 
 * Uses color VARIATIONS per metric within each selection:
 * - Spend/Impr: Base color (solid line)
 * - Conv/Clicks: Lighter shade (dashed line)
 * - CPA/CTR: Darker shade (dotted line)
 * 
 * Only shows lines for ACTIVE selections (hides empty/unselected)
 */
function cpBuildDashboardCharts(dashSheet, chartRows) {
  // Remove existing charts
  const existingCharts = dashSheet.getCharts();
  for (const chart of existingCharts) {
    dashSheet.removeChart(chart);
  }
  
  // Get metric toggle states (row 6, columns B-G)
  // B6=Spend, C6=CPA, D6=Conv, E6=Impr, F6=Clicks, G6=CTR
  const toggles = dashSheet.getRange('B6:G6').getValues()[0];
  const showSpend = toggles[0] === true;
  const showCPA = toggles[1] === true;
  const showConv = toggles[2] === true;
  const showImpr = toggles[3] === true;
  const showClicks = toggles[4] === true;
  const showCTR = toggles[5] === true;
  
  // Get selection names DIRECTLY from the dropdowns (B3:F3)
  const rawSelections = dashSheet.getRange('B3:F3').getDisplayValues()[0];
  
  // Determine which selections are ACTIVE (not empty, not "(Clear)")
  const activeSelections = [];
  for (let i = 0; i < 5; i++) {
    const sel = rawSelections[i];
    if (sel && sel !== '(Clear)' && sel.trim() !== '') {
      activeSelections.push({
        index: i,
        name: sel,
        shortName: sel.length > 12 ? sel.substring(0, 10) + '..' : sel
      });
    }
  }
  
  // If no active selections, don't build charts
  if (activeSelections.length === 0) {
    Logger.log('No active selections - skipping chart build');
    return;
  }
  
  Logger.log(`Active selections: ${activeSelections.map(s => s.name).join(', ')}`);
  
  // Color palette - DISTINCT colors for each selection (easy to differentiate)
  // Each selection gets 3 shades for the 3 metrics
  const colorPalette = {
    0: { base: '#2196F3', light: '#64B5F6', dark: '#1565C0' },  // S1 Blue
    1: { base: '#F44336', light: '#E57373', dark: '#C62828' },  // S2 Red
    2: { base: '#4CAF50', light: '#81C784', dark: '#2E7D32' },  // S3 Green
    3: { base: '#9C27B0', light: '#BA68C8', dark: '#6A1B9A' },  // S4 Purple
    4: { base: '#FF9800', light: '#FFB74D', dark: '#E65100' }   // S5 Orange
  };
  
  // Column mapping (0-indexed from B=2):
  // Spend: B-F (2-6), Conv: G-K (7-11), CPA: L-P (12-16)
  // Impr: Q-U (17-21), Clicks: V-Z (22-26), CTR: AA-AE (27-31)
  const metricColStarts = {
    'Spend': 2, 'Conv': 7, 'CPA': 12,
    'Impr': 17, 'Clicks': 22, 'CTR': 27
  };
  
  // =========================================================================
  // Update trend data headers with selection names (row 12) - ONLY for active
  // =========================================================================
  const allMetrics = ['Spend', 'Conv', 'CPA', 'Impr', 'Clicks', 'CTR'];
  
  for (const metric of allMetrics) {
    const baseCol = metricColStarts[metric];
    for (let s = 0; s < 5; s++) {
      const col = baseCol + s;
      // Find if this selection is active
      const activeSel = activeSelections.find(a => a.index === s);
      if (activeSel) {
        dashSheet.getRange(12, col).setValue(`${activeSel.shortName} ${metric}`);
      } else {
        dashSheet.getRange(12, col).setValue('');  // Clear unused headers
      }
    }
  }
  
  // =========================================================================
  // CHART 1: Spend, Conversions, CPA - ONLY active selections
  // =========================================================================
  if (showSpend || showConv || showCPA) {
    const chart1Builder = dashSheet.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(dashSheet.getRange(12, 1, chartRows, 1))  // Date column
      .setNumHeaders(1)
      .setPosition(1, 9, 0, 0)
      .setOption('title', 'Spend, Conversions & CPA')
      .setOption('width', 700)
      .setOption('height', 400)
      .setOption('legend', { position: 'right', maxLines: 15 })
      .setOption('hAxis', { title: 'Date', slantedText: true, slantedTextAngle: 45 })
      .setOption('vAxes', {
        0: { title: 'Spend / CPA ($)', format: '$#,##0', titleTextStyle: { color: '#333' } },
        1: { title: 'Conversions', format: '#,##0', titleTextStyle: { color: '#ea4335' } }
      })
      .setOption('useFirstColumnAsDomain', true);
    
    let seriesConfig = {};
    let seriesIndex = 0;
    
    // Add Spend series - THICK solid line, CIRCLE markers
    if (showSpend) {
      for (const sel of activeSelections) {
        const col = metricColStarts['Spend'] + sel.index;
        chart1Builder.addRange(dashSheet.getRange(12, col, chartRows, 1));
        seriesConfig[seriesIndex++] = { 
          color: colorPalette[sel.index].base, 
          lineWidth: 4, 
          pointSize: 8,
          pointShape: 'circle',
          type: 'line', 
          targetAxisIndex: 0 
        };
      }
    }
    
    // Add CPA series - MEDIUM line, SQUARE markers (darker shade)
    if (showCPA) {
      for (const sel of activeSelections) {
        const col = metricColStarts['CPA'] + sel.index;
        chart1Builder.addRange(dashSheet.getRange(12, col, chartRows, 1));
        seriesConfig[seriesIndex++] = { 
          color: colorPalette[sel.index].dark, 
          lineWidth: 2, 
          pointSize: 6,
          pointShape: 'square',
          type: 'line', 
          targetAxisIndex: 0
        };
      }
    }
    
    // Add Conversions series - THIN line, TRIANGLE markers (lighter shade)
    if (showConv) {
      for (const sel of activeSelections) {
        const col = metricColStarts['Conv'] + sel.index;
        chart1Builder.addRange(dashSheet.getRange(12, col, chartRows, 1));
        seriesConfig[seriesIndex++] = { 
          color: colorPalette[sel.index].light, 
          lineWidth: 1, 
          pointSize: 5,
          pointShape: 'triangle',
          type: 'line', 
          targetAxisIndex: 1
        };
      }
    }
    
    if (seriesIndex > 0) {
      chart1Builder.setOption('series', seriesConfig);
      dashSheet.insertChart(chart1Builder.build());
    }
  }
  
  // =========================================================================
  // CHART 2: Impressions, Clicks, CTR - ONLY active selections
  // =========================================================================
  if (showImpr || showClicks || showCTR) {
    const chart2Builder = dashSheet.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(dashSheet.getRange(12, 1, chartRows, 1))  // Date column
      .setNumHeaders(1)
      .setPosition(23, 9, 0, 0)
      .setOption('title', 'Impressions, Clicks & CTR')
      .setOption('width', 700)
      .setOption('height', 400)
      .setOption('legend', { position: 'right', maxLines: 15 })
      .setOption('hAxis', { title: 'Date', slantedText: true, slantedTextAngle: 45 })
      .setOption('vAxes', {
        0: { title: 'Impressions / Clicks', format: '#,##0', titleTextStyle: { color: '#333' } },
        1: { title: 'CTR (%)', format: '0.00"%"', titleTextStyle: { color: '#ea4335' } }
      })
      .setOption('useFirstColumnAsDomain', true);
    
    let seriesConfig2 = {};
    let seriesIndex2 = 0;
    
    // Add Impressions series - THIN line, TRIANGLE markers (lighter shade)
    if (showImpr) {
      for (const sel of activeSelections) {
        const col = metricColStarts['Impr'] + sel.index;
        chart2Builder.addRange(dashSheet.getRange(12, col, chartRows, 1));
        seriesConfig2[seriesIndex2++] = { 
          color: colorPalette[sel.index].light, 
          lineWidth: 1, 
          pointSize: 5,
          pointShape: 'triangle',
          type: 'line', 
          targetAxisIndex: 0
        };
      }
    }
    
    // Add Clicks series - MEDIUM line, SQUARE markers (darker shade)
    if (showClicks) {
      for (const sel of activeSelections) {
        const col = metricColStarts['Clicks'] + sel.index;
        chart2Builder.addRange(dashSheet.getRange(12, col, chartRows, 1));
        seriesConfig2[seriesIndex2++] = { 
          color: colorPalette[sel.index].dark, 
          lineWidth: 2, 
          pointSize: 6,
          pointShape: 'square',
          type: 'line', 
          targetAxisIndex: 0
        };
      }
    }
    
    // Add CTR series - THICK solid line, CIRCLE markers
    if (showCTR) {
      for (const sel of activeSelections) {
        const col = metricColStarts['CTR'] + sel.index;
        chart2Builder.addRange(dashSheet.getRange(12, col, chartRows, 1));
        seriesConfig2[seriesIndex2++] = { 
          color: colorPalette[sel.index].base, 
          lineWidth: 4, 
          pointSize: 8,
          pointShape: 'circle',
          type: 'line', 
          targetAxisIndex: 1
        };
      }
    }
    
    if (seriesIndex2 > 0) {
      chart2Builder.setOption('series', seriesConfig2);
      dashSheet.insertChart(chart2Builder.build());
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
// SPOT-LEVEL DAILY STATS FUNCTIONS (V7)
// ============================================================================

/**
 * Pull spot-level (bid-level) daily stats for the last 7 days
 * This provides granular data at the Spot/Bid level for multi-level Dashboard analysis
 * Writes to "Daily Stats (Spot)" sheet
 */
function cpPullDailyStatsSpotLevel() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Get campaign IDs from Legend sheet
    let campaignIds = [];
    const legendSheet = ss.getSheetByName('Legend');
    
    if (legendSheet) {
      const lastRow = legendSheet.getLastRow();
      if (lastRow >= 2) {
        const legendData = legendSheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (const row of legendData) {
          const id = String(row[0]).trim();
          if (id) campaignIds.push(id);
        }
      }
    }
    
    if (campaignIds.length === 0) {
      ui.alert('Error', 'No campaign IDs found in Legend sheet (Column A).', ui.ButtonSet.OK);
      return;
    }
    
    // Build lookup from Control Panel for Strategy/Sub-Strategy
    const cpLookup = cpBuildControlPanelLookup(ss);
    
    Logger.log(`Pulling SPOT-LEVEL daily stats for ${campaignIds.length} campaigns...`);
    
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
    
    // Fetch bid-level stats for each day
    const allSpotStats = [];
    const dailyCounts = {};  // Track entries per day for debugging
    
    for (const dateInfo of dates) {
      Logger.log(`Fetching spot-level stats for ${dateInfo.displayDate} (API date: ${dateInfo.apiDate})...`);
      
      const dayBidData = cpFetchBidDataWithDetails(campaignIds, dateInfo.apiDate);
      const bidCount = Object.keys(dayBidData).length;
      let dayEntries = 0;
      let dayWithData = 0;
      
      // Process each bid
      for (const bidId in dayBidData) {
        const bid = dayBidData[bidId];
        const campaignId = String(bid.campaign_id || '');
        const spotName = bid.spot_name || '';
        
        // Lookup strategy/sub-strategy/campaignName from Control Panel
        // Note: API doesn't always return campaign_name, so use lookup as primary source
        const cpInfo = cpLookup[campaignId] || {};
        
        // Extract countries from geos
        const countries = cpExtractCountriesFromGeos(bid.geos || {});
        
        // Calculate CPA
        const cpa = bid.conversions > 0 ? bid.cost / bid.conversions : 0;
        
        // Use lookup first, then API, then fallback to "Campaign {ID}"
        let campaignNameValue = cpInfo.campaignName || bid.campaign_name || '';
        if (!campaignNameValue && campaignId) {
          campaignNameValue = `Campaign ${campaignId}`;
        }
        
        allSpotStats.push({
          date: dateInfo.displayDate,
          tier1Strategy: cpInfo.strategy || '',
          subStrategy: cpInfo.subStrategy || '',
          campaignId: campaignId,
          campaignName: campaignNameValue,
          spotId: bid.spot_id || '',
          spotName: spotName,
          country: countries,
          bidId: bidId,
          impressions: bid.impressions || 0,
          clicks: bid.clicks || 0,
          conversions: bid.conversions || 0,
          spend: bid.cost || 0,
          cpa: cpa,
          ecpm: bid.ecpm || 0,
          ctr: bid.ctr || 0
        });
        
        dayEntries++;
        if (bid.impressions > 0) dayWithData++;
      }
      
      dailyCounts[dateInfo.displayDate] = { total: dayEntries, withData: dayWithData };
      Logger.log(`  ${dateInfo.displayDate}: ${dayEntries} bids fetched, ${dayWithData} with impressions`);
      
      Utilities.sleep(500);  // Pause between days
    }
    
    // Log summary per day
    Logger.log('=== DAILY SUMMARY ===');
    for (const date in dailyCounts) {
      Logger.log(`  ${date}: ${dailyCounts[date].total} entries, ${dailyCounts[date].withData} with data`);
    }
    
    // Write to sheet
    cpWriteDailyStatsSpotLevel(allSpotStats);
    
    // Count stats
    const dataRows = allSpotStats.filter(s => s.impressions > 0).length;
    const uniqueSpots = [...new Set(allSpotStats.map(s => s.spotName))].length;
    const uniqueDates = [...new Set(allSpotStats.map(s => s.date))];
    
    // Build per-day summary
    let dailySummary = '';
    for (const date in dailyCounts) {
      dailySummary += `${date}: ${dailyCounts[date].withData}/${dailyCounts[date].total}\n`;
    }
    
    ui.alert('Success', 
      `Pulled SPOT-LEVEL daily stats for ${campaignIds.length} campaign(s).\n\n` +
      `Total entries: ${allSpotStats.length}\n` +
      `Entries with data: ${dataRows}\n` +
      `Unique spots: ${uniqueSpots}\n` +
      `Unique dates in data: ${uniqueDates.length}\n` +
      `Date range: ${dates[0].displayDate} to ${dates[dates.length - 1].displayDate}\n\n` +
      `Per-day breakdown (with data / total):\n${dailySummary}`,
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    ui.alert('Error', `Failed to pull spot-level stats: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Build lookup table from Control Panel sheet
 * Returns { campaignId: { strategy, subStrategy, campaignName } }
 */
function cpBuildControlPanelLookup(ss) {
  const lookup = {};
  const cpSheet = ss.getSheetByName(CP_SHEET_NAME);
  
  if (!cpSheet) return lookup;
  
  const lastRow = cpSheet.getLastRow();
  if (lastRow < 3) return lookup;
  
  // Control Panel columns: A=Strategy, B=Sub-Strategy, C=Campaign Name, D=Campaign ID
  const data = cpSheet.getRange(3, 1, lastRow - 2, 4).getValues();
  
  for (const row of data) {
    const campaignId = String(row[3] || '').trim();
    if (campaignId && !lookup[campaignId]) {
      lookup[campaignId] = {
        strategy: String(row[0] || '').trim(),
        subStrategy: String(row[1] || '').trim(),
        campaignName: String(row[2] || '').trim()
      };
    }
  }
  
  Logger.log(`Built Control Panel lookup with ${Object.keys(lookup).length} campaigns`);
  return lookup;
}

/**
 * Extract country codes from bid geos object
 * Returns comma-separated string of country codes
 */
function cpExtractCountriesFromGeos(geos) {
  const countries = new Set();
  
  if (typeof geos === 'object' && geos !== null) {
    for (const geoId in geos) {
      const geo = geos[geoId];
      if (geo && geo.countryCode) {
        countries.add(geo.countryCode);
      } else if (geo && geo.country_code) {
        countries.add(geo.country_code);
      }
    }
  }
  
  return Array.from(countries).sort().join(', ');
}

/**
 * Fetch bid-level data with full details for a single date
 * Returns object keyed by bid_id with stats AND bid details
 */
function cpFetchBidDataWithDetails(campaignIds, dateStr) {
  Logger.log(`Fetching bid details for ${dateStr}...`);
  
  const allBids = {};
  const BATCH_SIZE = 5;
  const MAX_RETRIES = 3;
  
  for (let batchStart = 0; batchStart < campaignIds.length; batchStart += BATCH_SIZE) {
    const batchIds = campaignIds.slice(batchStart, batchStart + BATCH_SIZE);
    
    let pendingIds = [...batchIds];
    let retryCount = 0;
    
    while (pendingIds.length > 0 && retryCount < MAX_RETRIES) {
      if (retryCount > 0) {
        const backoffMs = Math.pow(2, retryCount) * 2000;
        Logger.log(`  Retry ${retryCount}/${MAX_RETRIES} after ${backoffMs}ms...`);
        Utilities.sleep(backoffMs);
      }
      
      const requests = pendingIds.map(id => ({
        url: `${CP_API_BASE_URL}/bids/${id}.json?api_key=${CP_API_KEY}&startDate=${dateStr}&endDate=${dateStr}`,
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
                  
                  allBids[bidId] = {
                    campaign_id: campaignId,
                    campaign_name: bid.campaign_name || '',
                    spot_id: bid.spot_id || '',
                    spot_name: bid.spot_name || '',
                    geos: bid.geos || {},
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
        Logger.log(`Batch error: ${e}`);
        break;
      }
      
      retryCount++;
    }
    
    // Delay between batches
    if (batchStart + BATCH_SIZE < campaignIds.length) {
      Utilities.sleep(2000);
    }
  }
  
  Logger.log(`Got data for ${Object.keys(allBids).length} bids on ${dateStr}`);
  return allBids;
}

/**
 * Write spot-level daily stats to sheet
 */
function cpWriteDailyStatsSpotLevel(spotStats) {
  const sheet = cpGetOrCreateSheet(CP_DAILY_SPOT_SHEET_NAME);
  
  // Clear existing data
  sheet.clear();
  
  // Define headers
  const headers = [
    'Date',           // A
    'Tier 1 Strategy',// B
    'Sub-Strategy',   // C
    'Campaign ID',    // D
    'Campaign Name',  // E
    'Spot ID',        // F
    'Spot Name',      // G
    'Country',        // H
    'Bid ID',         // I
    'Impressions',    // J
    'Clicks',         // K
    'Conversions',    // L
    'Spend',          // M
    'CPA',            // N
    'eCPM',           // O
    'CTR'             // P
  ];
  
  // Write headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white');
  
  if (spotStats.length === 0) {
    Logger.log('No spot-level stats to write');
    return;
  }
  
  // Convert to rows
  const dataRows = spotStats.map(s => [
    s.date,
    s.tier1Strategy,
    s.subStrategy,
    s.campaignId,
    s.campaignName,
    s.spotId,
    s.spotName,
    s.country,
    s.bidId,
    s.impressions,
    s.clicks,
    s.conversions,
    s.spend,
    s.cpa,
    s.ecpm,
    s.ctr
  ]);
  
  // Write data
  sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  
  // Format columns
  const numRows = dataRows.length;
  
  // Date (A)
  sheet.getRange(2, 1, numRows, 1).setNumberFormat('yyyy-mm-dd');
  
  // IDs as text (D, F, I)
  sheet.getRange(2, 4, numRows, 1).setNumberFormat('@');
  sheet.getRange(2, 6, numRows, 1).setNumberFormat('@');
  sheet.getRange(2, 9, numRows, 1).setNumberFormat('@');
  
  // Impressions, Clicks, Conversions (J-L)
  sheet.getRange(2, 10, numRows, 3).setNumberFormat('#,##0');
  
  // Spend, CPA (M-N)
  sheet.getRange(2, 13, numRows, 2).setNumberFormat('$#,##0.00');
  
  // eCPM (O)
  sheet.getRange(2, 15, numRows, 1).setNumberFormat('$#,##0.000');
  
  // CTR (P)
  sheet.getRange(2, 16, numRows, 1).setNumberFormat('0.00"%"');
  
  // Alternating row colors
  for (let i = 2; i <= Math.min(numRows + 1, 500); i++) {  // Limit for performance
    if (i % 2 === 0) {
      sheet.getRange(i, 1, 1, headers.length).setBackground('#e8f5e9');
    }
  }
  
  // Add filter
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  sheet.getRange(1, 1, numRows + 1, headers.length).createFilter();
  
  // Freeze header
  sheet.setFrozenRows(1);
  
  // Auto-resize and adjust key columns
  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(5, 250);  // Campaign Name
  sheet.setColumnWidth(7, 200);  // Spot Name
  
  Logger.log(`Wrote ${dataRows.length} rows to ${CP_DAILY_SPOT_SHEET_NAME}`);
}

// ============================================================================
// DASHBOARD FUNCTIONS (V7) - Multi-Level Dashboard
// ============================================================================

/**
 * Build the V7 Multi-Level Dashboard
 * Supports viewing data at Strategy, Sub-Strategy, Spot, Campaign, and Country levels
 * Allows comparing up to 3 items with overlaid charts
 */
function cpBuildDashboardV7() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if spot-level daily stats exist
  const spotSheet = ss.getSheetByName(CP_DAILY_SPOT_SHEET_NAME);
  if (!spotSheet || spotSheet.getLastRow() < 2) {
    ui.alert('Error', 
      'Please run "Pull Daily Stats (Spot Level)" first to get data.\n\n' +
      'This pulls granular spot-level data required for the V7 Dashboard.',
      ui.ButtonSet.OK);
    return;
  }
  
  try {
    Logger.log('Building V7 Dashboard...');
    
    // Get unique values for dropdowns
    const uniqueValues = cpGetUniqueValuesForDashboard(spotSheet);
    
    Logger.log(`Found: ${uniqueValues.strategies.length} strategies, ` +
               `${uniqueValues.subStrategies.length} sub-strategies, ` +
               `${uniqueValues.spots.length} spots, ` +
               `${uniqueValues.campaigns.length} campaigns, ` +
               `${uniqueValues.countries.length} countries`);
    
    // Get or create V7 dashboard sheet (separate from V8)
    let dashSheet = ss.getSheetByName(CP_DASHBOARD_V7_SHEET_NAME);
    if (dashSheet) {
      // Remove existing charts
      const charts = dashSheet.getCharts();
      for (const chart of charts) {
        dashSheet.removeChart(chart);
      }
      // Clear content, formatting, AND data validations
      dashSheet.clear();
      dashSheet.getRange(1, 1, dashSheet.getMaxRows(), dashSheet.getMaxColumns()).clearDataValidations();
      Logger.log('Cleared existing Dashboard-V7 sheet');
    } else {
      dashSheet = ss.insertSheet(CP_DASHBOARD_V7_SHEET_NAME);
      Logger.log('Created Dashboard-V7 sheet');
    }
    
    // =========================================================================
    // ROW 1: Title
    // =========================================================================
    dashSheet.getRange('A1').setValue('Campaign Performance Dashboard V7');
    dashSheet.getRange('A1').setFontSize(18).setFontWeight('bold');
    
    // =========================================================================
    // ROW 2: View Level Selector
    // =========================================================================
    dashSheet.getRange('A2').setValue('View Level:');
    dashSheet.getRange('A2').setFontWeight('bold');
    
    const viewLevels = ['All', 'Tier 1 Strategy', 'Sub-Strategy', 'Spot Name', 'Campaign', 'Country'];
    const viewLevelRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(viewLevels, true)
      .setAllowInvalid(false)
      .build();
    dashSheet.getRange('B2').setDataValidation(viewLevelRule);
    dashSheet.getRange('B2').setValue('All');
    dashSheet.getRange('B2').setBackground('#e3f2fd').setFontWeight('bold');
    
    // =========================================================================
    // ROW 3: Selection Dropdowns (up to 5 for comparison)
    // =========================================================================
    dashSheet.getRange('A3').setValue('Compare:');
    dashSheet.getRange('A3').setFontWeight('bold');
    
    // Create dropdowns for selections (initially with all options combined)
    // These will be dynamically updated based on View Level
    // Add "(Clear)" option at the start to allow clearing selection
    const allOptions = ['(Clear)', ...cpBuildSelectionOptions(uniqueValues, 'All')];
    const selectionRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(allOptions.length > 1 ? allOptions : ['(Clear)', '(Select View Level first)'], true)
      .setAllowInvalid(true)
      .build();
    
    // 5 Selection colors (matching chart colors - more distinct)
    const selColors = ['#BBDEFB', '#FFCDD2', '#C8E6C9', '#E1BEE7', '#FFE0B2'];  // Blue, Red, Green, Purple, Orange
    const selLabels = ['Blue', 'Red', 'Green', 'Purple', 'Orange'];
    
    // Selection 1-5 (B3-F3)
    for (let i = 0; i < 5; i++) {
      const col = 2 + i;  // B=2, C=3, D=4, E=5, F=6
      const cell = dashSheet.getRange(3, col);
      cell.setDataValidation(selectionRule);
      cell.setValue(i === 0 ? (allOptions[1] || '') : '');
      cell.setBackground(selColors[i]);
      cell.setNote(`Selection ${i + 1} (${selLabels[i]} in chart)${i > 0 ? ' - Optional' : ''}`);
    }
    
    // =========================================================================
    // ROW 4: Date Range Info (Editable - user can type dates manually)
    // =========================================================================
    dashSheet.getRange('A4').setValue('Start Date:');
    dashSheet.getRange('A4').setFontWeight('bold');
    
    // Get actual date range from data
    const spotLastRow = spotSheet.getLastRow();
    const allDates = spotSheet.getRange(2, 1, spotLastRow - 1, 1).getValues().flat().filter(d => d);
    const minDate = allDates.length > 0 ? new Date(Math.min(...allDates.map(d => new Date(d)))) : new Date();
    const maxDate = allDates.length > 0 ? new Date(Math.max(...allDates.map(d => new Date(d)))) : new Date();
    
    // Start date (B4) - editable
    dashSheet.getRange('B4').setValue(minDate);
    dashSheet.getRange('B4').setNumberFormat('yyyy-mm-dd');
    dashSheet.getRange('B4').setBackground('#fff9c4');
    dashSheet.getRange('B4').setNote('Edit to change start date filter');
    
    // End date label and value
    dashSheet.getRange('C4').setValue('End Date:');
    dashSheet.getRange('C4').setFontWeight('bold');
    dashSheet.getRange('D4').setValue(maxDate);
    dashSheet.getRange('D4').setNumberFormat('yyyy-mm-dd');
    dashSheet.getRange('D4').setBackground('#fff9c4');
    dashSheet.getRange('D4').setNote('Edit to change end date filter');
    
    // =========================================================================
    // HIDDEN HELPER CELLS (Column Z) - Store extracted values for formulas
    // =========================================================================
    dashSheet.getRange('Z1').setValue('View Level');
    dashSheet.getRange('Z2').setFormula('=B2');  // Current view level
    // Selections 1-5
    dashSheet.getRange('Z3').setValue('Sel1');
    dashSheet.getRange('Z4').setFormula('=IF(B3="(Clear)","",B3)');
    dashSheet.getRange('Z5').setValue('Sel2');
    dashSheet.getRange('Z6').setFormula('=IF(C3="(Clear)","",C3)');
    dashSheet.getRange('Z7').setValue('Sel3');
    dashSheet.getRange('Z8').setFormula('=IF(D3="(Clear)","",D3)');
    dashSheet.getRange('Z9').setValue('Sel4');
    dashSheet.getRange('Z10').setFormula('=IF(E3="(Clear)","",E3)');
    dashSheet.getRange('Z11').setValue('Sel5');
    dashSheet.getRange('Z12').setFormula('=IF(F3="(Clear)","",F3)');
    // Date range
    dashSheet.getRange('Z13').setValue('Start Date');
    dashSheet.getRange('Z14').setFormula('=B4');
    dashSheet.getRange('Z15').setValue('End Date');
    dashSheet.getRange('Z16').setFormula('=D4');
    
    // Hide column Z
    dashSheet.hideColumns(26);
    
    // =========================================================================
    // ROW 5-6: Metric Toggle - Labels on row 5, Checkboxes on row 6
    // =========================================================================
    dashSheet.getRange('A5').setValue('Metrics:');
    dashSheet.getRange('A5').setFontWeight('bold');
    dashSheet.getRange('A6').setValue('Show:');
    dashSheet.getRange('A6').setFontWeight('bold');
    
    // Clear any existing data validation on rows 5-6
    dashSheet.getRange('B5:G6').clearDataValidations();
    
    const metricNames = ['Spend', 'CPA', 'Conv', 'Impr', 'Clicks', 'CTR'];
    const checkboxRule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .setAllowInvalid(false)
      .build();
    
    // Row 5: Labels, Row 6: Checkboxes (columns B-G)
    for (let i = 0; i < metricNames.length; i++) {
      const col = 2 + i;  // B=2, C=3, D=4, E=5, F=6, G=7
      
      // Label on row 5
      dashSheet.getRange(5, col).setValue(metricNames[i]);
      dashSheet.getRange(5, col).setFontWeight('bold').setHorizontalAlignment('center');
      
      // Checkbox on row 6
      const checkCell = dashSheet.getRange(6, col);
      checkCell.setValue(true);  // Set value BEFORE validation
      checkCell.setDataValidation(checkboxRule);
      checkCell.setHorizontalAlignment('center');
    }
    
    // Set column widths for metric columns
    for (let col = 2; col <= 7; col++) {
      dashSheet.setColumnWidth(col, 60);
    }
    
    // =========================================================================
    // ROW 7: KPI Summary Cards
    // =========================================================================
    dashSheet.getRange('A7').setValue('SUMMARY (Selection 1)');
    dashSheet.getRange('A7:G7').merge();
    dashSheet.getRange('A7').setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    
    // KPI Headers
    const kpiHeaders = ['Total Spend', 'Conversions', 'CPA', 'eCPM', 'Impressions', 'Clicks', 'CTR'];
    dashSheet.getRange('A8:G8').setValues([kpiHeaders]);
    dashSheet.getRange('A8:G8').setFontWeight('bold').setBackground('#e3f2fd');
    
    // KPI Formulas - Aggregate based on Selection 1 (row 9)
    const spotSheetRef = `'${CP_DAILY_SPOT_SHEET_NAME}'`;
    
    // Helper function to build SUMIFS formula for a metric column
    const buildSumFormula = (dataCol, selCell) => {
      return `=IF(OR(${selCell}=""),0,SUMIFS(${spotSheetRef}!$${dataCol}$2:$${dataCol},` +
        `${spotSheetRef}!$B$2:$B,${selCell})+SUMIFS(${spotSheetRef}!$${dataCol}$2:$${dataCol},` +
        `${spotSheetRef}!$C$2:$C,${selCell})+SUMIFS(${spotSheetRef}!$${dataCol}$2:$${dataCol},` +
        `${spotSheetRef}!$E$2:$E,${selCell})+SUMIFS(${spotSheetRef}!$${dataCol}$2:$${dataCol},` +
        `${spotSheetRef}!$G$2:$G,${selCell})+SUMIFS(${spotSheetRef}!$${dataCol}$2:$${dataCol},` +
        `${spotSheetRef}!$H$2:$H,${selCell}))`;
    };
    
    // KPI Values (row 9) - using Selection 1 ($Z$4)
    dashSheet.getRange('A9').setFormula(buildSumFormula('M', '$Z$4'));  // Spend
    dashSheet.getRange('B9').setFormula(buildSumFormula('L', '$Z$4'));  // Conversions
    dashSheet.getRange('C9').setFormula('=IF(B9>0,A9/B9,0)');  // CPA
    dashSheet.getRange('D9').setFormula('=IF(E9>0,(A9/E9)*1000,0)');  // eCPM
    dashSheet.getRange('E9').setFormula(buildSumFormula('J', '$Z$4'));  // Impressions
    dashSheet.getRange('F9').setFormula(buildSumFormula('K', '$Z$4'));  // Clicks
    dashSheet.getRange('G9').setFormula('=IF(E9>0,F9/E9*100,0)');  // CTR
    
    // Format KPI values
    dashSheet.getRange('A9').setNumberFormat('$#,##0.00');
    dashSheet.getRange('B9').setNumberFormat('#,##0');
    dashSheet.getRange('C9').setNumberFormat('$#,##0.00');
    dashSheet.getRange('D9').setNumberFormat('$#,##0.000');
    dashSheet.getRange('E9').setNumberFormat('#,##0');
    dashSheet.getRange('F9').setNumberFormat('#,##0');
    dashSheet.getRange('G9').setNumberFormat('0.00"%"');
    dashSheet.getRange('A9:G9').setFontSize(14).setFontWeight('bold');
    
    // =========================================================================
    // ROW 11-12: Chart Data Headers (expanded for 5 selections × 6 metrics)
    // =========================================================================
    dashSheet.getRange('A11').setValue('TREND DATA');
    dashSheet.getRange('A11').setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    
    // Build headers for all 5 selections × all metrics
    // Columns: Date, Sel1-5 Spend, Sel1-5 Conv, Sel1-5 CPA, Sel1-5 Impr, Sel1-5 Clicks, Sel1-5 CTR
    const metricsOrder = ['Spend', 'Conv', 'CPA', 'Impr', 'Clicks', 'CTR'];
    const trendHeaders = ['Date'];
    for (const metric of metricsOrder) {
      for (let s = 1; s <= 5; s++) {
        trendHeaders.push(`S${s} ${metric}`);
      }
    }
    // That's 1 + 30 = 31 columns (A through AE)
    dashSheet.getRange(12, 1, 1, trendHeaders.length).setValues([trendHeaders]);
    dashSheet.getRange(12, 1, 1, trendHeaders.length).setFontWeight('bold').setBackground('#e3f2fd').setFontSize(9);
    
    // Get unique dates from spot data
    // Use getDisplayValues() to get strings (getValues() returns Date objects which don't dedupe correctly in Set)
    const lastRow = spotSheet.getLastRow();
    const dateStrings = spotSheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
    const dates = [...new Set(dateStrings)].filter(d => d).sort();
    
    // Helper function to build SUMPRODUCT formula for daily metrics
    const buildDailyFormula = (dataCol, selCell, row) => {
      return `=IF(OR(${selCell}="",A${row}=""),0,SUMPRODUCT((${spotSheetRef}!$A$2:$A=A${row})*` +
        `((${spotSheetRef}!$B$2:$B=${selCell})+(${spotSheetRef}!$C$2:$C=${selCell})+(${spotSheetRef}!$E$2:$E=${selCell})+` +
        `(${spotSheetRef}!$G$2:$G=${selCell})+(${spotSheetRef}!$H$2:$H=${selCell}))*(${spotSheetRef}!$${dataCol}$2:$${dataCol})))`;
    };
    
    // Selection cell references
    const selCells = ['$Z$4', '$Z$6', '$Z$8', '$Z$10', '$Z$12'];
    
    // Row 12 = headers, Row 13 = selection labels, Row 14+ = data
    const dataStartRow = 14;
    const trendRows = Math.min(dates.length, 14);
    
    // Add formulas for each date
    for (let i = 0; i < trendRows; i++) {
      const row = dataStartRow + i;
      const dateVal = dates[i];
      
      // Column A: Date
      dashSheet.getRange(row, 1).setValue(dateVal);
      
      let col = 2;  // Start at column B
      
      // For each metric (Spend, Conv, CPA, Impr, Clicks, CTR)
      // Spend (M) for all 5 selections
      for (let s = 0; s < 5; s++) {
        dashSheet.getRange(row, col++).setFormula(buildDailyFormula('M', selCells[s], row));
      }
      
      // Conversions (L) for all 5 selections
      for (let s = 0; s < 5; s++) {
        dashSheet.getRange(row, col++).setFormula(buildDailyFormula('L', selCells[s], row));
      }
      
      // CPA (Spend/Conv) for all 5 selections - columns B-F have spend, G-K have conv
      // CPA formula: =IF(conv>0, spend/conv, 0)
      for (let s = 0; s < 5; s++) {
        const spendCol = 2 + s;  // B=2 for S1 Spend
        const convCol = 7 + s;   // G=7 for S1 Conv
        const spendLetter = String.fromCharCode(65 + spendCol - 1);  // B, C, D, E, F
        const convLetter = String.fromCharCode(65 + convCol - 1);    // G, H, I, J, K
        dashSheet.getRange(row, col++).setFormula(`=IF(${convLetter}${row}>0,${spendLetter}${row}/${convLetter}${row},0)`);
      }
      
      // Impressions (J) for all 5 selections
      for (let s = 0; s < 5; s++) {
        dashSheet.getRange(row, col++).setFormula(buildDailyFormula('J', selCells[s], row));
      }
      
      // Clicks (K) for all 5 selections
      for (let s = 0; s < 5; s++) {
        dashSheet.getRange(row, col++).setFormula(buildDailyFormula('K', selCells[s], row));
      }
      
      // CTR (Clicks/Impr*100) for all 5 selections
      for (let s = 0; s < 5; s++) {
        const imprCol = 17 + s;  // Impressions start at column 17 (Q=17)
        const clicksCol = 22 + s; // Clicks start at column 22 (V=22)
        const imprLetter = String.fromCharCode(65 + imprCol - 1);
        const clicksLetter = String.fromCharCode(65 + clicksCol - 1);
        dashSheet.getRange(row, col++).setFormula(`=IF(${imprLetter}${row}>0,${clicksLetter}${row}/${imprLetter}${row}*100,0)`);
      }
    }
    
    // Format trend data columns
    // Spend: columns 2-6 (B-F)
    dashSheet.getRange(dataStartRow, 2, trendRows, 5).setNumberFormat('$#,##0.00');
    // Conv: columns 7-11 (G-K)
    dashSheet.getRange(dataStartRow, 7, trendRows, 5).setNumberFormat('#,##0');
    // CPA: columns 12-16 (L-P)
    dashSheet.getRange(dataStartRow, 12, trendRows, 5).setNumberFormat('$#,##0.00');
    // Impr: columns 17-21 (Q-U)
    dashSheet.getRange(dataStartRow, 17, trendRows, 5).setNumberFormat('#,##0');
    // Clicks: columns 22-26 (V-Z)
    dashSheet.getRange(dataStartRow, 22, trendRows, 5).setNumberFormat('#,##0');
    // CTR: columns 27-31 (AA-AE)
    dashSheet.getRange(dataStartRow, 27, trendRows, 5).setNumberFormat('0.00"%"');
    
    // Date column
    dashSheet.getRange(dataStartRow, 1, trendRows, 1).setNumberFormat('yyyy-mm-dd');
    
    // =========================================================================
    // BUILD CHARTS using helper function (allows rebuild on metric toggle)
    // =========================================================================
    const chartRows = trendRows + 2;  // headers (12) + labels (13) + data rows
    cpBuildDashboardCharts(dashSheet, chartRows);
    
    // =========================================================================
    // Add Selection Name Labels (row 13, below header, to show what S1-S5 represent)
    // =========================================================================
    dashSheet.getRange('A13').setValue('Selection:');
    // Show selection names for spend columns (B-F)
    dashSheet.getRange('B13').setFormula('=IF(B3="(Clear)","",B3)');
    dashSheet.getRange('C13').setFormula('=IF(C3="(Clear)","",C3)');
    dashSheet.getRange('D13').setFormula('=IF(D3="(Clear)","",D3)');
    dashSheet.getRange('E13').setFormula('=IF(E3="(Clear)","",E3)');
    dashSheet.getRange('F13').setFormula('=IF(F3="(Clear)","",F3)');
    // Repeat selection names for other metric groups
    const selFormulas = ['=IF(B3="(Clear)","",B3)', '=IF(C3="(Clear)","",C3)', '=IF(D3="(Clear)","",D3)', 
                         '=IF(E3="(Clear)","",E3)', '=IF(F3="(Clear)","",F3)'];
    for (let col = 7; col <= 27; col += 5) {
      for (let s = 0; s < 5; s++) {
        dashSheet.getRange(13, col + s).setFormula(selFormulas[s]);
      }
    }
    dashSheet.getRange(13, 1, 1, 31).setFontStyle('italic').setFontColor('#666666').setFontSize(9).setBackground('#f5f5f5');
    
    // =========================================================================
    // BREAKDOWN TABLE (Row 36+) - moved down for expanded trend data + selection labels
    // =========================================================================
    const breakdownStartRow = 36;
    dashSheet.getRange(`A${breakdownStartRow}`).setValue('BREAKDOWN (Child Items for Selection 1)');
    dashSheet.getRange(`A${breakdownStartRow}:I${breakdownStartRow}`).merge();
    dashSheet.getRange(`A${breakdownStartRow}`).setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    
    // Breakdown headers
    const breakdownHeaders = ['Item', 'Impressions', 'Clicks', 'Conversions', 'Spend', 'CPA', 'eCPM', 'CTR', 'vs Yesterday'];
    dashSheet.getRange(breakdownStartRow + 1, 1, 1, breakdownHeaders.length).setValues([breakdownHeaders]);
    dashSheet.getRange(breakdownStartRow + 1, 1, 1, breakdownHeaders.length)
      .setFontWeight('bold')
      .setBackground('#e3f2fd');
    
    // Add note about breakdown
    dashSheet.getRange(breakdownStartRow + 2, 1).setValue('(Breakdown data populated based on selection)');
    dashSheet.getRange(breakdownStartRow + 2, 1).setFontStyle('italic').setFontColor('#666666');
    
    // =========================================================================
    // COLUMN WIDTHS
    // =========================================================================
    dashSheet.setColumnWidth(1, 100);  // A - Labels/Date
    // Selections B-F
    for (let i = 2; i <= 6; i++) {
      dashSheet.setColumnWidth(i, 140);
    }
    // Metric columns (narrower)
    for (let i = 7; i <= 31; i++) {
      dashSheet.setColumnWidth(i, 70);
    }
    
    // =========================================================================
    // STORE UNIQUE VALUES FOR DYNAMIC DROPDOWNS (Hidden sheet)
    // =========================================================================
    cpStoreDashboardOptions(ss, uniqueValues);
    
    ui.alert('Dashboard V7 Built!', 
      `Multi-Level Dashboard created with:\n\n` +
      `• ${uniqueValues.strategies.length} Strategies\n` +
      `• ${uniqueValues.subStrategies.length} Sub-Strategies\n` +
      `• ${uniqueValues.spots.length} Spots\n` +
      `• ${uniqueValues.campaigns.length} Campaigns\n` +
      `• ${uniqueValues.countries.length} Countries\n\n` +
      `Features:\n` +
      `• Compare up to 5 items\n` +
      `• 2 Charts: Spend + Conversions\n` +
      `• Metric toggles (row 5)\n` +
      `• Breakdown table with vs Yesterday`,
      ui.ButtonSet.OK);
    
    Logger.log('V7 Dashboard built successfully');
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    ui.alert('Error', `Failed to build V7 Dashboard: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ============================================================================
// V8 DASHBOARD WITH FILTERS AND COMPARISON DATE RANGE
// ============================================================================

/**
 * Build the V8 Dashboard with Advanced Filtering and Period Comparison
 * 
 * Features:
 * - Multi-level filtering with "equals" and "contains" modes
 * - Multiple filters (up to 3) with AND logic
 * - Comparison date range for period-over-period analysis
 * - All V7 features (multi-select, charts, breakdown)
 */
function cpBuildDashboardV8() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if spot-level daily stats exist
  const spotSheet = ss.getSheetByName(CP_DAILY_SPOT_SHEET_NAME);
  if (!spotSheet || spotSheet.getLastRow() < 2) {
    ui.alert('Error', 
      'Please run "Pull Daily Stats (Spot Level)" first to get data.\n\n' +
      'This pulls granular spot-level data required for the Dashboard.',
      ui.ButtonSet.OK);
    return;
  }
  
  try {
    Logger.log('Building V8 Dashboard with Filters...');
    
    // Get unique values for dropdowns
    const uniqueValues = cpGetUniqueValuesForDashboard(spotSheet);
    
    Logger.log(`Found: ${uniqueValues.strategies.length} strategies, ` +
               `${uniqueValues.subStrategies.length} sub-strategies, ` +
               `${uniqueValues.spots.length} spots, ` +
               `${uniqueValues.campaigns.length} campaigns, ` +
               `${uniqueValues.countries.length} countries`);
    
    // Get or create dashboard sheet
    let dashSheet = ss.getSheetByName(CP_DASHBOARD_SHEET_NAME);
    if (dashSheet) {
      const charts = dashSheet.getCharts();
      for (const chart of charts) {
        dashSheet.removeChart(chart);
      }
      dashSheet.clear();
      dashSheet.getRange(1, 1, dashSheet.getMaxRows(), dashSheet.getMaxColumns()).clearDataValidations();
      Logger.log('Cleared existing Dashboard sheet');
    } else {
      dashSheet = ss.insertSheet(CP_DASHBOARD_SHEET_NAME);
      Logger.log('Created Dashboard sheet');
    }
    
    // =========================================================================
    // ROW 1: Title
    // =========================================================================
    dashSheet.getRange('A1').setValue('Campaign Performance Dashboard V8');
    dashSheet.getRange('A1').setFontSize(18).setFontWeight('bold');
    dashSheet.getRange('H1').setValue('With Filters & Period Comparison');
    dashSheet.getRange('H1').setFontStyle('italic').setFontColor('#666666');
    
    // =========================================================================
    // ROW 2: View Level Selector
    // =========================================================================
    dashSheet.getRange('A2').setValue('View Level:');
    dashSheet.getRange('A2').setFontWeight('bold');
    
    const viewLevels = ['All', 'Tier 1 Strategy', 'Sub-Strategy', 'Spot Name', 'Campaign', 'Country'];
    const viewLevelRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(viewLevels, true)
      .setAllowInvalid(false)
      .build();
    dashSheet.getRange('B2').setDataValidation(viewLevelRule);
    dashSheet.getRange('B2').setValue('All');
    dashSheet.getRange('B2').setBackground('#e3f2fd').setFontWeight('bold');
    
    // =========================================================================
    // ROW 3: Compare Selections (up to 5)
    // =========================================================================
    dashSheet.getRange('A3').setValue('Compare:');
    dashSheet.getRange('A3').setFontWeight('bold');
    
    // Include "(All)" option to see breakdown of all items matching filters
    const allOptions = ['(Clear)', '(All)', ...cpBuildSelectionOptions(uniqueValues, 'All')];
    const selectionRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(allOptions.length > 2 ? allOptions : ['(Clear)', '(All)', '(Select View Level first)'], true)
      .setAllowInvalid(true)
      .build();
    
    const selColors = ['#BBDEFB', '#FFCDD2', '#C8E6C9', '#E1BEE7', '#FFE0B2'];
    const selLabels = ['Blue', 'Red', 'Green', 'Purple', 'Orange'];
    
    for (let i = 0; i < 5; i++) {
      const col = 2 + i;
      const cell = dashSheet.getRange(3, col);
      cell.setDataValidation(selectionRule);
      cell.setValue(i === 0 ? (allOptions[1] || '') : '');
      cell.setBackground(selColors[i]);
      cell.setNote(`Selection ${i + 1} (${selLabels[i]} in chart)${i > 0 ? ' - Optional' : ''}`);
    }
    
    // =========================================================================
    // ROW 4: Breakdown By (own row below Compare)
    // =========================================================================
    dashSheet.getRange('A4').setValue('Breakdown:');
    dashSheet.getRange('A4').setFontWeight('bold');
    
    const breakdownByOptions = ['Tier 1 Strategy', 'Sub-Strategy', 'Campaign', 'Spot Name', 'Country'];
    const breakdownByRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(breakdownByOptions, true)
      .setAllowInvalid(false)
      .build();
    dashSheet.getRange('B4').setDataValidation(breakdownByRule);
    dashSheet.getRange('B4').setValue('Spot Name');  // Default to Spot Name
    dashSheet.getRange('B4').setBackground('#e1bee7');  // Purple to stand out
    dashSheet.getRange('B4').setNote('Choose how to group the breakdown table');
    
    // =========================================================================
    // ROW 5: Primary Date Range
    // =========================================================================
    dashSheet.getRange('A5').setValue('Date Range:');
    dashSheet.getRange('A5').setFontWeight('bold');
    
    const spotLastRow = spotSheet.getLastRow();
    const allDates = spotSheet.getRange(2, 1, spotLastRow - 1, 1).getValues().flat().filter(d => d);
    const minDate = allDates.length > 0 ? new Date(Math.min(...allDates.map(d => new Date(d)))) : new Date();
    const maxDate = allDates.length > 0 ? new Date(Math.max(...allDates.map(d => new Date(d)))) : new Date();
    
    dashSheet.getRange('B5').setValue(minDate);
    dashSheet.getRange('B5').setNumberFormat('yyyy-mm-dd');
    dashSheet.getRange('B5').setBackground('#e8f5e9');
    dashSheet.getRange('B5').setNote('Primary period start date');
    
    dashSheet.getRange('C5').setValue('to');
    dashSheet.getRange('C5').setHorizontalAlignment('center');
    
    dashSheet.getRange('D5').setValue(maxDate);
    dashSheet.getRange('D5').setNumberFormat('yyyy-mm-dd');
    dashSheet.getRange('D5').setBackground('#e8f5e9');
    dashSheet.getRange('D5').setNote('Primary period end date');
    
    // =========================================================================
    // ROW 6: Comparison Date Range (for period-over-period)
    // =========================================================================
    dashSheet.getRange('A6').setValue('Compare To:');
    dashSheet.getRange('A6').setFontWeight('bold');
    
    // Default comparison: previous period of same length
    const periodLength = Math.round((maxDate - minDate) / (1000 * 60 * 60 * 24)) + 1;
    const compEnd = new Date(minDate);
    compEnd.setDate(compEnd.getDate() - 1);
    const compStart = new Date(compEnd);
    compStart.setDate(compStart.getDate() - periodLength + 1);
    
    dashSheet.getRange('B6').setValue(compStart);
    dashSheet.getRange('B6').setNumberFormat('yyyy-mm-dd');
    dashSheet.getRange('B6').setBackground('#fff3e0');
    dashSheet.getRange('B6').setNote('Comparison period start (leave blank to disable)');
    
    dashSheet.getRange('C6').setValue('to');
    dashSheet.getRange('C6').setHorizontalAlignment('center');
    
    dashSheet.getRange('D6').setValue(compEnd);
    dashSheet.getRange('D6').setNumberFormat('yyyy-mm-dd');
    dashSheet.getRange('D6').setBackground('#fff3e0');
    dashSheet.getRange('D6').setNote('Comparison period end');
    
    dashSheet.getRange('E6').setValue('(Leave blank to disable comparison)');
    dashSheet.getRange('E6').setFontStyle('italic').setFontColor('#666666').setFontSize(9);
    
    // =========================================================================
    // ROWS 7-11: Filter Section
    // =========================================================================
    dashSheet.getRange('A7').setValue('FILTERS');
    dashSheet.getRange('A7:G7').merge();
    dashSheet.getRange('A7').setFontWeight('bold').setBackground('#ff9800').setFontColor('white');
    
    // Filter headers
    dashSheet.getRange('A8').setValue('');
    dashSheet.getRange('B8').setValue('Field');
    dashSheet.getRange('C8').setValue('Mode');
    dashSheet.getRange('D8').setValue('Value');
    dashSheet.getRange('B8:D8').setFontWeight('bold').setBackground('#ffe0b2');
    
    // Filter field options
    const filterFields = ['(None)', 'Tier 1 Strategy', 'Sub-Strategy', 'Campaign', 'Spot Name', 'Country'];
    const filterFieldRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(filterFields, true)
      .setAllowInvalid(false)
      .build();
    
    // Filter mode options
    const filterModes = ['equals', 'contains'];
    const filterModeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(filterModes, true)
      .setAllowInvalid(false)
      .build();
    
    // Create 3 filter rows (rows 9, 10, 11)
    for (let i = 0; i < 3; i++) {
      const row = 9 + i;
      dashSheet.getRange(row, 1).setValue(`Filter ${i + 1}:`);
      dashSheet.getRange(row, 1).setFontWeight('bold');
      
      // Field dropdown
      dashSheet.getRange(row, 2).setDataValidation(filterFieldRule);
      dashSheet.getRange(row, 2).setValue('(None)');
      dashSheet.getRange(row, 2).setBackground('#fff9c4');
      
      // Mode dropdown
      dashSheet.getRange(row, 3).setDataValidation(filterModeRule);
      dashSheet.getRange(row, 3).setValue('contains');
      dashSheet.getRange(row, 3).setBackground('#fff9c4');
      
      // Value input (columns D-F merged)
      dashSheet.getRange(row, 4, 1, 3).merge();
      dashSheet.getRange(row, 4).setValue('');
      dashSheet.getRange(row, 4).setBackground('#fff9c4');
      dashSheet.getRange(row, 4).setNote('Enter filter value (text to match)');
    }
    
    // Apply Filters / Refresh note
    dashSheet.getRange('G9').setValue('← Menu: Apply Filters / Refresh');
    dashSheet.getRange('G9').setFontStyle('italic').setFontColor('#666666').setFontSize(9);
    
    // =========================================================================
    // ROW 12-13: Metric Toggle - Labels on row 12, Checkboxes on row 13
    // =========================================================================
    dashSheet.getRange('A12').setValue('Metrics:');
    dashSheet.getRange('A12').setFontWeight('bold');
    dashSheet.getRange('A13').setValue('Show:');
    dashSheet.getRange('A13').setFontWeight('bold');
    
    dashSheet.getRange('B12:H13').clearDataValidations();
    
    const metricNames = ['Spend', 'CPA', 'Conv', 'Impr', 'Clicks', 'CTR', 'eCPM'];
    const checkboxRule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .setAllowInvalid(false)
      .build();
    
    for (let i = 0; i < metricNames.length; i++) {
      const col = 2 + i;
      dashSheet.getRange(12, col).setValue(metricNames[i]);
      dashSheet.getRange(12, col).setFontWeight('bold').setHorizontalAlignment('center');
      const checkCell = dashSheet.getRange(13, col);
      checkCell.setValue(true);
      checkCell.setDataValidation(checkboxRule);
      checkCell.setHorizontalAlignment('center');
    }
    
    for (let col = 2; col <= 8; col++) {
      dashSheet.setColumnWidth(col, 60);
    }
    
    // =========================================================================
    // ROW 15: KPI Summary Cards
    // =========================================================================
    dashSheet.getRange('A15').setValue('SUMMARY (Selection 1 - Primary Period)');
    dashSheet.getRange('A15:H15').merge();
    dashSheet.getRange('A15').setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    
    const kpiHeaders = ['Total Spend', 'Conversions', 'CPA', 'eCPM', 'Impressions', 'Clicks', 'CTR', 'Avg eCPM'];
    dashSheet.getRange('A16:H16').setValues([kpiHeaders]);
    dashSheet.getRange('A16:H16').setFontWeight('bold').setBackground('#e3f2fd');
    
    // KPI Formulas using helper cells
    const spotSheetRef = `'${CP_DAILY_SPOT_SHEET_NAME}'`;
    
    const buildSumFormula = (dataCol, selCell) => {
      return `=IF(OR(${selCell}=""),0,SUMIFS(${spotSheetRef}!$${dataCol}$2:$${dataCol},` +
        `${spotSheetRef}!$B$2:$B,${selCell})+SUMIFS(${spotSheetRef}!$${dataCol}$2:$${dataCol},` +
        `${spotSheetRef}!$C$2:$C,${selCell})+SUMIFS(${spotSheetRef}!$${dataCol}$2:$${dataCol},` +
        `${spotSheetRef}!$E$2:$E,${selCell})+SUMIFS(${spotSheetRef}!$${dataCol}$2:$${dataCol},` +
        `${spotSheetRef}!$G$2:$G,${selCell})+SUMIFS(${spotSheetRef}!$${dataCol}$2:$${dataCol},` +
        `${spotSheetRef}!$H$2:$H,${selCell}))`;
    };
    
    // KPI row 17 (values)
    dashSheet.getRange('A17').setFormula(buildSumFormula('M', '$AA$4'));  // Spend
    dashSheet.getRange('B17').setFormula(buildSumFormula('L', '$AA$4'));  // Conversions
    dashSheet.getRange('C17').setFormula('=IF(B17>0,A17/B17,0)');  // CPA
    dashSheet.getRange('D17').setFormula('=IF(E17>0,(A17/E17)*1000,0)');  // eCPM
    dashSheet.getRange('E17').setFormula(buildSumFormula('J', '$AA$4'));  // Impressions
    dashSheet.getRange('F17').setFormula(buildSumFormula('K', '$AA$4'));  // Clicks
    dashSheet.getRange('G17').setFormula('=IF(E17>0,F17/E17*100,0)');  // CTR
    
    // Avg eCPM - simple average of eCPM values matching the selection (where impressions > 0)
    // Uses SUMPRODUCT to calculate: sum of eCPM / count of matching rows with impressions
    const avgEcpmFormula = `=IF(OR($AA$4=""),0,IFERROR(` +
      `SUMPRODUCT((${spotSheetRef}!$J$2:$J>0)*((${spotSheetRef}!$B$2:$B=$AA$4)+(${spotSheetRef}!$C$2:$C=$AA$4)+` +
      `(${spotSheetRef}!$E$2:$E=$AA$4)+(${spotSheetRef}!$G$2:$G=$AA$4)+(${spotSheetRef}!$H$2:$H=$AA$4))*(${spotSheetRef}!$O$2:$O))/` +
      `SUMPRODUCT((${spotSheetRef}!$J$2:$J>0)*((${spotSheetRef}!$B$2:$B=$AA$4)+(${spotSheetRef}!$C$2:$C=$AA$4)+` +
      `(${spotSheetRef}!$E$2:$E=$AA$4)+(${spotSheetRef}!$G$2:$G=$AA$4)+(${spotSheetRef}!$H$2:$H=$AA$4))*1),0))`;
    dashSheet.getRange('H17').setFormula(avgEcpmFormula);

    // Format KPI values
    dashSheet.getRange('A17').setNumberFormat('$#,##0.00');
    dashSheet.getRange('B17').setNumberFormat('#,##0');
    dashSheet.getRange('C17').setNumberFormat('$#,##0.00');
    dashSheet.getRange('D17').setNumberFormat('$#,##0.000');
    dashSheet.getRange('E17').setNumberFormat('#,##0');
    dashSheet.getRange('F17').setNumberFormat('#,##0');
    dashSheet.getRange('G17').setNumberFormat('0.00"%"');
    dashSheet.getRange('H17').setNumberFormat('$#,##0.000');
    dashSheet.getRange('A17:H17').setFontSize(14).setFontWeight('bold');
    
    // =========================================================================
    // ROW 19: Comparison Period Summary (if enabled)
    // =========================================================================
    dashSheet.getRange('A19').setValue('COMPARISON PERIOD (vs Above)');
    dashSheet.getRange('A19:G19').merge();
    dashSheet.getRange('A19').setFontWeight('bold').setBackground('#ff9800').setFontColor('white');
    
    dashSheet.getRange('A20:G20').setValues([['Spend', 'Conv', 'CPA', 'Δ Spend', 'Δ Conv', 'Δ CPA', 'Status']]);
    dashSheet.getRange('A20:G20').setFontWeight('bold').setBackground('#ffe0b2');
    
    // Row 21: Comparison values (placeholder - populated by Apply Filters)
    dashSheet.getRange('A21').setValue('(Run "Apply Filters" to see comparison)');
    dashSheet.getRange('A21:G21').merge();
    dashSheet.getRange('A21').setFontStyle('italic').setFontColor('#666666');
    
    // =========================================================================
    // HIDDEN HELPER CELLS (Column AA) - Store filter/selection values
    // =========================================================================
    dashSheet.getRange('AA1').setValue('V8 Config');
    dashSheet.getRange('AA2').setValue('View Level');
    dashSheet.getRange('AA3').setFormula('=B2');
    dashSheet.getRange('AA4').setFormula('=IF(B3="(Clear)","",B3)');  // Selection 1
    dashSheet.getRange('AA5').setFormula('=IF(C3="(Clear)","",C3)');  // Selection 2
    dashSheet.getRange('AA6').setFormula('=IF(D3="(Clear)","",D3)');  // Selection 3
    dashSheet.getRange('AA7').setFormula('=IF(E3="(Clear)","",E3)');  // Selection 4
    dashSheet.getRange('AA8').setFormula('=IF(F3="(Clear)","",F3)');  // Selection 5
    // Date ranges (Row 5 = Primary, Row 6 = Comparison)
    dashSheet.getRange('AA10').setValue('Primary Start');
    dashSheet.getRange('AA11').setFormula('=B5');
    dashSheet.getRange('AA12').setValue('Primary End');
    dashSheet.getRange('AA13').setFormula('=D5');
    dashSheet.getRange('AA14').setValue('Comp Start');
    dashSheet.getRange('AA15').setFormula('=B6');
    dashSheet.getRange('AA16').setValue('Comp End');
    dashSheet.getRange('AA17').setFormula('=D6');
    // Filter values (Rows 9, 10, 11)
    dashSheet.getRange('AA20').setValue('Filter1 Field');
    dashSheet.getRange('AA21').setFormula('=B9');
    dashSheet.getRange('AA22').setValue('Filter1 Mode');
    dashSheet.getRange('AA23').setFormula('=C9');
    dashSheet.getRange('AA24').setValue('Filter1 Value');
    dashSheet.getRange('AA25').setFormula('=D9');
    dashSheet.getRange('AA26').setValue('Filter2 Field');
    dashSheet.getRange('AA27').setFormula('=B10');
    dashSheet.getRange('AA28').setValue('Filter2 Mode');
    dashSheet.getRange('AA29').setFormula('=C10');
    dashSheet.getRange('AA30').setValue('Filter2 Value');
    dashSheet.getRange('AA31').setFormula('=D10');
    dashSheet.getRange('AA32').setValue('Filter3 Field');
    dashSheet.getRange('AA33').setFormula('=B11');
    dashSheet.getRange('AA34').setValue('Filter3 Mode');
    dashSheet.getRange('AA35').setFormula('=C11');
    dashSheet.getRange('AA36').setValue('Filter3 Value');
    dashSheet.getRange('AA37').setFormula('=D11');
    // Breakdown By setting (Row 4)
    dashSheet.getRange('AA38').setValue('Breakdown By');
    dashSheet.getRange('AA39').setFormula('=B4');
    
    // Hide column AA
    dashSheet.hideColumns(27);
    
    // =========================================================================
    // ROW 23+: TREND DATA (header row)
    // =========================================================================
    dashSheet.getRange('A23').setValue('TREND DATA');
    dashSheet.getRange('A23').setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    
    // Headers for trend data (V8 adds eCPM)
    const metricsOrder = ['Spend', 'Conv', 'CPA', 'Impr', 'Clicks', 'CTR', 'eCPM'];
    const trendHeaders = ['Date'];
    for (const metric of metricsOrder) {
      for (let s = 1; s <= 5; s++) {
        trendHeaders.push(`S${s} ${metric}`);
      }
    }
    dashSheet.getRange(24, 1, 1, trendHeaders.length).setValues([trendHeaders]);
    dashSheet.getRange(24, 1, 1, trendHeaders.length).setFontWeight('bold').setBackground('#e3f2fd').setFontSize(9);
    
    // Get dates and populate trend data
    const lastRow = spotSheet.getLastRow();
    const dateStrings = spotSheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
    const dates = [...new Set(dateStrings)].filter(d => d).sort();
    
    const buildDailyFormula = (dataCol, selCell, row) => {
      return `=IF(OR(${selCell}="",A${row}=""),0,SUMPRODUCT((${spotSheetRef}!$A$2:$A=A${row})*` +
        `((${spotSheetRef}!$B$2:$B=${selCell})+(${spotSheetRef}!$C$2:$C=${selCell})+(${spotSheetRef}!$E$2:$E=${selCell})+` +
        `(${spotSheetRef}!$G$2:$G=${selCell})+(${spotSheetRef}!$H$2:$H=${selCell}))*(${spotSheetRef}!$${dataCol}$2:$${dataCol})))`;
    };
    
    const selCells = ['$AA$4', '$AA$5', '$AA$6', '$AA$7', '$AA$8'];
    const dataStartRow = 25;  // Row after headers (row 24 is headers)
    const trendRows = Math.min(dates.length, 14);
    
    for (let i = 0; i < trendRows; i++) {
      const row = dataStartRow + i;
      const dateVal = dates[i];
      
      dashSheet.getRange(row, 1).setValue(dateVal);
      
      let col = 2;
      // Spend
      for (let s = 0; s < 5; s++) {
        dashSheet.getRange(row, col++).setFormula(buildDailyFormula('M', selCells[s], row));
      }
      // Conv
      for (let s = 0; s < 5; s++) {
        dashSheet.getRange(row, col++).setFormula(buildDailyFormula('L', selCells[s], row));
      }
      // CPA
      for (let s = 0; s < 5; s++) {
        const spendCol = 2 + s;
        const convCol = 7 + s;
        const spendLetter = String.fromCharCode(65 + spendCol - 1);
        const convLetter = String.fromCharCode(65 + convCol - 1);
        dashSheet.getRange(row, col++).setFormula(`=IF(${convLetter}${row}>0,${spendLetter}${row}/${convLetter}${row},0)`);
      }
      // Impr
      for (let s = 0; s < 5; s++) {
        dashSheet.getRange(row, col++).setFormula(buildDailyFormula('J', selCells[s], row));
      }
      // Clicks
      for (let s = 0; s < 5; s++) {
        dashSheet.getRange(row, col++).setFormula(buildDailyFormula('K', selCells[s], row));
      }
      // CTR
      for (let s = 0; s < 5; s++) {
        const imprCol = 17 + s;
        const clicksCol = 22 + s;
        const imprLetter = String.fromCharCode(65 + imprCol - 1);
        const clicksLetter = String.fromCharCode(65 + clicksCol - 1);
        dashSheet.getRange(row, col++).setFormula(`=IF(${imprLetter}${row}>0,${clicksLetter}${row}/${imprLetter}${row}*100,0)`);
      }
      
      // eCPM = (Spend / Impressions) * 1000
      for (let s = 0; s < 5; s++) {
        const spendCol = 2 + s;
        const imprCol = 17 + s;
        const spendLetter = String.fromCharCode(65 + spendCol - 1);
        const imprLetter = String.fromCharCode(65 + imprCol - 1);
        dashSheet.getRange(row, col++).setFormula(`=IF(${imprLetter}${row}>0,(${spendLetter}${row}/${imprLetter}${row})*1000,0)`);
      }
    }
    
    // Format trend data
    dashSheet.getRange(dataStartRow, 1, trendRows, 1).setNumberFormat('yyyy-mm-dd');
    dashSheet.getRange(dataStartRow, 2, trendRows, 5).setNumberFormat('$#,##0.00');  // Spend
    dashSheet.getRange(dataStartRow, 7, trendRows, 5).setNumberFormat('#,##0');      // Conv
    dashSheet.getRange(dataStartRow, 12, trendRows, 5).setNumberFormat('$#,##0.00'); // CPA
    dashSheet.getRange(dataStartRow, 17, trendRows, 5).setNumberFormat('#,##0');     // Impr
    dashSheet.getRange(dataStartRow, 22, trendRows, 5).setNumberFormat('#,##0');     // Clicks
    dashSheet.getRange(dataStartRow, 27, trendRows, 5).setNumberFormat('0.00"%"');   // CTR
    dashSheet.getRange(dataStartRow, 32, trendRows, 5).setNumberFormat('$#,##0.000'); // eCPM
    
    // =========================================================================
    // BUILD CHARTS
    // =========================================================================
    const chartRows = trendRows + 1;
    cpBuildDashboardChartsV8(dashSheet, chartRows);
    
    // =========================================================================
    // BREAKDOWN TABLE (Row 65+) - moved down to accommodate 3 charts
    // =========================================================================
    const breakdownStartRow = 66;
    dashSheet.getRange(`A${breakdownStartRow}`).setValue('BREAKDOWN (Child Items for Selection 1)');
    dashSheet.getRange(`A${breakdownStartRow}:I${breakdownStartRow}`).merge();
    dashSheet.getRange(`A${breakdownStartRow}`).setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    
    const breakdownHeaders = ['Item', 'Impressions', 'Clicks', 'Conversions', 'Spend', 'CPA', 'eCPM', 'CTR', 'vs Comparison'];
    dashSheet.getRange(breakdownStartRow + 1, 1, 1, breakdownHeaders.length).setValues([breakdownHeaders]);
    dashSheet.getRange(breakdownStartRow + 1, 1, 1, breakdownHeaders.length)
      .setFontWeight('bold')
      .setBackground('#e3f2fd');
    
    dashSheet.getRange(breakdownStartRow + 2, 1).setValue('(Run "Apply Filters" to populate breakdown)');
    dashSheet.getRange(breakdownStartRow + 2, 1).setFontStyle('italic').setFontColor('#666666');
    
    // =========================================================================
    // COLUMN WIDTHS
    // =========================================================================
    dashSheet.setColumnWidth(1, 100);
    for (let i = 2; i <= 7; i++) {
      dashSheet.setColumnWidth(i, 80);
    }
    
    // Store unique values for dynamic dropdowns
    cpStoreDashboardOptions(ss, uniqueValues);
    
    ui.alert('Dashboard V8 Built!', 
      `Advanced Dashboard created with:\n\n` +
      `• ${uniqueValues.strategies.length} Strategies\n` +
      `• ${uniqueValues.subStrategies.length} Sub-Strategies\n` +
      `• ${uniqueValues.spots.length} Spots\n` +
      `• ${uniqueValues.campaigns.length} Campaigns\n` +
      `• ${uniqueValues.countries.length} Countries\n\n` +
      `NEW in V8:\n` +
      `• 3 Filter rows (Field/Mode/Value)\n` +
      `• Comparison date range\n` +
      `• Use "Apply Filters" from menu to refresh data`,
      ui.ButtonSet.OK);
    
    Logger.log('V8 Dashboard built successfully');
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    ui.alert('Error', `Failed to build V8 Dashboard: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Build charts for V8 Dashboard (copy of V7 chart builder with adjusted rows)
 */
function cpBuildDashboardChartsV8(dashSheet, chartRows) {
  // Remove existing charts
  const existingCharts = dashSheet.getCharts();
  for (const chart of existingCharts) {
    dashSheet.removeChart(chart);
  }
  
  // Get metric toggle states (row 12 for V8)
  const toggles = dashSheet.getRange('B13:G13').getValues()[0];  // Row 13 = checkboxes
  const showSpend = toggles[0] === true;
  const showCPA = toggles[1] === true;
  const showConv = toggles[2] === true;
  const showImpr = toggles[3] === true;
  const showClicks = toggles[4] === true;
  const showCTR = toggles[5] === true;
  
  // Get active selections
  const rawSelections = dashSheet.getRange('B3:F3').getDisplayValues()[0];
  const activeSelections = [];
  for (let i = 0; i < 5; i++) {
    const sel = rawSelections[i];
    if (sel && sel !== '(Clear)' && sel.trim() !== '') {
      activeSelections.push({
        index: i,
        name: sel,
        shortName: sel.length > 12 ? sel.substring(0, 10) + '..' : sel
      });
    }
  }
  
  if (activeSelections.length === 0) {
    Logger.log('No active selections - skipping chart build');
    return;
  }
  
  const colorPalette = {
    0: { base: '#2196F3', light: '#64B5F6', dark: '#1565C0' },
    1: { base: '#F44336', light: '#E57373', dark: '#C62828' },
    2: { base: '#4CAF50', light: '#81C784', dark: '#2E7D32' },
    3: { base: '#9C27B0', light: '#BA68C8', dark: '#6A1B9A' },
    4: { base: '#FF9800', light: '#FFB74D', dark: '#E65100' }
  };
  
  const metricColStarts = {
    'Spend': 2, 'Conv': 7, 'CPA': 12,
    'Impr': 17, 'Clicks': 22, 'CTR': 27, 'eCPM': 32
  };
  
  // Update headers with selection names
  const allMetrics = ['Spend', 'Conv', 'CPA', 'Impr', 'Clicks', 'CTR', 'eCPM'];
  for (const metric of allMetrics) {
    const baseCol = metricColStarts[metric];
    for (let s = 0; s < 5; s++) {
      const col = baseCol + s;
      const activeSel = activeSelections.find(a => a.index === s);
      if (activeSel) {
        dashSheet.getRange(24, col).setValue(`${activeSel.shortName} ${metric}`);
      } else {
        dashSheet.getRange(24, col).setValue('');
      }
    }
  }
  
  // Chart 1: Spend, CPA, Conversions (row 23 headers, row 24+ data)
  if (showSpend || showConv || showCPA) {
    const chart1Builder = dashSheet.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(dashSheet.getRange(24, 1, chartRows, 1))
      .setNumHeaders(1)
      .setPosition(1, 9, 0, 0)
      .setOption('title', 'Spend, Conversions & CPA')
      .setOption('width', 700)
      .setOption('height', 400)
      .setOption('legend', { position: 'right', maxLines: 15 })
      .setOption('hAxis', { title: 'Date', slantedText: true, slantedTextAngle: 45 })
      .setOption('vAxes', {
        0: { title: 'Spend / CPA ($)', format: '$#,##0' },
        1: { title: 'Conversions', format: '#,##0' }
      })
      .setOption('useFirstColumnAsDomain', true);
    
    let seriesConfig = {};
    let seriesIndex = 0;
    
    if (showSpend) {
      for (const sel of activeSelections) {
        const col = metricColStarts['Spend'] + sel.index;
        chart1Builder.addRange(dashSheet.getRange(24, col, chartRows, 1));
        seriesConfig[seriesIndex++] = { 
          color: colorPalette[sel.index].base, lineWidth: 4, pointSize: 8,
          pointShape: 'circle', type: 'line', targetAxisIndex: 0 
        };
      }
    }
    
    if (showCPA) {
      for (const sel of activeSelections) {
        const col = metricColStarts['CPA'] + sel.index;
        chart1Builder.addRange(dashSheet.getRange(24, col, chartRows, 1));
        seriesConfig[seriesIndex++] = { 
          color: colorPalette[sel.index].dark, lineWidth: 2, pointSize: 6,
          pointShape: 'square', type: 'line', targetAxisIndex: 0
        };
      }
    }
    
    if (showConv) {
      for (const sel of activeSelections) {
        const col = metricColStarts['Conv'] + sel.index;
        chart1Builder.addRange(dashSheet.getRange(24, col, chartRows, 1));
        seriesConfig[seriesIndex++] = { 
          color: colorPalette[sel.index].light, lineWidth: 1, pointSize: 5,
          pointShape: 'triangle', type: 'line', targetAxisIndex: 1
        };
      }
    }
    
    if (seriesIndex > 0) {
      chart1Builder.setOption('series', seriesConfig);
      dashSheet.insertChart(chart1Builder.build());
    }
  }
  
  // Chart 2: Impressions, Clicks, CTR
  if (showImpr || showClicks || showCTR) {
    const chart2Builder = dashSheet.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(dashSheet.getRange(24, 1, chartRows, 1))
      .setNumHeaders(1)
      .setPosition(23, 9, 0, 0)
      .setOption('title', 'Impressions, Clicks & CTR')
      .setOption('width', 700)
      .setOption('height', 400)
      .setOption('legend', { position: 'right', maxLines: 15 })
      .setOption('hAxis', { title: 'Date', slantedText: true, slantedTextAngle: 45 })
      .setOption('vAxes', {
        0: { title: 'Impressions / Clicks', format: '#,##0' },
        1: { title: 'CTR (%)', format: '0.00"%"' }
      })
      .setOption('useFirstColumnAsDomain', true);
    
    let seriesConfig2 = {};
    let seriesIndex2 = 0;
    
    if (showImpr) {
      for (const sel of activeSelections) {
        const col = metricColStarts['Impr'] + sel.index;
        chart2Builder.addRange(dashSheet.getRange(24, col, chartRows, 1));
        seriesConfig2[seriesIndex2++] = { 
          color: colorPalette[sel.index].light, lineWidth: 1, pointSize: 5,
          pointShape: 'triangle', type: 'line', targetAxisIndex: 0
        };
      }
    }
    
    if (showClicks) {
      for (const sel of activeSelections) {
        const col = metricColStarts['Clicks'] + sel.index;
        chart2Builder.addRange(dashSheet.getRange(24, col, chartRows, 1));
        seriesConfig2[seriesIndex2++] = { 
          color: colorPalette[sel.index].dark, lineWidth: 2, pointSize: 6,
          pointShape: 'square', type: 'line', targetAxisIndex: 0
        };
      }
    }
    
    if (showCTR) {
      for (const sel of activeSelections) {
        const col = metricColStarts['CTR'] + sel.index;
        chart2Builder.addRange(dashSheet.getRange(24, col, chartRows, 1));
        seriesConfig2[seriesIndex2++] = { 
          color: colorPalette[sel.index].base, lineWidth: 4, pointSize: 8,
          pointShape: 'circle', type: 'line', targetAxisIndex: 1
        };
      }
    }
    
    if (seriesIndex2 > 0) {
      chart2Builder.setOption('series', seriesConfig2);
      dashSheet.insertChart(chart2Builder.build());
    }
  }
  
  // Chart 3: eCPM (Average Cost Per Mille)
  // Always show eCPM chart if we have active selections
  if (activeSelections.length > 0) {
    const chart3Builder = dashSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dashSheet.getRange(24, 1, chartRows, 1))  // Date column
      .setNumHeaders(1)
      .setPosition(45, 9, 0, 0)  // Position below Chart 2
      .setOption('title', 'Average eCPM (Cost Per 1000 Impressions)')
      .setOption('width', 700)
      .setOption('height', 350)
      .setOption('legend', { position: 'right', maxLines: 10 })
      .setOption('hAxis', { title: 'Date', slantedText: true, slantedTextAngle: 45 })
      .setOption('vAxis', { title: 'eCPM ($)', format: '$#,##0.000' })
      .setOption('useFirstColumnAsDomain', true);
    
    let seriesConfig3 = {};
    let seriesIndex3 = 0;
    
    for (const sel of activeSelections) {
      const col = metricColStarts['eCPM'] + sel.index;
      chart3Builder.addRange(dashSheet.getRange(24, col, chartRows, 1));
      seriesConfig3[seriesIndex3++] = { 
        color: colorPalette[sel.index].base, 
        lineWidth: 3, 
        pointSize: 6,
        pointShape: 'circle'
      };
    }
    
    if (seriesIndex3 > 0) {
      chart3Builder.setOption('series', seriesConfig3);
      dashSheet.insertChart(chart3Builder.build());
    }
  }
}

/**
 * Refresh V8 Dashboard - quick refresh of charts and breakdown
 * Use this when you've changed selections/filters and want to force a refresh
 */
function cpRefreshDashboardV8() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const dashSheet = ss.getSheetByName(CP_DASHBOARD_SHEET_NAME);
  if (!dashSheet) {
    ui.alert('Error', 'Dashboard not found. Please build Dashboard V8 first.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    Logger.log('Refreshing V8 Dashboard...');
    
    // Rebuild charts with current settings (data starts row 25)
    let dataRowCount = 0;
    for (let r = 25; r <= 42; r++) {
      if (dashSheet.getRange(r, 1).getValue()) dataRowCount++;
    }
    if (dataRowCount === 0) dataRowCount = 7;
    
    cpBuildDashboardChartsV8(dashSheet, dataRowCount + 1);
    
    // Also run apply filters to update breakdown
    cpApplyDashboardFilters();
    
    Logger.log('V8 Dashboard refreshed');
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    ui.alert('Error', `Failed to refresh: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Apply filters to V8 Dashboard
 * Reads filter settings and refreshes data with filtered results
 */
function cpApplyDashboardFilters() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const dashSheet = ss.getSheetByName(CP_DASHBOARD_SHEET_NAME);
  if (!dashSheet) {
    ui.alert('Error', 'Dashboard not found. Please build Dashboard V8 first.', ui.ButtonSet.OK);
    return;
  }
  
  const spotSheet = ss.getSheetByName(CP_DAILY_SPOT_SHEET_NAME);
  if (!spotSheet || spotSheet.getLastRow() < 2) {
    ui.alert('Error', 'No spot-level data found. Please pull data first.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    Logger.log('Applying V8 Dashboard filters...');
    
    // Read filter settings (filters on rows 9, 10, 11)
    const filters = [];
    for (let i = 0; i < 3; i++) {
      const row = 9 + i;
      const field = dashSheet.getRange(row, 2).getValue();
      const mode = dashSheet.getRange(row, 3).getValue();
      const value = String(dashSheet.getRange(row, 4).getValue()).trim();
      
      if (field && field !== '(None)' && value) {
        filters.push({ field, mode, value });
      }
    }
    
    Logger.log(`Active filters: ${filters.length}`);
    filters.forEach((f, i) => Logger.log(`  Filter ${i+1}: ${f.field} ${f.mode} "${f.value}"`));
    
    // Read date ranges (Row 5 = Primary, Row 6 = Comparison)
    const primaryStart = dashSheet.getRange('B5').getValue();
    const primaryEnd = dashSheet.getRange('D5').getValue();
    const compStart = dashSheet.getRange('B6').getValue();
    const compEnd = dashSheet.getRange('D6').getValue();
    
    const hasComparison = compStart && compEnd;
    
    Logger.log(`Primary: ${primaryStart} to ${primaryEnd}`);
    if (hasComparison) {
      Logger.log(`Comparison: ${compStart} to ${compEnd}`);
    }
    
    // Get all spot data
    const spotLastRow = spotSheet.getLastRow();
    const spotData = spotSheet.getRange(2, 1, spotLastRow - 1, 16).getValues();
    
    // Column indices in spot data (0-based)
    // A=Date(0), B=Strategy(1), C=SubStrategy(2), D=CampID(3), E=CampName(4), F=SpotID(5), G=SpotName(6), H=Country(7)
    // I=BidID(8), J=Impr(9), K=Clicks(10), L=Conv(11), M=Spend(12), N=CPA(13), O=eCPM(14), P=CTR(15)
    
    const fieldColMap = {
      'Tier 1 Strategy': 1,
      'Sub-Strategy': 2,
      'Campaign': 4,
      'Spot Name': 6,
      'Country': 7
    };
    
    // Filter function
    const matchesFilters = (row) => {
      for (const filter of filters) {
        const colIndex = fieldColMap[filter.field];
        if (colIndex === undefined) continue;
        
        const cellValue = String(row[colIndex] || '').toLowerCase();
        const filterValue = filter.value.toLowerCase();
        
        if (filter.mode === 'equals') {
          if (cellValue !== filterValue) return false;
        } else {  // contains
          if (!cellValue.includes(filterValue)) return false;
        }
      }
      return true;
    };
    
    // Date filter function
    const inDateRange = (rowDate, start, end) => {
      if (!rowDate || !start || !end) return false;
      const d = new Date(rowDate);
      return d >= new Date(start) && d <= new Date(end);
    };
    
    // Calculate metrics for primary period
    let primaryMetrics = { spend: 0, conv: 0, impr: 0, clicks: 0 };
    let compMetrics = { spend: 0, conv: 0, impr: 0, clicks: 0 };
    
    for (const row of spotData) {
      if (!matchesFilters(row)) continue;
      
      const rowDate = row[0];
      
      if (inDateRange(rowDate, primaryStart, primaryEnd)) {
        primaryMetrics.spend += Number(row[12]) || 0;
        primaryMetrics.conv += Number(row[11]) || 0;
        primaryMetrics.impr += Number(row[9]) || 0;
        primaryMetrics.clicks += Number(row[10]) || 0;
      }
      
      if (hasComparison && inDateRange(rowDate, compStart, compEnd)) {
        compMetrics.spend += Number(row[12]) || 0;
        compMetrics.conv += Number(row[11]) || 0;
        compMetrics.impr += Number(row[9]) || 0;
        compMetrics.clicks += Number(row[10]) || 0;
      }
    }
    
    // Calculate CPA
    primaryMetrics.cpa = primaryMetrics.conv > 0 ? primaryMetrics.spend / primaryMetrics.conv : 0;
    compMetrics.cpa = compMetrics.conv > 0 ? compMetrics.spend / compMetrics.conv : 0;
    
    // Update comparison section (row 21)
    if (hasComparison) {
      const deltaSpend = primaryMetrics.spend - compMetrics.spend;
      const deltaConv = primaryMetrics.conv - compMetrics.conv;
      const deltaCPA = primaryMetrics.cpa - compMetrics.cpa;
      const pctSpend = compMetrics.spend > 0 ? (deltaSpend / compMetrics.spend * 100) : 0;
      const pctConv = compMetrics.conv > 0 ? (deltaConv / compMetrics.conv * 100) : 0;
      
      let status = '→ Stable';
      if (deltaSpend > 0 && deltaConv > 0) status = '📈 Growing';
      else if (deltaSpend < 0 && deltaConv < 0) status = '📉 Declining';
      else if (deltaCPA < 0) status = '✅ More Efficient';
      else if (deltaCPA > 0) status = '⚠️ Less Efficient';
      
      dashSheet.getRange('A21:G21').breakApart();
      dashSheet.getRange('A21:G21').setValues([[
        compMetrics.spend,
        compMetrics.conv,
        compMetrics.cpa,
        `${deltaSpend >= 0 ? '+' : ''}${pctSpend.toFixed(1)}%`,
        `${deltaConv >= 0 ? '+' : ''}${pctConv.toFixed(1)}%`,
        `${deltaCPA >= 0 ? '+' : ''}$${deltaCPA.toFixed(2)}`,
        status
      ]]);
      dashSheet.getRange('A21').setNumberFormat('$#,##0.00');
      dashSheet.getRange('B21').setNumberFormat('#,##0');
      dashSheet.getRange('C21').setNumberFormat('$#,##0.00');
      
      // Color the delta cells
      const spendCell = dashSheet.getRange('D21');
      const convCell = dashSheet.getRange('E21');
      const cpaCell = dashSheet.getRange('F21');
      
      spendCell.setBackground(deltaSpend >= 0 ? '#c8e6c9' : '#ffcdd2');
      convCell.setBackground(deltaConv >= 0 ? '#c8e6c9' : '#ffcdd2');
      cpaCell.setBackground(deltaCPA <= 0 ? '#c8e6c9' : '#ffcdd2');  // Lower CPA is better
    } else {
      dashSheet.getRange('A21:G21').breakApart();
      dashSheet.getRange('A21').setValue('(No comparison period set)');
      dashSheet.getRange('A21:G21').merge();
      dashSheet.getRange('A21').setFontStyle('italic').setFontColor('#666666');
    }
    
    // Rebuild charts
    const lastDataRow = dashSheet.getLastRow();
    let dataRowCount = 0;
    for (let r = 25; r <= Math.min(lastDataRow, 42); r++) {
      if (dashSheet.getRange(r, 1).getValue()) dataRowCount++;
    }
    if (dataRowCount === 0) dataRowCount = 7;
    
    cpBuildDashboardChartsV8(dashSheet, dataRowCount + 1);
    
    // =========================================================================
    // BREAKDOWN TABLE - Populate when Selection 1 = "(All)"
    // Shows breakdown by "Breakdown By" selection for all items matching filters
    // =========================================================================
    const selection1 = dashSheet.getRange('B3').getValue();
    const breakdownBy = dashSheet.getRange('B4').getValue() || 'Spot Name';  // Read from Breakdown By dropdown (Row 4)
    const breakdownStartRow = 66;
    
    // Clear existing breakdown data (rows 67+)
    const breakdownDataStart = breakdownStartRow + 2;
    dashSheet.getRange(breakdownDataStart, 1, 50, 9).clearContent();
    
    if (selection1 === '(All)' && breakdownBy) {
      Logger.log(`Building breakdown by: ${breakdownBy}`);
      
      // Group data by breakdown level
      const breakdownColMap = {
        'Tier 1 Strategy': 1,
        'Sub-Strategy': 2,
        'Campaign': 4,
        'Spot Name': 6,
        'Country': 7
      };
      
      const groupCol = breakdownColMap[breakdownBy];
      const grouped = {};
      const compGrouped = {};
      
      for (const row of spotData) {
        if (!matchesFilters(row)) continue;
        
        const groupKey = String(row[groupCol] || 'Unknown');
        const rowDate = row[0];
        
        // Primary period aggregation
        if (inDateRange(rowDate, primaryStart, primaryEnd)) {
          if (!grouped[groupKey]) {
            grouped[groupKey] = { impr: 0, clicks: 0, conv: 0, spend: 0 };
          }
          grouped[groupKey].impr += Number(row[9]) || 0;
          grouped[groupKey].clicks += Number(row[10]) || 0;
          grouped[groupKey].conv += Number(row[11]) || 0;
          grouped[groupKey].spend += Number(row[12]) || 0;
        }
        
        // Comparison period aggregation
        if (hasComparison && inDateRange(rowDate, compStart, compEnd)) {
          if (!compGrouped[groupKey]) {
            compGrouped[groupKey] = { impr: 0, clicks: 0, conv: 0, spend: 0 };
          }
          compGrouped[groupKey].impr += Number(row[9]) || 0;
          compGrouped[groupKey].clicks += Number(row[10]) || 0;
          compGrouped[groupKey].conv += Number(row[11]) || 0;
          compGrouped[groupKey].spend += Number(row[12]) || 0;
        }
      }
      
      // Convert to array and sort by spend descending
      const breakdownRows = Object.entries(grouped).map(([name, metrics]) => {
        const cpa = metrics.conv > 0 ? metrics.spend / metrics.conv : 0;
        const ecpm = metrics.impr > 0 ? (metrics.spend / metrics.impr) * 1000 : 0;
        const ctr = metrics.impr > 0 ? (metrics.clicks / metrics.impr) * 100 : 0;
        
        // Calculate comparison delta
        let comparisonStr = '-';
        if (hasComparison && compGrouped[name]) {
          const compSpend = compGrouped[name].spend;
          if (compSpend > 0) {
            const delta = ((metrics.spend - compSpend) / compSpend * 100);
            comparisonStr = `${delta >= 0 ? '+' : ''}${delta.toFixed(1)}%`;
          }
        }
        
        return {
          name,
          impr: metrics.impr,
          clicks: metrics.clicks,
          conv: metrics.conv,
          spend: metrics.spend,
          cpa,
          ecpm,
          ctr,
          comparison: comparisonStr
        };
      }).sort((a, b) => b.spend - a.spend);
      
      // Write breakdown data (limit to 30 rows)
      const maxRows = Math.min(breakdownRows.length, 30);
      if (maxRows > 0) {
        const dataToWrite = breakdownRows.slice(0, maxRows).map(r => [
          r.name, r.impr, r.clicks, r.conv, r.spend, r.cpa, r.ecpm, r.ctr, r.comparison
        ]);
        
        dashSheet.getRange(breakdownDataStart, 1, maxRows, 9).setValues(dataToWrite);
        
        // Format breakdown columns
        dashSheet.getRange(breakdownDataStart, 2, maxRows, 1).setNumberFormat('#,##0');      // Impr
        dashSheet.getRange(breakdownDataStart, 3, maxRows, 1).setNumberFormat('#,##0');      // Clicks
        dashSheet.getRange(breakdownDataStart, 4, maxRows, 1).setNumberFormat('#,##0');      // Conv
        dashSheet.getRange(breakdownDataStart, 5, maxRows, 1).setNumberFormat('$#,##0.00'); // Spend
        dashSheet.getRange(breakdownDataStart, 6, maxRows, 1).setNumberFormat('$#,##0.00'); // CPA
        dashSheet.getRange(breakdownDataStart, 7, maxRows, 1).setNumberFormat('$#,##0.000'); // eCPM
        dashSheet.getRange(breakdownDataStart, 8, maxRows, 1).setNumberFormat('0.00"%"');   // CTR
        
        // Color comparison column
        for (let i = 0; i < maxRows; i++) {
          const cell = dashSheet.getRange(breakdownDataStart + i, 9);
          const val = breakdownRows[i].comparison;
          if (val.startsWith('+')) {
            cell.setBackground('#c8e6c9');  // Green for positive
          } else if (val.startsWith('-') && val !== '-') {
            cell.setBackground('#ffcdd2');  // Red for negative
          }
        }
        
        Logger.log(`Wrote ${maxRows} breakdown rows`);
      }
      
      // Update breakdown header to show what we're showing
      dashSheet.getRange(`A${breakdownStartRow}`).setValue(`BREAKDOWN by ${breakdownBy} (${maxRows} items matching filters)`);
    } else if (selection1 === '(All)' && !breakdownBy) {
      dashSheet.getRange(breakdownDataStart, 1).setValue('Select a "Breakdown By" option (B4) to see breakdown');
      dashSheet.getRange(breakdownDataStart, 1).setFontStyle('italic').setFontColor('#666666');
    } else {
      dashSheet.getRange(breakdownDataStart, 1).setValue('Select "(All)" in Compare to see breakdown table');
      dashSheet.getRange(breakdownDataStart, 1).setFontStyle('italic').setFontColor('#666666');
    }
    
    ui.alert('Filters Applied!', 
      `Filters: ${filters.length} active\n\n` +
      `Primary Period:\n` +
      `  Spend: $${primaryMetrics.spend.toFixed(2)}\n` +
      `  Conversions: ${primaryMetrics.conv}\n` +
      `  CPA: $${primaryMetrics.cpa.toFixed(2)}\n\n` +
      (hasComparison ? 
        `Comparison Period:\n` +
        `  Spend: $${compMetrics.spend.toFixed(2)}\n` +
        `  Conversions: ${compMetrics.conv}\n` +
        `  CPA: $${compMetrics.cpa.toFixed(2)}` :
        '(No comparison period)'),
      ui.ButtonSet.OK);
    
    Logger.log('V8 Filters applied successfully');
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    ui.alert('Error', `Failed to apply filters: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Get unique values from spot-level daily stats for dashboard dropdowns
 */
function cpGetUniqueValuesForDashboard(spotSheet) {
  const lastRow = spotSheet.getLastRow();
  if (lastRow < 2) {
    return { strategies: [], subStrategies: [], spots: [], campaigns: [], countries: [] };
  }
  
  // Get all data columns we need
  // B=Strategy, C=Sub-Strategy, E=Campaign Name, G=Spot Name, H=Country
  const data = spotSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  
  const strategies = new Set();
  const subStrategies = new Set();
  const campaigns = new Set();
  const spots = new Set();
  const countries = new Set();
  
  for (const row of data) {
    if (row[1]) strategies.add(String(row[1]).trim());      // B - Strategy
    if (row[2]) subStrategies.add(String(row[2]).trim());   // C - Sub-Strategy
    
    // E - Campaign Name (use Campaign ID as fallback if name is empty)
    const campaignName = String(row[4] || '').trim();
    const campaignId = String(row[3] || '').trim();  // D - Campaign ID
    if (campaignName) {
      campaigns.add(campaignName);
    } else if (campaignId) {
      // Fallback: use Campaign ID if name is missing
      campaigns.add(`Campaign ${campaignId}`);
    }
    
    if (row[6]) spots.add(String(row[6]).trim());           // G - Spot Name
    if (row[7]) {                                            // H - Country (may be comma-separated)
      const countryList = String(row[7]).split(',');
      for (const c of countryList) {
        const trimmed = c.trim();
        if (trimmed) countries.add(trimmed);
      }
    }
  }
  
  return {
    strategies: Array.from(strategies).sort(),
    subStrategies: Array.from(subStrategies).sort(),
    spots: Array.from(spots).sort(),
    campaigns: Array.from(campaigns).sort(),
    countries: Array.from(countries).sort()
  };
}

/**
 * Build selection options based on view level
 */
function cpBuildSelectionOptions(uniqueValues, viewLevel) {
  switch (viewLevel) {
    case 'Tier 1 Strategy':
      return uniqueValues.strategies;
    case 'Sub-Strategy':
      return uniqueValues.subStrategies;
    case 'Spot Name':
      return uniqueValues.spots;
    case 'Campaign':
      return uniqueValues.campaigns;
    case 'Country':
      return uniqueValues.countries;
    case 'All':
    default:
      // Combine all options with prefixes
      const all = [];
      for (const s of uniqueValues.strategies) all.push(s);
      for (const s of uniqueValues.subStrategies) all.push(s);
      for (const s of uniqueValues.campaigns) all.push(s);
      for (const s of uniqueValues.spots) all.push(s);
      for (const c of uniqueValues.countries) all.push(c);
      return [...new Set(all)].sort();
  }
}

/**
 * Store dashboard options in a hidden sheet for dynamic dropdowns
 */
function cpStoreDashboardOptions(ss, uniqueValues) {
  let optionsSheet = ss.getSheetByName('_DashboardOptions');
  if (optionsSheet) {
    optionsSheet.clear();
  } else {
    optionsSheet = ss.insertSheet('_DashboardOptions');
  }
  
  // Hide the sheet
  optionsSheet.hideSheet();
  
  // Store each type in separate columns
  // A = Strategies, B = Sub-Strategies, C = Spots, D = Campaigns, E = Countries
  const headers = ['Strategies', 'Sub-Strategies', 'Spots', 'Campaigns', 'Countries'];
  optionsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Write each list
  const maxRows = Math.max(
    uniqueValues.strategies.length,
    uniqueValues.subStrategies.length,
    uniqueValues.spots.length,
    uniqueValues.campaigns.length,
    uniqueValues.countries.length
  );
  
  if (maxRows > 0) {
    const data = [];
    for (let i = 0; i < maxRows; i++) {
      data.push([
        uniqueValues.strategies[i] || '',
        uniqueValues.subStrategies[i] || '',
        uniqueValues.spots[i] || '',
        uniqueValues.campaigns[i] || '',
        uniqueValues.countries[i] || ''
      ]);
    }
    optionsSheet.getRange(2, 1, maxRows, 5).setValues(data);
  }
  
  Logger.log('Stored dashboard options in hidden sheet');
}

// ============================================================================
// LEGACY DASHBOARD FUNCTIONS (V2) - Campaign-level only
// ============================================================================

/**
 * Build or refresh the Dashboard sheet with campaign selector and chart
 * LEGACY: Use cpBuildDashboardV7() for multi-level views
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
    
    Logger.log(`🚀 OPTIMIZED V7 - Processing ${campaignIds.length} campaigns`);
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
// PIVOT TABLE FUNCTIONS (V7) - Native Google Sheets Pivot Table
// ============================================================================

/**
 * STEP 1: Prepare the Pivot Data Source sheet
 * 
 * Creates a clean data source sheet that you can use to manually create
 * a pivot table through the Google Sheets UI.
 * 
 * After running this:
 * 1. Go to the "Pivot Data" sheet
 * 2. Select all data (Ctrl+A)
 * 3. Insert > Pivot table > New sheet
 * 4. Add Row groups: Tier 1 Strategy, Sub Strategy, Campaign Name, Spot Name
 * 5. Add Values: T Spend (SUM), T Conv (SUM), Current eCPM Bid (AVG), Daily Budget (SUM)
 * 6. Then run "Add Edit Columns to Pivot"
 */
function cpPreparePivotSource() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if Control Panel exists and has data
  const cpSheet = ss.getSheetByName(CP_SHEET_NAME);
  if (!cpSheet || cpSheet.getLastRow() < 3) {
    ui.alert('Error', 'Please refresh Control Panel data first.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const lastCpRow = cpSheet.getLastRow();
    const dataRowCount = lastCpRow - 2;  // Subtract header and totals row
    
    Logger.log(`Preparing pivot data source with ${dataRowCount} rows...`);
    
    // Create or clear the Pivot Data sheet
    let pivotDataSheet = ss.getSheetByName('Pivot Data');
    if (pivotDataSheet) {
      pivotDataSheet.clear();
      Logger.log('Cleared existing Pivot Data sheet');
    } else {
      pivotDataSheet = ss.insertSheet('Pivot Data');
      Logger.log('Created Pivot Data sheet');
    }
    
    // Copy header row (row 1) - only the columns needed for pivot
    // A=Strategy, B=Sub-Strategy, C=Campaign, D=Campaign ID, H=Spot, I=Bid, O=Budget, T=T Spend, Z=T Conv, AH=Bid ID
    const headers = cpSheet.getRange(1, 1, 1, 34).getValues()[0];
    pivotDataSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Copy data rows (row 3 onwards, skip totals row 2)
    const dataRows = cpSheet.getRange(3, 1, lastCpRow - 2, 34).getValues();
    pivotDataSheet.getRange(2, 1, dataRows.length, 34).setValues(dataRows);
    
    // Format header row
    pivotDataSheet.getRange(1, 1, 1, 34)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('white');
    
    // Freeze header row
    pivotDataSheet.setFrozenRows(1);
    
    // Auto-resize key columns
    pivotDataSheet.autoResizeColumns(1, 10);
    
    // Activate the sheet
    ss.setActiveSheet(pivotDataSheet);
    
    ui.alert('Pivot Data Ready!', 
      `Created "Pivot Data" sheet with ${dataRowCount} rows.\n\n` +
      `NOW CREATE YOUR PIVOT TABLE:\n` +
      `1. Select cell A1 in this sheet\n` +
      `2. Go to Insert > Pivot table\n` +
      `3. Choose "New sheet" and click Create\n` +
      `4. In the Pivot table editor:\n` +
      `   • Rows: Add "Tier 1 Strategy", "Sub Strategy", "Campaign Name", "Spot Name"\n` +
      `   • Values: Add "T Spend" (SUM), "T Conv" (SUM), "Current eCPM Bid" (AVG), "Daily Budget" (SUM)\n` +
      `5. Rename the pivot table sheet to "${CP_PIVOT_SHEET_NAME}"\n` +
      `6. Then run "Add Edit Columns to Pivot"`,
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    ui.alert('Error', `Failed to prepare pivot data: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * STEP 2: Add Edit Columns to an existing Pivot Table
 * 
 * Detects your manually-created pivot table and adds:
 * - Bid ID lookup formula (matches pivot row to Control Panel data)
 * - Campaign ID lookup formula
 * - New CPM column (editable)
 * - New Budget column (editable)
 * - Comment column (editable)
 */
function cpAddEditColumnsToPivot() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check for pivot sheet
  let pivotSheet = ss.getSheetByName(CP_PIVOT_SHEET_NAME);
  if (!pivotSheet) {
    ui.alert('Error', 
      `Sheet "${CP_PIVOT_SHEET_NAME}" not found.\n\n` +
      `Please either:\n` +
      `1. Rename your pivot table sheet to "${CP_PIVOT_SHEET_NAME}"\n` +
      `2. Or create the pivot table first using "Prepare Pivot Data Source"`,
      ui.ButtonSet.OK);
    return;
  }
  
  try {
    // Check if there's a pivot table on this sheet
    const pivotTables = pivotSheet.getPivotTables();
    if (pivotTables.length === 0) {
      ui.alert('Error', 
        `No pivot table found on "${CP_PIVOT_SHEET_NAME}" sheet.\n\n` +
        `Please create a pivot table first, then run this function.`,
        ui.ButtonSet.OK);
      return;
    }
    
    Logger.log(`Found ${pivotTables.length} pivot table(s) on sheet`);
    
    // Get the pivot table and its source range to calculate actual row count
    const pivotTable = pivotTables[0];
    const sourceRange = pivotTable.getSourceDataRange();
    const sourceSheet = sourceRange.getSheet();
    const sourceRowCount = sourceRange.getNumRows() - 1;  // Minus header
    
    Logger.log(`Pivot source: ${sourceSheet.getName()}, ${sourceRowCount} data rows`);
    
    // Try to expand/show all rows
    const maxRows = pivotSheet.getMaxRows();
    if (maxRows > 1) {
      pivotSheet.showRows(1, maxRows);
    }
    SpreadsheetApp.flush();
    
    // Get pivot dimensions - use getDataRange for more accurate count
    const pivotLastCol = pivotSheet.getLastColumn();
    let pivotLastRow = pivotSheet.getLastRow();
    
    // If pivot is collapsed, the visible rows will be much less than source data
    // In that case, estimate based on source data + subtotals
    // A fully expanded pivot should have at least sourceRowCount rows plus subtotals
    // We'll add formulas for MORE rows than currently visible to cover expansion
    const estimatedMaxRows = sourceRowCount + Math.ceil(sourceRowCount * 0.3) + 10;  // Add 30% for subtotals + buffer
    
    // Use the larger of visible rows or estimated rows
    const formulaRowCount = Math.max(pivotLastRow, estimatedMaxRows);
    
    Logger.log(`Visible rows: ${pivotLastRow}, Estimated max: ${estimatedMaxRows}, Using: ${formulaRowCount}`);
    
    Logger.log(`Pivot table spans columns 1-${pivotLastCol}, rows 1-${pivotLastRow}`);
    
    // Determine where to add edit columns (after a spacer)
    // Order: Spacer → Checkbox → New CPM → New Budget → Comment → Bid ID → Campaign ID
    // Daily-use columns first, admin/reference columns last
    const spacerCol = pivotLastCol + 1;
    const checkboxCol = pivotLastCol + 2;   // Dashboard navigation checkbox
    const newCpmCol = pivotLastCol + 3;     // Editable
    const newBudgetCol = pivotLastCol + 4;  // Editable
    const commentCol = pivotLastCol + 5;    // Editable
    const bidIdCol = pivotLastCol + 6;      // Reference/Admin
    const campIdCol = pivotLastCol + 7;     // Reference/Admin
    
    // Clear any existing edit columns (in case re-running)
    if (pivotSheet.getMaxColumns() > pivotLastCol) {
      // Check if edit columns already exist by looking for headers we know
      const existingHeaders = pivotSheet.getRange(1, pivotLastCol + 1, 1, 10).getValues()[0];
      const hasEditCols = existingHeaders.includes('Bid ID') || existingHeaders.includes('New CPM') || existingHeaders.includes('📊');
      if (hasEditCols) {
        // Clear existing edit columns
        pivotSheet.getRange(1, pivotLastCol + 1, formulaRowCount + 10, 10).clear();
        // Also clear any data validation (checkboxes)
        pivotSheet.getRange(2, pivotLastCol + 1, formulaRowCount + 10, 10).clearDataValidations();
        Logger.log('Cleared existing edit columns');
      }
    }
    
    // Headers for edit columns
    // Order: Spacer, Checkbox, New CPM, New Budget, Comment, Bid ID, Campaign ID
    const editHeaders = ['', '📊', 'New CPM', 'New Budget', 'Comment', 'Bid ID', 'Campaign ID'];
    pivotSheet.getRange(1, spacerCol, 1, editHeaders.length).setValues([editHeaders]);
    
    // Format editable column headers (Checkbox, New CPM, New Budget, Comment)
    pivotSheet.getRange(1, checkboxCol, 1, 4)
      .setFontWeight('bold')
      .setBackground('#ff9800')
      .setFontColor('white');
    
    // Format admin column headers (Bid ID, Campaign ID) - different color
    pivotSheet.getRange(1, bidIdCol, 1, 2)
      .setFontWeight('bold')
      .setBackground('#9e9e9e')
      .setFontColor('white');
    
    // Determine pivot table row group columns
    // We need to find which columns contain the row labels (Strategy, Sub-Strategy, Campaign, Spot)
    const headerRow = pivotSheet.getRange(1, 1, 1, pivotLastCol).getValues()[0];
    Logger.log(`Pivot headers: ${headerRow.join(', ')}`);
    
    // Find the row group columns (they typically have names like "Tier 1 Strategy", "Sub Strategy", etc.)
    let strategyCol = -1, subStratCol = -1, campaignCol = -1, spotCol = -1;
    
    for (let i = 0; i < headerRow.length; i++) {
      const h = String(headerRow[i]).toLowerCase();
      if (h.includes('strategy') && !h.includes('sub')) strategyCol = i + 1;
      else if (h.includes('sub') && h.includes('strategy')) subStratCol = i + 1;
      else if (h.includes('campaign') && h.includes('name')) campaignCol = i + 1;
      else if (h.includes('spot') && h.includes('name')) spotCol = i + 1;
    }
    
    Logger.log(`Row group columns: Strategy=${strategyCol}, SubStrat=${subStratCol}, Campaign=${campaignCol}, Spot=${spotCol}`);
    
    // If we couldn't detect columns, use defaults (A, B, C, D)
    if (strategyCol === -1) strategyCol = 1;
    if (subStratCol === -1) subStratCol = 2;
    if (campaignCol === -1) campaignCol = 3;
    if (spotCol === -1) spotCol = 4;
    
    // Build lookup formulas for each row
    // These match the pivot row labels back to Control Panel to get Bid ID
    // Use formulaRowCount to ensure we cover all rows even when pivot is collapsed
    if (formulaRowCount > 1) {
      const formulas = [];
      const sCol = String.fromCharCode(64 + strategyCol);
      const ssCol = String.fromCharCode(64 + subStratCol);
      const cCol = String.fromCharCode(64 + campaignCol);
      const spCol = String.fromCharCode(64 + spotCol);
      
      for (let row = 2; row <= formulaRowCount; row++) {
        // Bid ID lookup - matches all 4 grouping columns to find unique bid
        // NOTE: Enable "Repeat row labels" in pivot table settings for this to work
        const bidFormula = `=IFERROR(INDEX('${CP_SHEET_NAME}'!$AH$3:$AH,MATCH(1,('${CP_SHEET_NAME}'!$A$3:$A=${sCol}${row})*('${CP_SHEET_NAME}'!$B$3:$B=${ssCol}${row})*('${CP_SHEET_NAME}'!$C$3:$C=${cCol}${row})*('${CP_SHEET_NAME}'!$H$3:$H=${spCol}${row}),0)),"")`;
        
        // Campaign ID lookup - same matching logic
        const campFormula = `=IFERROR(INDEX('${CP_SHEET_NAME}'!$D$3:$D,MATCH(1,('${CP_SHEET_NAME}'!$A$3:$A=${sCol}${row})*('${CP_SHEET_NAME}'!$B$3:$B=${ssCol}${row})*('${CP_SHEET_NAME}'!$C$3:$C=${cCol}${row})*('${CP_SHEET_NAME}'!$H$3:$H=${spCol}${row}),0)),"")`;
        
        formulas.push([bidFormula, campFormula]);
      }
      
      // Set all formulas at once (batch operation)
      pivotSheet.getRange(2, bidIdCol, formulaRowCount - 1, 2).setFormulas(formulas);
      Logger.log(`Added ${formulas.length} lookup formulas for Bid ID and Campaign ID`);
    }
    
    // Add checkbox data validation for Dashboard navigation column
    const checkboxRule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .setAllowInvalid(false)
      .build();
    pivotSheet.getRange(2, checkboxCol, formulaRowCount - 1, 1).setDataValidation(checkboxRule);
    
    // Format columns - use formulaRowCount to cover all potential rows
    // Gray out ID columns (reference only - at the end)
    pivotSheet.getRange(2, bidIdCol, formulaRowCount - 1, 2).setFontColor('#999999');
    
    // Yellow background for editable columns (New CPM, New Budget, Comment)
    pivotSheet.getRange(2, newCpmCol, formulaRowCount - 1, 3).setBackground('#fff9c4');
    
    // Number formats
    pivotSheet.getRange(2, newCpmCol, formulaRowCount - 1, 1).setNumberFormat('$#,##0.000');
    pivotSheet.getRange(2, newBudgetCol, formulaRowCount - 1, 1).setNumberFormat('$#,##0.00');
    
    // Auto-resize edit columns
    pivotSheet.autoResizeColumns(checkboxCol, 6);
    
    // Set checkbox column width (narrower)
    pivotSheet.setColumnWidth(checkboxCol, 30);
    
    ui.alert('Edit Columns Added!', 
      `Added edit columns to your pivot table:\n\n` +
      `• 📊 Checkbox - Click to view campaign on Dashboard\n` +
      `• New CPM - Enter new bid values here\n` +
      `• New Budget - Enter new budget values here\n` +
      `• Comment - Optional notes\n` +
      `• Bid ID / Campaign ID - Reference (admin)\n\n` +
      `To update bids/budgets:\n` +
      `1. Enter values in New CPM or New Budget columns\n` +
      `2. Run "UPDATE FROM PIVOT"\n\n` +
      `To view campaign graph:\n` +
      `Click the checkbox in the 📊 column`,
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    ui.alert('Error', `Failed to add edit columns: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Refresh the edit columns on an existing pivot table
 * 
 * This is faster than "Add Edit Columns" because it:
 * - Keeps the existing column structure
 * - Only refreshes the Bid ID and Campaign ID formulas
 * - Preserves your New CPM, New Budget, and Comment values
 * 
 * Use after refreshing Control Panel data to update the lookups.
 */
function cpRefreshPivotEditColumns() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check for pivot sheet
  const pivotSheet = ss.getSheetByName(CP_PIVOT_SHEET_NAME);
  if (!pivotSheet) {
    ui.alert('Error', `Sheet "${CP_PIVOT_SHEET_NAME}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  try {
    // Find the edit columns by looking for "Bid ID" header
    const lastCol = pivotSheet.getLastColumn();
    const headers = pivotSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    let bidIdCol = -1;
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === 'Bid ID') {
        bidIdCol = i + 1;
        break;
      }
    }
    
    if (bidIdCol === -1) {
      ui.alert('Error', 
        'Edit columns not found. Please run "Add Edit Columns to Pivot" first.',
        ui.ButtonSet.OK);
      return;
    }
    
    const campIdCol = bidIdCol + 1;
    
    // Find the row group columns (before Bid ID, after the values)
    // Look for Strategy, Sub-Strategy, Campaign, Spot columns
    let strategyCol = -1, subStratCol = -1, campaignCol = -1, spotCol = -1;
    
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i]).toLowerCase();
      if (h.includes('strategy') && !h.includes('sub')) strategyCol = i + 1;
      else if (h.includes('sub') && h.includes('strategy')) subStratCol = i + 1;
      else if (h.includes('campaign') && h.includes('name')) campaignCol = i + 1;
      else if (h.includes('spot') && h.includes('name')) spotCol = i + 1;
    }
    
    // Default to A, B, C, D if not found
    if (strategyCol === -1) strategyCol = 1;
    if (subStratCol === -1) subStratCol = 2;
    if (campaignCol === -1) campaignCol = 3;
    if (spotCol === -1) spotCol = 4;
    
    // Get source data row count for formula generation
    const pivotTables = pivotSheet.getPivotTables();
    let formulaRowCount;
    
    if (pivotTables.length > 0) {
      const sourceRange = pivotTables[0].getSourceDataRange();
      const sourceRowCount = sourceRange.getNumRows() - 1;
      formulaRowCount = sourceRowCount + Math.ceil(sourceRowCount * 0.3) + 10;
    } else {
      // Fallback: use visible last row + buffer
      formulaRowCount = pivotSheet.getLastRow() + 50;
    }
    
    Logger.log(`Refreshing formulas for up to ${formulaRowCount} rows...`);
    
    // Generate and set the formulas
    const formulas = [];
    const sCol = String.fromCharCode(64 + strategyCol);
    const ssCol = String.fromCharCode(64 + subStratCol);
    const cCol = String.fromCharCode(64 + campaignCol);
    const spCol = String.fromCharCode(64 + spotCol);
    
    for (let row = 2; row <= formulaRowCount; row++) {
      const bidFormula = `=IFERROR(INDEX('${CP_SHEET_NAME}'!$AH$3:$AH,MATCH(1,('${CP_SHEET_NAME}'!$A$3:$A=${sCol}${row})*('${CP_SHEET_NAME}'!$B$3:$B=${ssCol}${row})*('${CP_SHEET_NAME}'!$C$3:$C=${cCol}${row})*('${CP_SHEET_NAME}'!$H$3:$H=${spCol}${row}),0)),"")`;
      
      const campFormula = `=IFERROR(INDEX('${CP_SHEET_NAME}'!$D$3:$D,MATCH(1,('${CP_SHEET_NAME}'!$A$3:$A=${sCol}${row})*('${CP_SHEET_NAME}'!$B$3:$B=${ssCol}${row})*('${CP_SHEET_NAME}'!$C$3:$C=${cCol}${row})*('${CP_SHEET_NAME}'!$H$3:$H=${spCol}${row}),0)),"")`;
      
      formulas.push([bidFormula, campFormula]);
    }
    
    // Set formulas (only for Bid ID and Campaign ID columns - preserves edit columns)
    pivotSheet.getRange(2, bidIdCol, formulaRowCount - 1, 2).setFormulas(formulas);
    
    // Re-apply formatting to ID columns
    pivotSheet.getRange(2, bidIdCol, formulaRowCount - 1, 2).setFontColor('#999999');
    
    ui.alert('Refreshed!', 
      `Updated ${formulas.length} Bid ID and Campaign ID formulas.\n\n` +
      `Your New CPM, New Budget, and Comment values have been preserved.`,
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    ui.alert('Error', `Failed to refresh edit columns: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * LEGACY: Build pivot table programmatically (kept for reference)
 * Use cpPreparePivotSource + cpAddEditColumnsToPivot instead for better styling
 */
function cpBuildPivotView() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if Control Panel exists and has data
  const cpSheet = ss.getSheetByName(CP_SHEET_NAME);
  if (!cpSheet || cpSheet.getLastRow() < 3) {
    ui.alert('Error', 'Please refresh Control Panel data first.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const lastCpRow = cpSheet.getLastRow();
    const dataRowCount = lastCpRow - 2;  // Subtract header and totals row
    
    Logger.log(`Building native pivot table from ${dataRowCount} Control Panel rows...`);
    
    // Get or create pivot sheet
    let pivotSheet = ss.getSheetByName(CP_PIVOT_SHEET_NAME);
    if (pivotSheet) {
      // Delete existing pivot tables on this sheet
      const existingPivots = pivotSheet.getPivotTables();
      for (const pt of existingPivots) {
        pt.remove();
      }
      pivotSheet.clear();
      Logger.log('Cleared existing Pivot Table sheet');
    } else {
      pivotSheet = ss.insertSheet(CP_PIVOT_SHEET_NAME);
      Logger.log('Created Pivot Table sheet');
    }
    
    // For native pivot table, we need to exclude the Totals row (row 2)
    // Solution: Copy data to a temp range without the totals row, or filter it out
    // We'll create a helper data range that skips row 2
    
    // Create a temporary data sheet for pivot source (without totals row)
    let tempSheet = ss.getSheetByName('_PivotSource');
    if (tempSheet) {
      ss.deleteSheet(tempSheet);
    }
    tempSheet = ss.insertSheet('_PivotSource');
    
    // Copy header row (row 1)
    const headerRow = cpSheet.getRange(1, 1, 1, 34).getValues();
    tempSheet.getRange(1, 1, 1, 34).setValues(headerRow);
    
    // Copy data rows (row 3 onwards, skip totals row 2)
    const dataRows = cpSheet.getRange(3, 1, lastCpRow - 2, 34).getValues();
    tempSheet.getRange(2, 1, dataRows.length, 34).setValues(dataRows);
    
    Logger.log(`Created temp pivot source with ${dataRows.length} data rows`);
    
    // Define source range from temp sheet (header + data, no totals)
    const sourceRange = tempSheet.getRange(1, 1, dataRows.length + 1, 34);
    
    // Create native pivot table starting at A1
    const pivotTable = pivotSheet.getRange('A1').createPivotTable(sourceRange);
    
    // Add ROW groupings (hierarchical) - 1-indexed column numbers
    // Strategy (Column A = 1)
    const strategyGroup = pivotTable.addRowGroup(1);
    strategyGroup.showTotals(true);
    
    // Sub Strategy (Column B = 2)
    const subStratGroup = pivotTable.addRowGroup(2);
    subStratGroup.showTotals(true);
    
    // Campaign Name (Column C = 3)
    const campaignGroup = pivotTable.addRowGroup(3);
    campaignGroup.showTotals(false);  // Don't subtotal per campaign
    
    // Spot Name (Column H = 8)
    const spotGroup = pivotTable.addRowGroup(8);
    spotGroup.showTotals(false);  // Individual bids, no subtotal needed
    
    // Add VALUE columns (metrics) - these will aggregate for subtotals
    // T Spend (Column T = 20)
    pivotTable.addPivotValue(20, SpreadsheetApp.PivotTableSummarizeFunction.SUM)
      .setDisplayName('T Spend');
    
    // T Conv (Column Z = 26)
    pivotTable.addPivotValue(26, SpreadsheetApp.PivotTableSummarizeFunction.SUM)
      .setDisplayName('T Conv');
    
    // Current eCPM Bid (Column I = 9) - Average for subtotals
    pivotTable.addPivotValue(9, SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE)
      .setDisplayName('Avg Bid');
    
    // Daily Budget (Column O = 15)
    pivotTable.addPivotValue(15, SpreadsheetApp.PivotTableSummarizeFunction.SUM)
      .setDisplayName('Budget');
    
    Logger.log('Native pivot table created with row groups and values');
    
    // Hide the temp source sheet
    tempSheet.hideSheet();
    
    // Wait for pivot to render
    SpreadsheetApp.flush();
    
    // Find where the pivot table ends
    const pivotLastCol = pivotSheet.getLastColumn();
    const pivotLastRow = pivotSheet.getLastRow();
    
    Logger.log(`Pivot table spans columns 1-${pivotLastCol}, rows 1-${pivotLastRow}`);
    
    // Add edit columns starting after a spacer column
    const spacerCol = pivotLastCol + 1;
    const bidIdCol = pivotLastCol + 2;      // J (or wherever pivot ends + 2)
    const campIdCol = pivotLastCol + 3;     // K
    const newCpmCol = pivotLastCol + 4;     // L
    const newBudgetCol = pivotLastCol + 5;  // M
    const commentCol = pivotLastCol + 6;    // N
    
    // Headers for edit columns
    const editHeaders = ['', 'Bid ID', 'Campaign ID', 'New CPM', 'New Budget', 'Comment'];
    pivotSheet.getRange(1, spacerCol, 1, editHeaders.length).setValues([editHeaders]);
    pivotSheet.getRange(1, bidIdCol, 1, 5)
      .setFontWeight('bold')
      .setBackground('#ff9800')
      .setFontColor('white');
    
    // Add lookup formulas for each data row (Bid ID and Campaign ID)
    // These formulas match on Strategy + Sub-Strategy + Campaign Name + Spot Name
    // For subtotal rows, the formula returns empty (no exact match)
    if (pivotLastRow > 1) {
      const formulas = [];
      for (let row = 2; row <= pivotLastRow; row++) {
        // Bid ID lookup - matches all 4 grouping columns to find unique bid
        // A=Strategy, B=Sub-Strategy, C=Campaign, D=Spot in pivot
        // A=Strategy, B=Sub-Strategy, C=Campaign, H=Spot in Control Panel
        const bidFormula = `=IFERROR(INDEX('${CP_SHEET_NAME}'!$AH$3:$AH,MATCH(1,('${CP_SHEET_NAME}'!$A$3:$A=A${row})*('${CP_SHEET_NAME}'!$B$3:$B=B${row})*('${CP_SHEET_NAME}'!$C$3:$C=C${row})*('${CP_SHEET_NAME}'!$H$3:$H=D${row}),0)),"")`;
        
        // Campaign ID lookup - same matching logic
        const campFormula = `=IFERROR(INDEX('${CP_SHEET_NAME}'!$D$3:$D,MATCH(1,('${CP_SHEET_NAME}'!$A$3:$A=A${row})*('${CP_SHEET_NAME}'!$B$3:$B=B${row})*('${CP_SHEET_NAME}'!$C$3:$C=C${row})*('${CP_SHEET_NAME}'!$H$3:$H=D${row}),0)),"")`;
        
        formulas.push([bidFormula, campFormula]);
      }
      
      // Set all formulas at once (batch operation)
      pivotSheet.getRange(2, bidIdCol, pivotLastRow - 1, 2).setFormulas(formulas);
      Logger.log(`Added ${formulas.length} lookup formulas for Bid ID and Campaign ID`);
    }
    
    // Format columns
    // Gray out ID columns (reference only)
    pivotSheet.getRange(2, bidIdCol, pivotLastRow - 1, 2).setFontColor('#999999');
    
    // Yellow background for editable columns
    pivotSheet.getRange(2, newCpmCol, pivotLastRow - 1, 3).setBackground('#fff9c4');
    
    // Number formats
    pivotSheet.getRange(2, newCpmCol, pivotLastRow - 1, 1).setNumberFormat('$#,##0.000');
    pivotSheet.getRange(2, newBudgetCol, pivotLastRow - 1, 1).setNumberFormat('$#,##0.00');
    
    // Auto-resize columns
    pivotSheet.autoResizeColumns(1, commentCol);
    
    // Freeze first row
    pivotSheet.setFrozenRows(1);
    
    // Activate the sheet
    ss.setActiveSheet(pivotSheet);
    
    ui.alert('Success', 
      `Native Pivot Table created!\n\n` +
      `• ${dataRowCount} bids in Control Panel\n` +
      `• Pivot table in columns A-${String.fromCharCode(64 + pivotLastCol)}\n` +
      `• Edit columns in columns ${String.fromCharCode(64 + bidIdCol)}-${String.fromCharCode(64 + commentCol)}\n\n` +
      `Use the native pivot table's +/- buttons to collapse/expand groups.\n\n` +
      `To edit bids:\n` +
      `1. Find the row you want to edit (Bid ID shows for detail rows only)\n` +
      `2. Enter New CPM, New Budget, or Comment\n` +
      `3. Run "Update from Pivot" to apply changes\n\n` +
      `Note: Subtotal rows have empty Bid ID - only detail rows can be edited.`,
      ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    ui.alert('Error', `Failed to build Pivot Table: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

// Store column positions for helper functions (set after pivot is built)
// New V6 order: Checkbox → New CPM → New Budget → Comment → Bid ID → Campaign ID
const PIVOT_EDIT_COLS = {
  CHECKBOX_OFFSET: 2,    // pivotLastCol + 2
  NEW_CPM_OFFSET: 3,     // pivotLastCol + 3
  NEW_BUDGET_OFFSET: 4,  // pivotLastCol + 4
  COMMENT_OFFSET: 5,     // pivotLastCol + 5
  BID_ID_OFFSET: 6,      // pivotLastCol + 6
  CAMP_ID_OFFSET: 7      // pivotLastCol + 7
};

// Helper to find edit column positions in pivot sheet
function getPivotEditColumns(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // Find columns by header name (dynamic detection)
  let checkboxCol = -1, newCpmCol = -1, newBudgetCol = -1, commentCol = -1, bidIdCol = -1, campIdCol = -1;
  
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (h === '📊') checkboxCol = i + 1;
    else if (h === 'New CPM') newCpmCol = i + 1;
    else if (h === 'New Budget') newBudgetCol = i + 1;
    else if (h === 'Comment') commentCol = i + 1;
    else if (h === 'Bid ID') bidIdCol = i + 1;
    else if (h === 'Campaign ID') campIdCol = i + 1;
  }
  
  if (bidIdCol === -1 || newCpmCol === -1) {
    throw new Error('Could not find edit columns. Please run "Add Edit Columns to Pivot" first.');
  }
  
  return {
    checkboxCol: checkboxCol,
    newCpmCol: newCpmCol,
    newBudgetCol: newBudgetCol,
    commentCol: commentCol,
    bidIdCol: bidIdCol,
    campIdCol: campIdCol
  };
}

// Find the 'Avg Bid' column in the pivot table (for copying to New CPM)
function getPivotBidColumn(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === 'Avg Bid') {
      return i + 1;  // 1-indexed
    }
  }
  return -1;
}

// Find the 'Budget' column in the pivot table (for copying to New Budget)
function getPivotBudgetColumn(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === 'Budget') {
      return i + 1;  // 1-indexed
    }
  }
  return -1;
}

/**
 * Copy current bids to New CPM column in Pivot Table
 * V6: Finds columns dynamically based on native pivot structure
 */
function cpCopyBidsPivot() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CP_PIVOT_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Error', `Sheet "${CP_PIVOT_SHEET_NAME}" not found. Please build Pivot Table first.`, ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('Error', 'No data found in Pivot Table.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // Find column positions dynamically
    const avgBidCol = getPivotBidColumn(sheet);
    const editCols = getPivotEditColumns(sheet);
    
    if (avgBidCol === -1) {
      ui.alert('Error', 'Could not find "Avg Bid" column. Please rebuild Pivot Table.', ui.ButtonSet.OK);
      return;
    }
    
    // Copy Avg Bid to New CPM - data starts at row 2 (row 1 is header)
    const currentBids = sheet.getRange(2, avgBidCol, lastRow - 1, 1).getValues();
    sheet.getRange(2, editCols.newCpmCol, lastRow - 1, 1).setValues(currentBids);
    
    ui.alert('Done', `Copied ${lastRow - 1} bid values to "New CPM" column.`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', `Failed to copy bids: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Copy current budgets to New Budget column in Pivot Table
 * V6: Finds columns dynamically based on native pivot structure
 */
function cpCopyBudgetsPivot() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CP_PIVOT_SHEET_NAME);
  
  if (!sheet) {
    ui.alert('Error', `Sheet "${CP_PIVOT_SHEET_NAME}" not found. Please build Pivot Table first.`, ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('Error', 'No data found in Pivot Table.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // Find column positions dynamically
    const budgetCol = getPivotBudgetColumn(sheet);
    const editCols = getPivotEditColumns(sheet);
    
    if (budgetCol === -1) {
      ui.alert('Error', 'Could not find "Budget" column. Please rebuild Pivot Table.', ui.ButtonSet.OK);
      return;
    }
    
    // Copy Budget to New Budget - data starts at row 2 (row 1 is header)
    const currentBudgets = sheet.getRange(2, budgetCol, lastRow - 1, 1).getValues();
    sheet.getRange(2, editCols.newBudgetCol, lastRow - 1, 1).setValues(currentBudgets);
    
    ui.alert('Done', `Copied ${lastRow - 1} budget values to "New Budget" column.`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', `Failed to copy budgets: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Update bids and budgets from Pivot Table
 * V6: Reads Bid ID and Campaign ID from lookup formula columns
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
  if (lastRow < 2) {
    ui.alert('Error', 'No data found.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // Find column positions dynamically
    const editCols = getPivotEditColumns(sheet);
    const avgBidCol = getPivotBidColumn(sheet);
    const budgetCol = getPivotBudgetColumn(sheet);
    
    // Get data from row 2 onwards (row 1 is header)
    // We need: row labels (A-D), avg bid value, budget value, and edit columns
    const lastCol = sheet.getLastColumn();
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    // Collect bid and budget updates
    const bidUpdates = [];
    const budgetUpdates = {};  // campaignId -> { newBudget, comment, rowIndices }
    
    // Column indices are 0-based in the data array
    const bidIdIdx = editCols.bidIdCol - 1;
    const campIdIdx = editCols.campIdCol - 1;
    const newCpmIdx = editCols.newCpmCol - 1;
    const newBudgetIdx = editCols.newBudgetCol - 1;
    const commentIdx = editCols.commentCol - 1;
    const avgBidIdx = avgBidCol - 1;
    const budgetIdx = budgetCol - 1;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const bidId = String(row[bidIdIdx] || '');
      const campaignId = String(row[campIdIdx] || '');
      
      // Skip subtotal rows (empty Bid ID from lookup formula)
      if (!bidId) continue;
      
      const currentBid = cpToNumeric(row[avgBidIdx], 0);
      const newCpm = cpToNumeric(row[newCpmIdx], 0);
      const newBudget = cpToNumeric(row[newBudgetIdx], 0);
      const comment = String(row[commentIdx] || '');
      const dailyBudget = cpToNumeric(row[budgetIdx], 0);
      
      // Row labels from pivot (A=0, B=1, C=2, D=3)
      const campaignName = row[2] || '';  // C - Campaign Name
      const spotName = row[3] || '';      // D - Spot Name
      
      // Check for bid update
      if (newCpm > 0 && newCpm !== currentBid) {
        bidUpdates.push({
          rowIndex: i + 2,  // Sheet row (data starts at row 2)
          bidId: bidId,
          campaignId: campaignId,
          campaignName: campaignName,
          spotId: '',  // Not available in V6 pivot
          spotName: spotName,
          deviceOS: '',  // Not available in V6 pivot
          country: '',   // Not available in V6 pivot
          currentBid: currentBid,
          newBid: newCpm,
          change: ((newCpm - currentBid) / currentBid * 100).toFixed(2),
          comment: comment
        });
      }
      
      // Check for budget update (aggregate by campaign)
      if (newBudget > 0 && newBudget !== dailyBudget) {
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
        
        // Clear edit columns in pivot sheet (V6: use dynamic columns)
        sheet.getRange(bid.rowIndex, editCols.newCpmCol).setValue('');
        sheet.getRange(bid.rowIndex, editCols.commentCol).setValue('');
        
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
        
        // Clear edit columns in pivot sheet rows for this campaign (V6: use dynamic columns)
        for (const rowIndex of budget.rowIndices) {
          sheet.getRange(rowIndex, editCols.newBudgetCol).setValue('');
          sheet.getRange(rowIndex, editCols.commentCol).setValue('');
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
  let resultMsg = `Update Complete!\n\n`;
  
  if (bidUpdates.length > 0) {
    resultMsg += `BIDS: ${bidSuccess} successful, ${bidFail} failed\n`;
  }
  if (budgetCampaigns.length > 0) {
    resultMsg += `BUDGETS: ${budgetSuccess} successful, ${budgetFail} failed\n`;
  }
  
  ui.alert('Update Results', resultMsg, ui.ButtonSet.OK);
  
  } catch (error) {
    ui.alert('Error', `Failed to process updates: ${error.toString()}`, ui.ButtonSet.OK);
  }
}
