/**
 * DIAGNOSTIC FUNCTION - Add this to your V5 script temporarily
 * This will show you exactly what the API is returning
 */
function debugAPIResponse() {
  const ui = SpreadsheetApp.getUi();
  
  // Use last 7 days to ensure we have data
  const endDate = getESTYesterday();
  const startDate = new Date(endDate);
  startDate.setDate(startDate.getDate() - 6);
  
  const formattedStartDate = formatDate(startDate);
  const formattedEndDate = formatDate(endDate);
  
  Logger.log(`Fetching data from ${formattedStartDate} to ${formattedEndDate}`);
  
  const url = `${API_URL}?api_key=${API_KEY}&startDate=${formattedStartDate}&endDate=${formattedEndDate}&limit=1000&offset=1`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      'method': 'get',
      'contentType': 'application/json',
      'muteHttpExceptions': true
    });
    
    const responseCode = response.getResponseCode();
    Logger.log(`Response Code: ${responseCode}`);
    
    if (responseCode !== 200) {
      Logger.log(`ERROR: ${response.getContentText()}`);
      ui.alert('API Error', `Status ${responseCode}: ${response.getContentText()}`, ui.ButtonSet.OK);
      return;
    }
    
    const jsonData = JSON.parse(response.getContentText());
    
    Logger.log('=== FULL API RESPONSE ===');
    Logger.log(JSON.stringify(jsonData, null, 2));
    Logger.log('======================');
    
    // Analyze structure
    Logger.log('\n=== ANALYSIS ===');
    Logger.log(`Response type: ${typeof jsonData}`);
    Logger.log(`Is array: ${Array.isArray(jsonData)}`);
    
    if (Array.isArray(jsonData)) {
      Logger.log(`Array length: ${jsonData.length}`);
      if (jsonData.length > 0) {
        Logger.log(`First item type: ${typeof jsonData[0]}`);
        Logger.log(`First item keys: ${Object.keys(jsonData[0]).join(', ')}`);
        Logger.log(`First item: ${JSON.stringify(jsonData[0], null, 2)}`);
      }
    } else if (typeof jsonData === 'object') {
      const keys = Object.keys(jsonData);
      Logger.log(`Object keys: ${keys.join(', ')}`);
      Logger.log(`Number of keys: ${keys.length}`);
      
      if (keys.length > 0) {
        const firstKey = keys[0];
        Logger.log(`\nFirst campaign (key: ${firstKey}):`);
        Logger.log(JSON.stringify(jsonData[firstKey], null, 2));
        
        const firstCampaign = jsonData[firstKey];
        Logger.log(`\nFirst campaign type: ${typeof firstCampaign}`);
        
        if (firstCampaign && typeof firstCampaign === 'object') {
          Logger.log(`Campaign fields: ${Object.keys(firstCampaign).join(', ')}`);
          Logger.log(`\nField values:`);
          Logger.log(`  campaignId: ${firstCampaign.campaignId}`);
          Logger.log(`  id: ${firstCampaign.id}`);
          Logger.log(`  campaignName: ${firstCampaign.campaignName}`);
          Logger.log(`  impressions: ${firstCampaign.impressions}`);
          Logger.log(`  clicks: ${firstCampaign.clicks}`);
          Logger.log(`  cost: ${firstCampaign.cost}`);
        }
      }
    }
    
    // Test the conversion logic
    Logger.log('\n=== TESTING CONVERSION LOGIC ===');
    let campaigns = [];
    if (Array.isArray(jsonData)) {
      campaigns = jsonData;
    } else if (typeof jsonData === 'object' && jsonData !== null) {
      campaigns = Object.values(jsonData);
    }
    
    Logger.log(`Campaigns array length: ${campaigns.length}`);
    
    let validCount = 0;
    for (let i = 0; i < campaigns.length; i++) {
      const campaign = campaigns[i];
      Logger.log(`\nCampaign ${i}:`);
      Logger.log(`  Is object: ${campaign && typeof campaign === 'object'}`);
      Logger.log(`  Type: ${typeof campaign}`);
      Logger.log(`  Null: ${campaign === null}`);
      Logger.log(`  Undefined: ${campaign === undefined}`);
      
      if (campaign && typeof campaign === 'object') {
        validCount++;
        Logger.log(`  ✓ Would be processed`);
        Logger.log(`  Campaign ID: ${campaign.campaignId || campaign.id || 'MISSING'}`);
      } else {
        Logger.log(`  ✗ Would be SKIPPED`);
      }
    }
    
    Logger.log(`\n=== RESULT ===`);
    Logger.log(`Valid campaigns to process: ${validCount}`);
    
    ui.alert('Debug Complete', `Check View > Logs for full details.\n\nValid campaigns: ${validCount}`, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`EXCEPTION: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    ui.alert('Error', error.toString(), ui.ButtonSet.OK);
  }
}

