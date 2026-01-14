# V8 - Hybrid Approach: Two Endpoints, Complete Data

## 🎯 The Strategy

Use BOTH endpoints to get complete data:

### Step 1: Get All Campaign Stats (Pagination Works!)
**Endpoint**: `/api/campaigns/stats.json`  
**Pagination**: `offset=501, 1001, 1501...` ✅ Works!  
**Data**: impressions, clicks, conversions, ctr, ecpm, ecpc, ads_count, ads_paused

### Step 2: Get Missing Fields (Single Call)
**Endpoint**: `/api/campaigns/bids/stats.json`  
**Pagination**: Just `limit=500&offset=1` (first 500 is enough to match)  
**Data**: status, dailyBudget, dailyBudgetLeft, numberOfBids, CPM

### Step 3: Merge by Campaign ID
Match campaigns from both endpoints and create complete rows

## 📊 Complete Data Mapping

| Field | Source Endpoint | API Field |
|-------|----------------|-----------|
| Campaign ID | stats | `campaign_id` |
| Campaign Name | stats | `campaign_name` |
| Campaign Type | stats | `campaign_type` |
| **Status** | **bids/stats** | `status` |
| **Daily Budget** | **bids/stats** | `dailyBudget` |
| **Daily Budget Left** | **bids/stats** | `dailyBudgetLeft` |
| Ads Paused | stats | `ads_paused` |
| **Number of Bids** | **bids/stats** | `numberOfBids` |
| **Number of Creatives** | **bids/stats** | `numberOfCreative` |
| Impressions | stats | `impressions` |
| Clicks | stats | `clicks` |
| Conversions | stats | `conversions` |
| Cost | stats | (calculate from ecpc * clicks or use cost if available) |
| CTR | stats | `ctr` |
| **CPM** | **bids/stats** | `CPM` |
| ECPM | stats | `ecpm` (alternative metric) |
| ECPC | stats | `ecpc` |

## 🔧 Implementation Plan

```javascript
// 1. Fetch all campaigns with stats (good pagination)
const allCampaigns = fetchAllCampaignStats(startDate, endDate);
// Result: { campaign_id_123: {stats}, campaign_id_456: {stats}, ... }

// 2. Fetch additional fields from bids endpoint (single call)
const additionalFields = fetchAdditionalFields(startDate, endDate);
// Result: { campaign_id_123: {status, budget, etc}, ... }

// 3. Merge data
for (let campaignId in allCampaigns) {
  const stats = allCampaigns[campaignId];
  const additional = additionalFields[campaignId] || {};
  
  const row = [
    campaignId,
    stats.campaign_name,
    stats.campaign_type,
    additional.status || 'unknown',
    additional.dailyBudget || 0,
    additional.dailyBudgetLeft || 0,
    stats.ads_paused || 0,
    additional.numberOfBids || 0,
    additional.numberOfCreative || 0,
    stats.impressions || 0,
    stats.clicks || 0,
    stats.conversions || 0,
    stats.ecpc * stats.clicks || 0, // Calculate cost
    stats.ctr || 0,
    additional.CPM || stats.ecpm || 0,
    new Date(),
    dateRange
  ];
}

// 4. Write merged data
writeToSheet(rows);
```

## ✅ Benefits

1. **Complete Data** - All fields present
2. **Fast Pagination** - stats endpoint works with offset=501
3. **No Timeout** - Only fetches what's needed
4. **Accurate** - Gets real values for status, budgets, CPM

## ⚠️ Limitations

- If you have > 500 campaigns, some won't have status/budget data
- Solution: Could also paginate the bids endpoint slowly with offset=1,2,3... but probably not worth it for just a few fields

---

Ready to implement in V8!

