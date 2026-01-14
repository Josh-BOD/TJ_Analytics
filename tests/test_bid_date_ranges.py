"""
Test if TrafficJunky API supports date-filtered bid-level stats.

We'll test multiple endpoints and parameter combinations to see if
we can get bid-level stats for specific date ranges.
"""

import requests
from datetime import datetime, timedelta
import pytz
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
API_BASE_URL = "https://api.trafficjunky.com/api"
CAMPAIGN_ID = "1013232471"

def format_date(date):
    """Format date as DD/MM/YYYY"""
    return date.strftime("%d/%m/%Y")

def get_dates():
    """Get test date ranges"""
    est = pytz.timezone("America/New_York")
    now = datetime.now(est)
    today = now.date()
    
    return {
        'today': format_date(today),
        'yesterday': format_date(today - timedelta(days=1)),
        '3_days_ago': format_date(today - timedelta(days=3)),
        '7_days_ago': format_date(today - timedelta(days=7)),
        '30_days_ago': format_date(today - timedelta(days=30)),
    }

def test_endpoint(name, url, params=None):
    """Test an endpoint and return response info"""
    print(f"\n{'='*60}")
    print(f"TEST: {name}")
    print(f"{'='*60}")
    
    full_url = f"{url}?api_key={API_KEY}"
    if params:
        param_str = "&".join([f"{k}={v}" for k, v in params.items()])
        full_url += f"&{param_str}"
        print(f"Params: {params}")
    
    print(f"URL: {url}")
    
    try:
        resp = requests.get(full_url)
        print(f"Status: {resp.status_code}")
        
        if resp.status_code != 200:
            print(f"Error: {resp.text[:200]}")
            return None
        
        data = resp.json()
        return data
        
    except Exception as e:
        print(f"Exception: {e}")
        return None

def extract_bid_stats(data, bid_id):
    """Extract stats for a specific bid from various response formats"""
    if isinstance(data, dict):
        # Direct bid data
        if bid_id in data:
            bid = data[bid_id]
            if 'stats' in bid:
                return bid['stats']
            return bid
        
        # Nested in campaign
        for camp_id, camp_data in data.items():
            if isinstance(camp_data, dict):
                if 'bids' in camp_data:
                    for bid in camp_data.get('bids', []):
                        if str(bid.get('bid_id', '')) == bid_id:
                            return bid
    
    return None

def main():
    dates = get_dates()
    print("="*60)
    print("TESTING BID-LEVEL DATE FILTERING")
    print("="*60)
    print(f"\nCampaign ID: {CAMPAIGN_ID}")
    print(f"Date ranges to test:")
    for name, date in dates.items():
        print(f"  {name}: {date}")
    
    # First, get the bid IDs from a baseline call
    print("\n" + "="*60)
    print("STEP 1: Get Baseline Bid Data (no date params)")
    print("="*60)
    
    baseline_url = f"{API_BASE_URL}/bids/{CAMPAIGN_ID}.json"
    baseline_data = test_endpoint("Baseline /api/bids/{campaignId}.json", baseline_url)
    
    if not baseline_data:
        print("Failed to get baseline data")
        return
    
    bid_ids = list(baseline_data.keys())
    test_bid_id = bid_ids[0] if bid_ids else None
    
    print(f"\nFound {len(bid_ids)} bids")
    print(f"Using bid_id {test_bid_id} for comparison")
    
    if test_bid_id:
        baseline_stats = baseline_data[test_bid_id].get('stats', {})
        print(f"\nBaseline stats for bid {test_bid_id}:")
        print(f"  impressions: {baseline_stats.get('impressions')}")
        print(f"  clicks: {baseline_stats.get('clicks')}")
        print(f"  revenue: {baseline_stats.get('revenue')}")
        print(f"  ecpm: {baseline_stats.get('ecpm')}")
        print(f"  ctr: {baseline_stats.get('ctr')}")
    
    # Test 1: /api/bids/{campaignId}.json with date params
    print("\n" + "="*60)
    print("TEST 1: /api/bids/{campaignId}.json WITH date params")
    print("="*60)
    
    results = {}
    
    for date_name, date_str in dates.items():
        data = test_endpoint(
            f"Bids for {date_name}",
            baseline_url,
            {'startDate': date_str, 'endDate': date_str}
        )
        
        if data and test_bid_id:
            stats = data.get(test_bid_id, {}).get('stats', {})
            results[date_name] = {
                'impressions': stats.get('impressions'),
                'revenue': stats.get('revenue'),
                'ecpm': stats.get('ecpm'),
            }
            print(f"  → impressions={stats.get('impressions')}, revenue={stats.get('revenue')}, ecpm={stats.get('ecpm')}")
    
    # Compare results
    print("\n" + "="*60)
    print("COMPARISON: Do stats differ by date?")
    print("="*60)
    
    unique_impressions = set(r.get('impressions') for r in results.values() if r)
    unique_revenue = set(r.get('revenue') for r in results.values() if r)
    
    print(f"\nUnique impression values: {unique_impressions}")
    print(f"Unique revenue values: {unique_revenue}")
    
    if len(unique_impressions) > 1 or len(unique_revenue) > 1:
        print("\n✅ STATS DIFFER BY DATE - Date filtering IS working!")
    else:
        print("\n❌ STATS ARE IDENTICAL - Date filtering NOT working on this endpoint")
    
    # Test 2: /api/campaigns/bids/stats.json (the endpoint we tried before)
    print("\n" + "="*60)
    print("TEST 2: /api/campaigns/bids/stats.json")
    print("="*60)
    
    bids_stats_url = f"{API_BASE_URL}/campaigns/bids/stats.json"
    
    for date_name in ['today', 'yesterday']:
        data = test_endpoint(
            f"Campaigns bids stats - {date_name}",
            bids_stats_url,
            {'startDate': dates[date_name], 'endDate': dates[date_name], 'limit': 500, 'offset': 1}
        )
        
        if data:
            # Check if our campaign is in there
            if CAMPAIGN_ID in data:
                camp_data = data[CAMPAIGN_ID]
                print(f"\n  Campaign {CAMPAIGN_ID} found!")
                if isinstance(camp_data, dict):
                    print(f"  Keys: {list(camp_data.keys())}")
                    bids = camp_data.get('bids', [])
                    print(f"  Number of bids: {len(bids)}")
                    if bids:
                        print(f"  First bid keys: {list(bids[0].keys())}")
                        # Check if it has stats
                        first_bid = bids[0]
                        print(f"  First bid data: {json.dumps(first_bid, indent=2)[:500]}")
    
    # Test 3: Try other potential endpoints
    print("\n" + "="*60)
    print("TEST 3: Exploring other endpoints")
    print("="*60)
    
    other_endpoints = [
        f"/bids/{CAMPAIGN_ID}/stats.json",
        f"/campaigns/{CAMPAIGN_ID}/bids/stats.json",
        f"/campaigns/{CAMPAIGN_ID}/bids.json",
        f"/bid/{test_bid_id}.json",
        f"/bid/{test_bid_id}/stats.json",
        f"/bids/{test_bid_id}/stats.json",
    ]
    
    for endpoint in other_endpoints:
        url = f"{API_BASE_URL}{endpoint}"
        data = test_endpoint(endpoint, url, {'startDate': dates['yesterday'], 'endDate': dates['yesterday']})
        
        if data:
            print(f"  ✅ Endpoint exists!")
            if isinstance(data, dict):
                print(f"  Keys: {list(data.keys())[:10]}")
            elif isinstance(data, list):
                print(f"  Array with {len(data)} items")
    
    # Test 4: Check if the stats object has date info
    print("\n" + "="*60)
    print("TEST 4: Full bid object inspection")
    print("="*60)
    
    if baseline_data and test_bid_id:
        bid = baseline_data[test_bid_id]
        print(f"\nFull bid object keys: {list(bid.keys())}")
        print(f"\nFull bid object:")
        print(json.dumps(bid, indent=2, default=str))
    
    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)


if __name__ == "__main__":
    main()
