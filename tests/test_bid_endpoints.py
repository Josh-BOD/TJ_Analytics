"""
Test TrafficJunky Bid API Endpoints
Run this to see what data each endpoint returns
"""

import requests
import json
from datetime import datetime, timedelta

# Configuration
API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
BASE_URL = "https://api.trafficjunky.com/api"
CAMPAIGN_ID = "1013232471"

def test_endpoint(name, url, params=None):
    """Test an endpoint and print the response"""
    print(f"\n{'='*80}")
    print(f"TESTING: {name}")
    print(f"{'='*80}")
    print(f"URL: {url.replace(API_KEY, 'HIDDEN')}")
    
    try:
        if params:
            response = requests.get(url, params=params, timeout=30)
        else:
            response = requests.get(url, timeout=30)
        
        print(f"Status: {response.status_code}")
        
        if response.status_code == 200:
            try:
                data = response.json()
                
                if isinstance(data, list):
                    print(f"Response: Array with {len(data)} items")
                    if len(data) > 0:
                        print(f"\nFirst item fields: {list(data[0].keys())}")
                        print(f"\nFirst item:")
                        print(json.dumps(data[0], indent=2))
                        
                        if len(data) > 1:
                            print(f"\n... and {len(data) - 1} more items")
                            
                        # Show all items if small array
                        if len(data) <= 10:
                            print(f"\n--- ALL ITEMS ---")
                            for i, item in enumerate(data):
                                print(f"\nItem {i+1}:")
                                print(json.dumps(item, indent=2))
                                
                elif isinstance(data, dict):
                    print(f"Response: Object with {len(data)} keys")
                    print(f"Keys: {list(data.keys())}")
                    print(f"\nFull response:")
                    print(json.dumps(data, indent=2)[:5000])
                else:
                    print(f"Response type: {type(data)}")
                    print(data)
                    
                return data
                
            except json.JSONDecodeError:
                print(f"Response (not JSON): {response.text[:500]}")
        else:
            print(f"Error: {response.text[:500]}")
            
    except Exception as e:
        print(f"Exception: {str(e)}")
    
    return None


def main():
    print("="*80)
    print("TRAFFICJUNKY BID API ENDPOINT TESTER")
    print("="*80)
    print(f"Campaign ID: {CAMPAIGN_ID}")
    
    # Calculate date range for endpoints that need it
    end_date = datetime.now() - timedelta(days=1)
    start_date = end_date - timedelta(days=30)
    date_params = {
        'api_key': API_KEY,
        'startDate': start_date.strftime('%d/%m/%Y'),
        'endDate': end_date.strftime('%d/%m/%Y'),
        'limit': 500,
        'offset': 1
    }
    
    # Test 1: /api/bids/{campaignId}.json - Full bid details
    print("\n" + "#"*80)
    print("# ENDPOINT 1: /api/bids/{campaignId}.json (Full bid details)")
    print("#"*80)
    data1 = test_endpoint(
        "/api/bids/{campaignId}.json",
        f"{BASE_URL}/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}"
    )
    
    # Test 2: /api/bids/{campaignId}/active.json - Active bids only
    print("\n" + "#"*80)
    print("# ENDPOINT 2: /api/bids/{campaignId}/active.json (Active bids)")
    print("#"*80)
    data2 = test_endpoint(
        "/api/bids/{campaignId}/active.json",
        f"{BASE_URL}/bids/{CAMPAIGN_ID}/active.json?api_key={API_KEY}"
    )
    
    # Test 3: /api/campaigns/bids/stats.json - Campaign stats with bids
    print("\n" + "#"*80)
    print("# ENDPOINT 3: /api/campaigns/bids/stats.json (Campaign stats)")
    print("#"*80)
    data3 = test_endpoint(
        "/api/campaigns/bids/stats.json",
        f"{BASE_URL}/campaigns/bids/stats.json",
        params=date_params
    )
    
    # Check if our campaign is in the response
    if data3:
        if isinstance(data3, list):
            found = [c for c in data3 if str(c.get('campaignId')) == CAMPAIGN_ID]
            if found:
                print(f"\n✅ Found campaign {CAMPAIGN_ID} in response!")
                print(json.dumps(found[0], indent=2))
            else:
                print(f"\n❌ Campaign {CAMPAIGN_ID} NOT in response")
                print(f"   Available campaign IDs: {[c.get('campaignId') for c in data3[:10]]}")
        elif isinstance(data3, dict):
            if CAMPAIGN_ID in data3:
                print(f"\n✅ Found campaign {CAMPAIGN_ID} in response!")
                print(json.dumps(data3[CAMPAIGN_ID], indent=2))
            else:
                print(f"\n❌ Campaign {CAMPAIGN_ID} NOT in response")
                print(f"   Available campaign IDs: {list(data3.keys())[:10]}")
    
    # Test 4: /api/campaigns/{campaignId}.json - Campaign details
    print("\n" + "#"*80)
    print("# ENDPOINT 4: /api/campaigns/{campaignId}.json (Campaign details)")
    print("#"*80)
    data4 = test_endpoint(
        "/api/campaigns/{campaignId}.json",
        f"{BASE_URL}/campaigns/{CAMPAIGN_ID}.json?api_key={API_KEY}"
    )
    
    # Test 5: Get bid details by bid_id (if we have one)
    if data2 and isinstance(data2, list) and len(data2) > 0:
        bid_id = data2[0].get('bid_id')
        if bid_id:
            print("\n" + "#"*80)
            print(f"# ENDPOINT 5: /api/bids/{bid_id}.json (Single bid details)")
            print("#"*80)
            data5 = test_endpoint(
                f"/api/bids/{bid_id}.json",
                f"{BASE_URL}/bids/{bid_id}.json?api_key={API_KEY}"
            )
    
    # Summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    
    print("\nEndpoint 1 (/api/bids/{campaignId}.json):")
    if data1:
        if isinstance(data1, list) and len(data1) > 0:
            print(f"  ✅ Returns {len(data1)} bids")
            print(f"  Fields: {list(data1[0].keys())}")
        elif isinstance(data1, dict):
            print(f"  ✅ Returns object with {len(data1)} keys")
    else:
        print("  ❌ Failed or empty")
    
    print("\nEndpoint 2 (/api/bids/{campaignId}/active.json):")
    if data2:
        if isinstance(data2, list) and len(data2) > 0:
            print(f"  ✅ Returns {len(data2)} active bids")
            print(f"  Fields: {list(data2[0].keys())}")
    else:
        print("  ❌ Failed or empty")
    
    print("\nEndpoint 4 (/api/campaigns/{campaignId}.json):")
    if data4:
        if isinstance(data4, dict):
            print(f"  ✅ Returns campaign details with {len(data4)} fields")
            print(f"  Fields: {list(data4.keys())}")
            if 'spots' in data4:
                print(f"  Spots: {data4['spots']}")
            if 'bids' in data4:
                print(f"  Has bids array: {len(data4['bids'])} bids")
    else:
        print("  ❌ Failed or empty")


if __name__ == "__main__":
    main()
