#!/usr/bin/env python3
"""
Simple TrafficJunky API Test
Tests multiple parameter combinations to find what works
"""

import requests
import json
from datetime import datetime, timedelta

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
BASE_URL = "https://api.trafficjunky.com/api/campaigns/bids/stats.json"

print("=" * 80)
print("TRAFFICJUNKY API DIAGNOSTIC TEST")
print("=" * 80)

# Calculate dates - FIXED last 7 days
end_date = datetime.now() - timedelta(days=1)  # Yesterday
start_date = end_date - timedelta(days=6)  # 7 days total

start_str = start_date.strftime('%d/%m/%Y')
end_str = end_date.strftime('%d/%m/%Y')

print(f"\nUsing FIXED date range (last 7 days):")
print(f"  Start: {start_str} ({start_date.strftime('%Y-%m-%d')})")
print(f"  End:   {end_str} ({end_date.strftime('%Y-%m-%d')})")
print(f"API URL: {BASE_URL}")
print(f"API Key: {API_KEY[:20]}...{API_KEY[-20:]}\n")

# Test configurations
tests = [
    {
        "name": "Test 1: With offset=1 (original)",
        "params": {
            'api_key': API_KEY,
            'startDate': start_str,
            'endDate': end_str,
            'limit': 1000,
            'offset': 1
        }
    },
    {
        "name": "Test 2: Without offset",
        "params": {
            'api_key': API_KEY,
            'startDate': start_str,
            'endDate': end_str,
            'limit': 1000
        }
    },
    {
        "name": "Test 3: Without limit or offset",
        "params": {
            'api_key': API_KEY,
            'startDate': start_str,
            'endDate': end_str
        }
    },
    {
        "name": "Test 4: Only API key (no dates)",
        "params": {
            'api_key': API_KEY
        }
    },
    {
        "name": "Test 5: With offset=0",
        "params": {
            'api_key': API_KEY,
            'startDate': start_str,
            'endDate': end_str,
            'limit': 1000,
            'offset': 0
        }
    }
]

for test in tests:
    print("\n" + "=" * 80)
    print(test["name"])
    print("=" * 80)
    print(f"Parameters: {json.dumps({k: v if k != 'api_key' else 'HIDDEN' for k, v in test['params'].items()}, indent=2)}")
    
    try:
        response = requests.get(BASE_URL, params=test["params"], timeout=30)
        
        print(f"\n✓ Status Code: {response.status_code}")
        
        if response.status_code == 200:
            try:
                data = response.json()
                
                # Check for error message
                if isinstance(data, dict) and 'message' in data:
                    print(f"❌ API ERROR: {data['message']}")
                    print(f"Full response: {json.dumps(data, indent=2)}")
                elif isinstance(data, dict) and 'error' in data:
                    print(f"❌ API ERROR: {json.dumps(data['error'], indent=2)}")
                else:
                    # Success!
                    if isinstance(data, dict):
                        campaign_count = len(data)
                        print(f"✅ SUCCESS! Got {campaign_count} campaigns")
                        
                        if campaign_count > 0:
                            first_key = list(data.keys())[0]
                            first_campaign = data[first_key]
                            
                            print(f"\nFirst Campaign ID: {first_key}")
                            print(f"First Campaign Type: {type(first_campaign)}")
                            
                            if isinstance(first_campaign, dict):
                                print(f"First Campaign Fields: {', '.join(list(first_campaign.keys())[:10])}")
                                print(f"Campaign Name: {first_campaign.get('campaignName', 'N/A')}")
                                print(f"Impressions: {first_campaign.get('impressions', 'N/A')}")
                                print(f"Clicks: {first_campaign.get('clicks', 'N/A')}")
                                print(f"Cost: {first_campaign.get('cost', 'N/A')}")
                                
                                print(f"\n✅✅✅ THIS CONFIGURATION WORKS! ✅✅✅")
                            else:
                                print(f"⚠️ First campaign is not a dict: {first_campaign}")
                        else:
                            print("⚠️ Empty response (no campaigns in date range)")
                    
                    elif isinstance(data, list):
                        print(f"✅ SUCCESS! Got {len(data)} campaigns (array format)")
                        if len(data) > 0:
                            print(f"First Campaign Fields: {', '.join(list(data[0].keys())[:10])}")
                            print(f"\n✅✅✅ THIS CONFIGURATION WORKS! ✅✅✅")
                    
            except json.JSONDecodeError as e:
                print(f"❌ Failed to parse JSON: {e}")
                print(f"Raw response: {response.text[:500]}")
        else:
            print(f"❌ HTTP Error: {response.status_code}")
            print(f"Response: {response.text[:500]}")
    
    except requests.exceptions.Timeout:
        print("❌ Request timed out")
    except requests.exceptions.RequestException as e:
        print(f"❌ Request failed: {e}")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")

print("\n" + "=" * 80)
print("TEST COMPLETE")
print("=" * 80)
print("\nLook for '✅✅✅ THIS CONFIGURATION WORKS! ✅✅✅' above to see which test succeeded")
print("If ALL tests show errors, there may be an issue with:")
print("  - API Key (expired or invalid)")
print("  - TrafficJunky account status")
print("  - API endpoint changed/deprecated")
print("  - Temporary API outage")

