#!/usr/bin/env python3
"""
Test to verify if TJ fixed the geo/country bug for multi-country campaigns.

Bug: /api/bids/{campaignId}.json returns same countryCode for all bids
     even when campaign targets multiple countries.
"""

import requests
import json
from datetime import datetime

# API Configuration
API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
BASE_URL = "https://api.trafficjunky.com"

# Test campaign with multiple countries (AU, CA, NZ, UK, US)
TEST_CAMPAIGN_ID = "1013205721"

def test_geo_data():
    print(f"=" * 60)
    print(f"GEO BUG TEST - {datetime.now()}")
    print(f"=" * 60)
    print(f"\nTesting campaign: {TEST_CAMPAIGN_ID}")
    
    url = f"{BASE_URL}/api/bids/{TEST_CAMPAIGN_ID}.json?api_key={API_KEY}"
    
    print(f"\nFetching: {url.replace(API_KEY, 'API_KEY')}")
    
    resp = requests.get(url)
    
    print(f"Response code: {resp.status_code}")
    
    if resp.status_code != 200:
        print(f"ERROR: {resp.text[:500]}")
        return
    
    data = resp.json()
    
    # Collect all unique geos and countries
    all_geos = {}
    all_countries = set()
    bid_count = 0
    
    print(f"\n--- ANALYZING BIDS ---")
    
    for bid_id, bid in data.items():
        bid_count += 1
        geos = bid.get('geos', {})
        
        for geo_id, geo_info in geos.items():
            country_code = geo_info.get('countryCode', 'N/A')
            country_name = geo_info.get('countryName', 'N/A')
            all_countries.add(country_code)
            
            if geo_id not in all_geos:
                all_geos[geo_id] = {
                    'countryCode': country_code,
                    'countryName': country_name,
                    'bid_count': 0
                }
            all_geos[geo_id]['bid_count'] += 1
    
    print(f"\nTotal bids: {bid_count}")
    print(f"Unique geo IDs: {len(all_geos)}")
    print(f"Unique countries: {len(all_countries)}")
    
    print(f"\n--- GEO BREAKDOWN ---")
    for geo_id, info in all_geos.items():
        print(f"  GeoID {geo_id}: {info['countryCode']} ({info['countryName']}) - {info['bid_count']} bids")
    
    print(f"\n--- COUNTRIES FOUND ---")
    print(f"  {sorted(all_countries)}")
    
    # Show sample bids
    print(f"\n--- SAMPLE BIDS (first 5) ---")
    for i, (bid_id, bid) in enumerate(list(data.items())[:5]):
        geos = bid.get('geos', {})
        geo_list = list(geos.values())
        geo_str = geo_list[0]['countryCode'] if geo_list else 'N/A'
        print(f"  Bid {bid_id}: ${bid.get('bid', 'N/A')} - {bid.get('spot_name', 'N/A')} - Country: {geo_str}")
    
    # VERDICT
    print(f"\n{'=' * 60}")
    if len(all_countries) == 1 and bid_count > 15:
        print("❌ BUG STILL EXISTS - All bids show same country!")
        print(f"   Expected: 5 countries (AU, CA, NZ, UK, US)")
        print(f"   Got: {list(all_countries)}")
    elif len(all_countries) >= 5:
        print("✅ BUG APPEARS FIXED - Multiple countries returned!")
        print(f"   Countries: {sorted(all_countries)}")
    else:
        print(f"⚠️ UNCLEAR - Got {len(all_countries)} countries for {bid_count} bids")
    print(f"{'=' * 60}")

if __name__ == "__main__":
    test_geo_data()
