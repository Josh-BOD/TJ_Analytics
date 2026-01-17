"""
Test the groupBy=geo parameter to see if geo-level stats are now available
"""

import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013232471"

from datetime import datetime, timedelta
yesterday = datetime.now() - timedelta(days=1)
date_str = yesterday.strftime('%d/%m/%Y')

print("=" * 80)
print("TESTING groupBy=geo PARAMETER")
print("=" * 80)

# 1. Without groupBy
print("\n1. WITHOUT groupBy (baseline):")
print("-" * 80)
url1 = f"https://api.trafficjunky.com/api/campaigns/bids/stats.json?api_key={API_KEY}&startDate={date_str}&endDate={date_str}&limit=1&offset=1"
response1 = requests.get(url1)
data1 = response1.json()

if data1:
    first_camp = list(data1.values())[0]
    print(f"Campaign: {first_camp.get('campaignName')}")
    print(f"Top-level keys: {list(first_camp.keys())}")
    
    # Check bids structure
    if 'bids' in first_camp:
        bids = first_camp['bids']
        print(f"\nBids count: {len(bids)}")
        if bids:
            first_bid = bids[0] if isinstance(bids, list) else list(bids.values())[0]
            print(f"Bid keys: {list(first_bid.keys()) if isinstance(first_bid, dict) else 'N/A'}")
            
            # Check geos in bid
            geos = first_bid.get('geos', {})
            if geos:
                first_geo = list(geos.values())[0] if isinstance(geos, dict) else geos[0] if geos else {}
                print(f"Geo keys: {list(first_geo.keys()) if isinstance(first_geo, dict) else 'N/A'}")

# 2. With groupBy=geo
print("\n\n2. WITH groupBy=geo:")
print("-" * 80)
url2 = f"https://api.trafficjunky.com/api/campaigns/bids/stats.json?api_key={API_KEY}&startDate={date_str}&endDate={date_str}&limit=1&offset=1&groupBy=geo"
response2 = requests.get(url2)
data2 = response2.json()

if data2:
    first_camp2 = list(data2.values())[0]
    print(f"Campaign: {first_camp2.get('campaignName')}")
    print(f"Top-level keys: {list(first_camp2.keys())}")
    
    # Check if there's a new geoStats or similar field
    for key in first_camp2.keys():
        if 'geo' in key.lower():
            print(f"\nFound geo field: {key}")
            print(f"Value: {json.dumps(first_camp2[key], indent=2)[:500]}")
    
    # Check bids structure
    if 'bids' in first_camp2:
        bids = first_camp2['bids']
        print(f"\nBids count: {len(bids)}")
        if bids:
            first_bid = bids[0] if isinstance(bids, list) else list(bids.values())[0]
            print(f"Bid keys: {list(first_bid.keys()) if isinstance(first_bid, dict) else 'N/A'}")
            
            # Check geos in bid - look for stats
            geos = first_bid.get('geos', {})
            if geos:
                print(f"\nGeos count: {len(geos)}")
                first_geo = list(geos.values())[0] if isinstance(geos, dict) else geos[0] if geos else {}
                print(f"Geo keys: {list(first_geo.keys()) if isinstance(first_geo, dict) else 'N/A'}")
                print(f"\nFull geo data:")
                print(json.dumps(first_geo, indent=2))

# 3. Direct bid endpoint with today's date
print("\n\n3. DIRECT BID ENDPOINT WITH TODAY'S DATE:")
print("-" * 80)
today = datetime.now()
today_str = today.strftime('%d/%m/%Y')

url3 = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}&startDate={today_str}&endDate={today_str}"
response3 = requests.get(url3)
data3 = response3.json()

if data3:
    first_bid3 = list(data3.values())[0]
    print(f"Bid ID: {first_bid3.get('bid_id')}")
    print(f"Spot: {first_bid3.get('spot_name')}")
    
    geos3 = first_bid3.get('geos', {})
    if geos3:
        print(f"\nGeos count: {len(geos3)}")
        for geo_id, geo_data in list(geos3.items())[:3]:
            print(f"\n  Geo ID {geo_id}:")
            if isinstance(geo_data, dict):
                for k, v in geo_data.items():
                    print(f"    {k}: {v}")

# 4. Check full response structure
print("\n\n4. FULL RESPONSE WITH groupBy=geo (first campaign):")
print("-" * 80)
print(json.dumps(first_camp2, indent=2)[:3000])

print("\n" + "=" * 80)
print("TEST COMPLETE")
print("=" * 80)
