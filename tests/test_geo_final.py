"""
Final test for geo stats - check all endpoints carefully
"""

import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013232471"

from datetime import datetime, timedelta
yesterday = datetime.now() - timedelta(days=1)
date_str = yesterday.strftime('%d/%m/%Y')

print("=" * 80)
print("FINAL GEO STATS TEST")
print("=" * 80)

# 1. Direct /api/bids/{campaignId}.json endpoint
print("\n1. DIRECT BIDS ENDPOINT (/api/bids/{campaignId}.json):")
print("-" * 80)

url1 = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}"
response1 = requests.get(url1)
data1 = response1.json()

print(f"Status: {response1.status_code}")
print(f"Response type: {type(data1)}")

if isinstance(data1, dict):
    print(f"Number of bids: {len(data1)}")
    
    for bid_id, bid_data in list(data1.items())[:2]:
        print(f"\n  Bid ID: {bid_id}")
        if isinstance(bid_data, dict):
            print(f"  Keys: {list(bid_data.keys())}")
            
            # Check geos
            geos = bid_data.get('geos', {})
            print(f"  Geos count: {len(geos)}")
            if geos:
                for geo_id, geo_data in list(geos.items())[:1]:
                    print(f"\n    Geo {geo_id}:")
                    if isinstance(geo_data, dict):
                        for k, v in geo_data.items():
                            print(f"      {k}: {v}")

# 2. Campaign bids stats endpoint 
print("\n\n2. CAMPAIGNS BIDS STATS ENDPOINT:")
print("-" * 80)

url2 = f"https://api.trafficjunky.com/api/campaigns/bids/stats.json?api_key={API_KEY}&startDate={date_str}&endDate={date_str}&limit=500&offset=1"
response2 = requests.get(url2)
data2 = response2.json()

# Find our campaign
if isinstance(data2, dict):
    our_camp = data2.get(CAMPAIGN_ID)
    if our_camp:
        print(f"Found campaign: {our_camp.get('campaignName')}")
        print(f"Keys: {list(our_camp.keys())}")
        
        bids = our_camp.get('bids', [])
        print(f"\nBids count: {len(bids)}")
        
        if bids:
            first_bid = bids[0] if isinstance(bids, list) else list(bids.values())[0]
            print(f"First bid structure:")
            print(json.dumps(first_bid, indent=2))
    else:
        # Show what campaigns are available
        print("Campaign not found. Available campaigns:")
        for cid, cdata in list(data2.items())[:3]:
            if isinstance(cdata, dict):
                print(f"  {cid}: {cdata.get('campaignName', 'N/A')}")
                
                # Show bids structure for first campaign
                if cid == list(data2.keys())[0]:
                    bids = cdata.get('bids', [])
                    if bids:
                        first_bid = bids[0] if isinstance(bids, list) else list(bids.values())[0]
                        print(f"\n  First bid for this campaign:")
                        print(json.dumps(first_bid, indent=2))

# 3. Check if there's stats at geo level in the bids/stats response
print("\n\n3. CHECKING FOR GEO-LEVEL STATS IN BIDS/STATS RESPONSE:")
print("-" * 80)

# Get all campaigns and look for any with geo stats
for camp_id, camp_data in list(data2.items())[:5]:
    if isinstance(camp_data, dict):
        bids = camp_data.get('bids', [])
        
        for bid in bids[:2]:
            if isinstance(bid, dict):
                # Check if this bid has impressions/clicks fields
                if 'impressions' in bid or 'clicks' in bid or 'stats' in bid:
                    print(f"Campaign {camp_id} bid has stats!")
                    print(json.dumps(bid, indent=2)[:500])
                    break

# 4. Full raw output of first campaign's bids
print("\n\n4. RAW BIDS DATA FROM FIRST AVAILABLE CAMPAIGN:")
print("-" * 80)

first_camp_id = list(data2.keys())[0]
first_camp_data = data2[first_camp_id]
print(f"Campaign: {first_camp_data.get('campaignName')}")
print(f"\nFull bids array:")
print(json.dumps(first_camp_data.get('bids', []), indent=2)[:2000])

print("\n" + "=" * 80)
print("SUMMARY")
print("=" * 80)
print("""
Two different endpoints return different bid structures:

1. /api/bids/{campaignId}.json
   - Returns bids with: geos object (geoId, countryCode, countryName)
   - Stats are at BID level, NOT geo level

2. /api/campaigns/bids/stats.json  
   - Returns bids array with: placementId, bid, countryCode, countryName, etc.
   - No stats per geo/country

CONCLUSION: Geo-level stats (impressions/clicks/cost per country) still NOT available.
""")
