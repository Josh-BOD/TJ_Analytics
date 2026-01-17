"""
Test script to check:
1. All available fields on bid objects (looking for min CPM)
2. Geo breakdown on bids (TJ said they fixed it)
"""

import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013232471"

print("=" * 80)
print("TESTING BID FIELDS AND GEO BREAKDOWN")
print("=" * 80)

# 1. Fetch bid data
print("\n1. FETCHING BID DATA...")
url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}"
response = requests.get(url)
data = response.json()

print(f"   Response status: {response.status_code}")
print(f"   Number of bids: {len(data)}")

# 2. Show ALL fields from first bid
print("\n2. ALL FIELDS IN BID OBJECT:")
print("-" * 80)

first_bid_id = list(data.keys())[0]
first_bid = data[first_bid_id]

for key, value in sorted(first_bid.items()):
    value_str = str(value)
    if len(value_str) > 60:
        value_str = value_str[:60] + "..."
    print(f"   {key:<20}: {value_str}")

# 3. Check for min/floor/suggested CPM fields
print("\n3. SEARCHING FOR MIN/FLOOR/SUGGESTED CPM FIELDS:")
print("-" * 80)

search_terms = ['min', 'floor', 'suggest', 'recommend', 'price', 'rate', 'base', 'default']
found_any = False

for key, value in first_bid.items():
    key_lower = key.lower()
    for term in search_terms:
        if term in key_lower:
            print(f"   FOUND: {key} = {value}")
            found_any = True
            break

if not found_any:
    print("   No min/floor/suggested CPM fields found in bid object")

# 4. Check geos structure
print("\n4. GEO BREAKDOWN STRUCTURE:")
print("-" * 80)

geos = first_bid.get('geos', {})
if geos:
    print(f"   Number of geos: {len(geos)}")
    print(f"\n   First 3 geo entries:")
    
    for i, (geo_id, geo_data) in enumerate(list(geos.items())[:3]):
        print(f"\n   Geo ID: {geo_id}")
        if isinstance(geo_data, dict):
            for gkey, gvalue in geo_data.items():
                print(f"      {gkey}: {gvalue}")
        else:
            print(f"      Value: {geo_data}")
else:
    print("   No geos data found")

# 5. Check if geos have stats breakdown
print("\n5. CHECKING IF GEOS HAVE STATS BREAKDOWN:")
print("-" * 80)

stats_fields = ['impressions', 'clicks', 'conversions', 'cost', 'revenue', 'ctr', 'ecpm']
geo_has_stats = False

if geos:
    first_geo = list(geos.values())[0]
    if isinstance(first_geo, dict):
        for field in stats_fields:
            if field in first_geo:
                print(f"   FOUND: {field} = {first_geo[field]}")
                geo_has_stats = True
        
        if not geo_has_stats:
            print("   No stats fields found in geo data")
            print(f"   Available geo fields: {list(first_geo.keys())}")

# 6. Check for stats at bid level
print("\n6. BID-LEVEL STATS:")
print("-" * 80)

bid_stats = first_bid.get('stats', {})
if bid_stats:
    for key, value in bid_stats.items():
        print(f"   {key}: {value}")
else:
    print("   No stats object found")

# 7. Try fetching with date range to see if geo stats appear
print("\n7. TESTING WITH DATE RANGE (yesterday):")
print("-" * 80)

from datetime import datetime, timedelta
yesterday = datetime.now() - timedelta(days=1)
date_str = yesterday.strftime('%d/%m/%Y')

url_with_date = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}&startDate={date_str}&endDate={date_str}"
response2 = requests.get(url_with_date)
data2 = response2.json()

if data2:
    first_bid2 = list(data2.values())[0]
    geos2 = first_bid2.get('geos', {})
    
    if geos2:
        first_geo2 = list(geos2.values())[0]
        print(f"   Geo fields with date filter: {list(first_geo2.keys()) if isinstance(first_geo2, dict) else 'N/A'}")
        
        # Check for stats in geo
        if isinstance(first_geo2, dict):
            for field in stats_fields:
                if field in first_geo2:
                    print(f"   GEO STAT FOUND: {field} = {first_geo2[field]}")

# 8. Full JSON dump of one bid for reference
print("\n8. FULL BID JSON (for reference):")
print("-" * 80)
print(json.dumps(first_bid, indent=2)[:2000])
if len(json.dumps(first_bid)) > 2000:
    print("   ... (truncated)")

print("\n" + "=" * 80)
print("TEST COMPLETE")
print("=" * 80)
