"""
Test various endpoints for:
1. Min CPM / floor price
2. Geo-level stats breakdown
"""

import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013232471"
BID_ID = "1201505001"
SPOT_ID = "1845481"
GEO_ID = "1055790301"

from datetime import datetime, timedelta
yesterday = datetime.now() - timedelta(days=1)
date_str = yesterday.strftime('%d/%m/%Y')

print("=" * 80)
print("TESTING GEO STATS AND MIN CPM ENDPOINTS")
print("=" * 80)

# 1. Try various geo-related endpoints
print("\n1. TESTING GEO-RELATED ENDPOINTS:")
print("-" * 80)

geo_endpoints = [
    f"/api/bids/{BID_ID}/geos.json",
    f"/api/bids/{BID_ID}/geos/{GEO_ID}.json",
    f"/api/bids/{BID_ID}/geos/{GEO_ID}/stats.json",
    f"/api/campaigns/{CAMPAIGN_ID}/geos.json",
    f"/api/campaigns/{CAMPAIGN_ID}/geos/stats.json",
    f"/api/geos/{GEO_ID}.json",
    f"/api/geos/{GEO_ID}/stats.json",
    f"/api/campaigns/bids/stats.json?groupBy=geo",
]

for endpoint in geo_endpoints:
    if "?" in endpoint:
        url = f"https://api.trafficjunky.com{endpoint}&api_key={API_KEY}&startDate={date_str}&endDate={date_str}"
    else:
        url = f"https://api.trafficjunky.com{endpoint}?api_key={API_KEY}&startDate={date_str}&endDate={date_str}"
    
    response = requests.get(url)
    print(f"\n   {endpoint}")
    print(f"   Status: {response.status_code}")
    
    if response.status_code == 200:
        data = response.json()
        preview = json.dumps(data, indent=2)[:500]
        print(f"   Response: {preview}")
        if len(json.dumps(data)) > 500:
            print("   ... (truncated)")

# 2. Check spot endpoint for min CPM / floor price
print("\n\n2. CHECKING SPOT ENDPOINTS FOR MIN CPM:")
print("-" * 80)

spot_endpoints = [
    f"/api/spots/{SPOT_ID}.json",
    f"/api/spots.json",
    f"/api/spots/{SPOT_ID}/pricing.json",
    f"/api/spots/{SPOT_ID}/rates.json",
]

for endpoint in spot_endpoints:
    url = f"https://api.trafficjunky.com{endpoint}?api_key={API_KEY}"
    response = requests.get(url)
    print(f"\n   {endpoint}: {response.status_code}")
    
    if response.status_code == 200:
        data = response.json()
        
        # Check if it's a list or dict
        if isinstance(data, list) and len(data) > 0:
            # Show first item
            first = data[0]
            print(f"   First item keys: {list(first.keys()) if isinstance(first, dict) else 'N/A'}")
            
            # Look for min/floor fields
            if isinstance(first, dict):
                for key in first.keys():
                    if any(x in key.lower() for x in ['min', 'floor', 'price', 'rate', 'cpm']):
                        print(f"   FOUND: {key} = {first[key]}")
        elif isinstance(data, dict):
            # Look for min/floor fields
            for key in data.keys():
                if any(x in key.lower() for x in ['min', 'floor', 'price', 'rate', 'cpm']):
                    print(f"   FOUND: {key} = {data[key]}")
            
            # Show all keys
            print(f"   Available fields: {list(data.keys())[:15]}")

# 3. Check API spec for geo stats endpoints
print("\n\n3. CHECKING API SPEC FOR GEO/PRICING ENDPOINTS:")
print("-" * 80)

spec_url = "https://api.trafficjunky.com/docs/api-docs.json"
response = requests.get(spec_url)
spec = response.json()

paths = spec.get('paths', {})
relevant_paths = []

for path in paths:
    path_lower = path.lower()
    if any(x in path_lower for x in ['geo', 'country', 'region', 'price', 'floor', 'min', 'rate']):
        relevant_paths.append(path)
        print(f"   {path}")

if not relevant_paths:
    print("   No geo/pricing specific endpoints found in API spec")

# 4. Try bids stats with different parameters
print("\n\n4. TESTING BIDS STATS WITH VARIOUS PARAMETERS:")
print("-" * 80)

params_to_try = [
    {"groupBy": "geo"},
    {"groupBy": "country"},
    {"breakdown": "geo"},
    {"includeGeoStats": "true"},
    {"geoStats": "true"},
]

base_url = f"https://api.trafficjunky.com/api/campaigns/bids/stats.json?api_key={API_KEY}&startDate={date_str}&endDate={date_str}&limit=5&offset=1"

for params in params_to_try:
    param_str = "&".join([f"{k}={v}" for k, v in params.items()])
    url = f"{base_url}&{param_str}"
    response = requests.get(url)
    print(f"\n   Params: {params}")
    print(f"   Status: {response.status_code}")
    
    if response.status_code == 200:
        data = response.json()
        # Check if response structure changed
        if isinstance(data, dict) and len(data) > 0:
            first_key = list(data.keys())[0]
            first_item = data[first_key]
            
            # Check for geo breakdown in bids
            if isinstance(first_item, dict) and 'bids' in first_item:
                bids = first_item['bids']
                if bids and len(bids) > 0:
                    first_bid = bids[0] if isinstance(bids, list) else list(bids.values())[0]
                    
                    # Check for geo stats
                    geos = first_bid.get('geos', {}) if isinstance(first_bid, dict) else {}
                    if geos:
                        first_geo = list(geos.values())[0] if isinstance(geos, dict) else geos[0]
                        print(f"   Geo fields: {list(first_geo.keys()) if isinstance(first_geo, dict) else 'N/A'}")

print("\n" + "=" * 80)
print("TEST COMPLETE")
print("=" * 80)
