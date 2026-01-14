import requests
import json
from datetime import datetime, timedelta

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013205721"

today = datetime.now()
date_30_days_ago = (today - timedelta(days=30)).strftime('%Y-%m-%d')
date_today = today.strftime('%Y-%m-%d')

print("="*70)
print("TESTING DIFFERENT API ENDPOINTS AND PARAMETERS")
print("="*70)

# Current endpoint we're using
print("\n1. Current endpoint: /api/bids/{campaignId}.json")
url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}"
response = requests.get(url)
data = response.json()
print(f"   Status: {response.status_code}")
print(f"   Bids returned: {len(data)}")
if len(data) > 0:
    first_bid = list(data.values())[0]
    print(f"   First bid geos: {first_bid.get('geos')}")

# Try with different parameters
print("\n2. With expand parameter")
url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}&expand=geos"
response = requests.get(url)
print(f"   Status: {response.status_code}")
if response.status_code == 200:
    data = response.json()
    if len(data) > 0:
        first_bid = list(data.values())[0]
        print(f"   First bid geos: {first_bid.get('geos')}")

print("\n3. With includeGeos parameter")
url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}&includeGeos=true"
response = requests.get(url)
print(f"   Status: {response.status_code}")
if response.status_code == 200:
    data = response.json()
    if len(data) > 0:
        first_bid = list(data.values())[0]
        print(f"   First bid geos: {first_bid.get('geos')}")

print("\n4. With detailed parameter")
url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}&detailed=true"
response = requests.get(url)
print(f"   Status: {response.status_code}")
if response.status_code == 200:
    data = response.json()
    if len(data) > 0:
        first_bid = list(data.values())[0]
        print(f"   First bid geos: {first_bid.get('geos')}")

print("\n5. With date range")
url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}&from={date_30_days_ago}&to={date_today}"
response = requests.get(url)
print(f"   Status: {response.status_code}")
if response.status_code == 200:
    data = response.json()
    if len(data) > 0:
        first_bid = list(data.values())[0]
        print(f"   First bid geos: {first_bid.get('geos')}")

# Try /active.json endpoint
print("\n6. Try /api/bids/{campaignId}/active.json")
url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}/active.json?api_key={API_KEY}"
response = requests.get(url)
print(f"   Status: {response.status_code}")
if response.status_code == 200:
    data = response.json()
    print(f"   Type: {type(data)}, Length: {len(data) if isinstance(data, list) else 'N/A'}")
    if isinstance(data, list) and len(data) > 0:
        print(f"   First bid: {json.dumps(data[0], indent=4)}")

# Try alternate endpoints
print("\n7. Try /api/campaigns/{id}/bids.json")
url = f"https://api.trafficjunky.com/api/campaigns/{CAMPAIGN_ID}/bids.json?api_key={API_KEY}"
response = requests.get(url)
print(f"   Status: {response.status_code}")
if response.status_code == 200:
    data = response.json()
    print(f"   Response type: {type(data)}")
    if isinstance(data, dict) and len(data) > 0:
        first_key = list(data.keys())[0]
        print(f"   First entry: {json.dumps(data[first_key], indent=4)[:500]}")

print("\n8. Try /api/placements/{campaignId}.json")
url = f"https://api.trafficjunky.com/api/placements/{CAMPAIGN_ID}.json?api_key={API_KEY}"
response = requests.get(url)
print(f"   Status: {response.status_code}")
if response.status_code == 200:
    data = response.json()
    print(f"   Response: {json.dumps(data, indent=2)[:1000]}")

# Check what fields a single bid has when queried differently
print("\n" + "="*70)
print("CHECKING SINGLE BID DETAIL")
print("="*70)

# Get a bid ID first
url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}"
response = requests.get(url)
data = response.json()
bid_ids = list(data.keys())[:5]  # First 5 bids

print(f"\nComparing 5 bids from same spot to see differences:")
for bid_id in bid_ids:
    bid = data[bid_id]
    geos = bid.get('geos', {})
    geo_info = list(geos.values())[0] if geos else {}
    print(f"  Bid {bid_id}: CPM=${bid.get('bid')} | Spot={bid.get('spot_name')} | GeoID={geo_info.get('geoId')} | Country={geo_info.get('countryCode')}")
