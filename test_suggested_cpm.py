import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013232471"
BID_ID = "1201505001"
SPOT_ID = "1845481"

print("="*70)
print("SEARCHING FOR SUGGESTED CPM / PRICING DATA")
print("="*70)

# 1. Check if bid response has any pricing suggestions
print("\n1. Checking bid data for pricing fields...")
url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}"
response = requests.get(url)
data = response.json()
first_bid = list(data.values())[0]
print(f"   All fields in bid: {list(first_bid.keys())}")
# Check for any pricing-related fields
for key, value in first_bid.items():
    key_lower = key.lower()
    if any(x in key_lower for x in ['price', 'suggest', 'recommend', 'min', 'max', 'avg', 'floor', 'ceiling', 'traffic']):
        print(f"   {key}: {value}")

# 2. Check spot endpoint for pricing
print("\n2. Checking spot endpoints...")
spot_endpoints = [
    f"/api/spots/{SPOT_ID}.json",
    f"/api/spots/{SPOT_ID}/pricing.json",
    f"/api/spots/{SPOT_ID}/stats.json",
    f"/api/spot/{SPOT_ID}.json",
]
for endpoint in spot_endpoints:
    url = f"https://api.trafficjunky.com{endpoint}?api_key={API_KEY}"
    response = requests.get(url)
    print(f"   {endpoint}: {response.status_code}")
    if response.status_code == 200:
        data = response.json()
        print(f"   Response: {json.dumps(data, indent=2)[:500]}")

# 3. Check for traffic/pricing endpoint
print("\n3. Checking other potential endpoints...")
other_endpoints = [
    "/api/traffic.json",
    "/api/pricing.json",
    "/api/rates.json",
    f"/api/campaigns/{CAMPAIGN_ID}/pricing.json",
    f"/api/campaigns/{CAMPAIGN_ID}/traffic.json",
]
for endpoint in other_endpoints:
    url = f"https://api.trafficjunky.com{endpoint}?api_key={API_KEY}"
    response = requests.get(url)
    print(f"   {endpoint}: {response.status_code}")
    if response.status_code == 200:
        data = response.json()
        print(f"   Response: {json.dumps(data, indent=2)[:300]}")

# 4. Check the full API spec for any pricing/suggestion endpoints
print("\n4. Checking API spec for pricing endpoints...")
spec_url = "https://api.trafficjunky.com/docs/api-docs.json"
response = requests.get(spec_url)
spec = response.json()
paths = spec.get('paths', {})
for path in paths:
    if any(x in path.lower() for x in ['price', 'suggest', 'traffic', 'rate', 'floor', 'estimate']):
        print(f"   Found: {path}")

# 5. Check the traffic field in more detail
print("\n5. Analyzing 'traffic' field from bids...")
url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}"
response = requests.get(url)
data = response.json()
print("   Bid ID | CPM | Traffic Share | Spot Name")
for bid_id, bid in list(data.items())[:10]:
    cpm = bid.get('bid', 'N/A')
    traffic = bid.get('traffic', 'N/A')
    spot = bid.get('spot_name', 'N/A')[:30]
    print(f"   {bid_id} | ${cpm} | {traffic:.10f} | {spot}")
