import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013205721"

# Try various geo-related endpoints
endpoints = [
    f"/api/campaigns/{CAMPAIGN_ID}/geos.json",
    f"/api/campaigns/{CAMPAIGN_ID}/targeting.json",
    f"/api/campaigns/{CAMPAIGN_ID}/audience.json",
    f"/api/campaigns/{CAMPAIGN_ID}/settings.json",
    f"/api/geos.json",
    f"/api/geos/{CAMPAIGN_ID}.json",
]

for endpoint in endpoints:
    url = f"https://api.trafficjunky.com{endpoint}?api_key={API_KEY}"
    print(f"\n{'='*60}")
    print(f"Testing: {endpoint}")
    
    try:
        response = requests.get(url)
        print(f"Status: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            print(f"Response: {json.dumps(data, indent=2)[:1000]}")
        else:
            print(f"Error: {response.text[:200]}")
    except Exception as e:
        print(f"Exception: {e}")

# Also try to get ONE specific bid with different format
print(f"\n{'='*60}")
print("Testing single bid detail with various formats...")

# Get first bid ID from the bids endpoint
url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}"
response = requests.get(url)
data = response.json()
first_bid_id = list(data.keys())[0]
print(f"First bid ID: {first_bid_id}")

# Now try to get detailed info for this specific bid
bid_endpoints = [
    f"/api/bids/{first_bid_id}.json",
    f"/api/bid/{first_bid_id}.json",
    f"/api/bids/{first_bid_id}/geos.json",
    f"/api/bids/{first_bid_id}/details.json",
]

for endpoint in bid_endpoints:
    url = f"https://api.trafficjunky.com{endpoint}?api_key={API_KEY}"
    print(f"\n{'='*60}")
    print(f"Testing: {endpoint}")
    
    try:
        response = requests.get(url)
        print(f"Status: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            print(f"Response: {json.dumps(data, indent=2)[:1500]}")
    except Exception as e:
        print(f"Exception: {e}")
