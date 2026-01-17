import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013232471"

print("="*80)
print("TESTING /api/campaigns/stats.json ENDPOINT")
print("="*80)

# Test the campaigns stats endpoint
url = f"https://api.trafficjunky.com/api/campaigns/stats.json?api_key={API_KEY}"
print(f"\n1. GET /api/campaigns/stats.json (all campaigns)")
response = requests.get(url)
print(f"   Status: {response.status_code}")
if response.status_code == 200:
    data = response.json()
    print(f"   Type: {type(data)}")
    if isinstance(data, list):
        print(f"   Count: {len(data)} campaigns")
        if len(data) > 0:
            print(f"\n   First campaign structure:")
            print(json.dumps(data[0], indent=4))
    elif isinstance(data, dict):
        print(f"   Keys: {list(data.keys())[:10]}")
        print(f"\n   Full response:")
        print(json.dumps(data, indent=4)[:2000])

# Test with specific campaign ID
print(f"\n2. GET /api/campaigns/stats.json?campaignIds={CAMPAIGN_ID}")
url = f"https://api.trafficjunky.com/api/campaigns/stats.json?api_key={API_KEY}&campaignIds={CAMPAIGN_ID}"
response = requests.get(url)
print(f"   Status: {response.status_code}")
if response.status_code == 200:
    data = response.json()
    print(f"\n   Response:")
    print(json.dumps(data, indent=4)[:3000])

# Test with campaign ID in path
print(f"\n3. GET /api/campaigns/{CAMPAIGN_ID}/stats.json")
url = f"https://api.trafficjunky.com/api/campaigns/{CAMPAIGN_ID}/stats.json?api_key={API_KEY}"
response = requests.get(url)
print(f"   Status: {response.status_code}")
if response.status_code == 200:
    data = response.json()
    print(f"\n   Response:")
    print(json.dumps(data, indent=4)[:3000])
