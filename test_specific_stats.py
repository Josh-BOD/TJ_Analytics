import requests
import json
from datetime import datetime, timedelta

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013232471"

# Calculate date range (last 30 days)
end_date = datetime.now()
start_date = end_date - timedelta(days=30)
start_str = start_date.strftime("%d/%m/%Y")  # DD/MM/YYYY format per API docs
end_str = end_date.strftime("%d/%m/%Y")

print("="*80)
print("TRYING DIFFERENT WAYS TO GET CAMPAIGN STATS")
print("="*80)

# Try 1: With campaignIds parameter
print(f"\n1. /api/campaigns/stats.json?campaignIds={CAMPAIGN_ID}")
url = f"https://api.trafficjunky.com/api/campaigns/stats.json?api_key={API_KEY}&campaignIds={CAMPAIGN_ID}"
response = requests.get(url)
print(f"   Status: {response.status_code}")
data = response.json()
print(f"   Campaigns returned: {list(data.keys())[:5]}")
if CAMPAIGN_ID in data:
    print(f"   ✅ Found! eCPM: ${data[CAMPAIGN_ID].get('ecpm')}")
else:
    print(f"   ❌ Campaign {CAMPAIGN_ID} not in response")

# Try 2: With date range
print(f"\n2. /api/campaigns/stats.json with dates {start_str} to {end_str}")
url = f"https://api.trafficjunky.com/api/campaigns/stats.json?api_key={API_KEY}&startDate={start_str}&endDate={end_str}"
response = requests.get(url)
print(f"   Status: {response.status_code}")
data = response.json()
print(f"   Campaigns returned: {len(data)}")
if CAMPAIGN_ID in data:
    print(f"   ✅ Found! eCPM: ${data[CAMPAIGN_ID].get('ecpm')}")
else:
    print(f"   ❌ Campaign {CAMPAIGN_ID} not in response")

# Try 3: With higher limit/offset
print(f"\n3. /api/campaigns/stats.json with limit=100")
url = f"https://api.trafficjunky.com/api/campaigns/stats.json?api_key={API_KEY}&limit=100"
response = requests.get(url)
print(f"   Status: {response.status_code}")
data = response.json()
print(f"   Campaigns returned: {len(data)}")
if CAMPAIGN_ID in data:
    print(f"   ✅ Found! eCPM: ${data[CAMPAIGN_ID].get('ecpm')}")
else:
    print(f"   ❌ Campaign {CAMPAIGN_ID} not in response")
    # Check if any of the target campaigns are found
    target_ids = ["1013232471", "1013225801", "1013205721"]
    for tid in target_ids:
        if tid in data:
            print(f"   Found {tid}: eCPM ${data[tid].get('ecpm')}")

# Try 4: Check /api/campaigns/bids/stats endpoint (different endpoint)
print(f"\n4. /api/campaigns/bids/stats.json (different endpoint)")
url = f"https://api.trafficjunky.com/api/campaigns/bids/stats.json?api_key={API_KEY}&limit=100"
response = requests.get(url)
print(f"   Status: {response.status_code}")
if response.status_code == 200:
    data = response.json()
    if isinstance(data, list):
        print(f"   Campaigns returned: {len(data)}")
        campaign_ids_found = [str(c.get('campaignId', c.get('campaign_id', ''))) for c in data[:20]]
        print(f"   First 20 campaign IDs: {campaign_ids_found}")
        for c in data:
            cid = str(c.get('campaignId', c.get('campaign_id', '')))
            if cid == CAMPAIGN_ID:
                print(f"   ✅ Found! Campaign data: {json.dumps(c, indent=4)[:500]}")
                break
    else:
        print(f"   Type: {type(data)}")
        print(f"   Keys: {list(data.keys())[:10]}")
