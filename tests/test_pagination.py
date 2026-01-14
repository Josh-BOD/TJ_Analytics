import requests
import json
from datetime import datetime, timedelta

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013205721"

today = datetime.now()
date_30_days_ago = (today - timedelta(days=30)).strftime('%Y-%m-%d')
date_today = today.strftime('%Y-%m-%d')

# Try with pagination
found = False
offset = 1
all_campaigns = {}

while not found:
    url = f"https://api.trafficjunky.com/api/campaigns/bids/stats.json?api_key={API_KEY}&from={date_30_days_ago}&to={date_today}&limit=500&offset={offset}"
    print(f"Fetching offset {offset}...")
    
    response = requests.get(url)
    if response.status_code != 200:
        print(f"Error: {response.status_code}")
        break
        
    data = response.json()
    print(f"  Got {len(data)} campaigns")
    
    if len(data) == 0:
        break
        
    all_campaigns.update(data)
    
    if CAMPAIGN_ID in data:
        print(f"\n✅ FOUND campaign {CAMPAIGN_ID} at offset {offset}!")
        campaign = data[CAMPAIGN_ID]
        bids = campaign.get('bids', [])
        
        # Collect countries
        countries = {}
        for bid in bids:
            country = bid.get('countryCode', 'Unknown')
            countries[country] = countries.get(country, 0) + 1
        
        print(f"Countries: {countries}")
        found = True
        break
    
    offset += 500
    
    if offset > 2000:  # Safety limit
        break

print(f"\nTotal campaigns scanned: {len(all_campaigns)}")

if not found:
    print(f"\nCampaign {CAMPAIGN_ID} not found in any page!")
    
    # List all campaign IDs to see what's there
    print("\nAll campaign IDs found:")
    for cid in sorted(all_campaigns.keys()):
        name = all_campaigns[cid].get('campaignName', 'N/A')
        print(f"  {cid}: {name}")
