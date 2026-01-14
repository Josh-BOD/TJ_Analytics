import requests
import json
from datetime import datetime, timedelta

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013205721"

# Try with different date ranges
today = datetime.now()
date_30_days_ago = (today - timedelta(days=30)).strftime('%Y-%m-%d')
date_today = today.strftime('%Y-%m-%d')

# Try various date parameter formats
date_formats = [
    f"from={date_30_days_ago}&to={date_today}",
    f"startDate={date_30_days_ago}&endDate={date_today}",
    f"dateFrom={date_30_days_ago}&dateTo={date_today}",
]

for params in date_formats:
    url = f"https://api.trafficjunky.com/api/campaigns/bids/stats.json?api_key={API_KEY}&{params}"
    print(f"\n{'='*60}")
    print(f"Testing with params: {params}")
    
    response = requests.get(url)
    print(f"Status: {response.status_code}")
    
    if response.status_code == 200:
        data = response.json()
        print(f"Campaigns returned: {len(data)}")
        
        if CAMPAIGN_ID in data:
            print(f"✅ FOUND campaign {CAMPAIGN_ID}!")
            campaign = data[CAMPAIGN_ID]
            bids = campaign.get('bids', [])
            
            # Collect countries
            countries = {}
            for bid in bids:
                country = bid.get('countryCode', 'Unknown')
                countries[country] = countries.get(country, 0) + 1
            
            print(f"Countries: {countries}")
            break
        else:
            print(f"Campaign {CAMPAIGN_ID} not found")

# Also try with campaignId filter
print(f"\n{'='*60}")
print("Testing with explicit campaignId filter...")
url = f"https://api.trafficjunky.com/api/campaigns/bids/stats.json?api_key={API_KEY}&campaignId={CAMPAIGN_ID}&from={date_30_days_ago}&to={date_today}"
response = requests.get(url)
print(f"Status: {response.status_code}")
if response.status_code == 200:
    data = response.json()
    print(f"Campaigns: {len(data)}")
    if CAMPAIGN_ID in data:
        print("✅ Found!")
        campaign = data[CAMPAIGN_ID]
        bids = campaign.get('bids', [])
        countries = {}
        for bid in bids:
            country = bid.get('countryCode', 'Unknown')
            countries[country] = countries.get(country, 0) + 1
        print(f"Countries: {countries}")
