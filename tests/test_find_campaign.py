import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013205721"

# Get all campaigns from bids/stats
url = f"https://api.trafficjunky.com/api/campaigns/bids/stats.json?api_key={API_KEY}"
print(f"Fetching all campaigns from bids/stats.json...")

response = requests.get(url)
print(f"Status: {response.status_code}")

if response.status_code == 200:
    data = response.json()
    
    print(f"Total campaigns returned: {len(data)}")
    
    # Check if our campaign is there
    if CAMPAIGN_ID in data:
        campaign = data[CAMPAIGN_ID]
        print(f"\n=== FOUND CAMPAIGN {CAMPAIGN_ID} ===")
        print(f"campaignName: {campaign.get('campaignName')}")
        print(f"numberOfBids: {campaign.get('numberOfBids')}")
        
        bids = campaign.get('bids', [])
        print(f"Bids in response: {len(bids)}")
        
        # Show first 10 bids
        print("\nFirst 10 bids:")
        for i, bid in enumerate(bids[:10]):
            print(f"  {i+1}. {bid.get('countryCode')} - ${bid.get('bid')} - placementId: {bid.get('placementId')}")
        
        # Collect all countries
        countries = {}
        for bid in bids:
            country = bid.get('countryCode', 'Unknown')
            countries[country] = countries.get(country, 0) + 1
        
        print(f"\n=== ALL COUNTRIES ===")
        for country, count in sorted(countries.items()):
            print(f"  {country}: {count} bids")
    else:
        print(f"\nCampaign {CAMPAIGN_ID} NOT found in response!")
        print(f"Available campaign IDs (first 20):")
        for i, cid in enumerate(list(data.keys())[:20]):
            print(f"  {cid}: {data[cid].get('campaignName', 'N/A')}")
else:
    print(f"Error: {response.text[:500]}")
