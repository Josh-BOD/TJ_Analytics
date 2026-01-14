import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013205721"

# Test the campaign details endpoint
url = f"https://api.trafficjunky.com/api/campaigns/{CAMPAIGN_ID}.json?api_key={API_KEY}"

print(f"Fetching campaign details for {CAMPAIGN_ID}...")
print()

response = requests.get(url)
print(f"Status: {response.status_code}")

if response.status_code == 200:
    data = response.json()
    
    print("\n=== FULL CAMPAIGN DATA ===")
    print(json.dumps(data, indent=2))
    
    print("\n=== KEY FIELDS ===")
    print(f"campaign_id: {data.get('campaign_id')}")
    print(f"campaign_name: {data.get('campaign_name')}")
    print(f"campaign_target_group: {data.get('campaign_target_group')}")
    print(f"number_of_bids: {data.get('number_of_bids')}")
    print(f"spots: {data.get('spots')}")
else:
    print(f"Error: {response.text}")

# Also try the campaigns/bids/stats endpoint
print("\n" + "="*60)
print("=== TRYING campaigns/bids/stats.json ===")
url2 = f"https://api.trafficjunky.com/api/campaigns/bids/stats.json?api_key={API_KEY}&campaignIds={CAMPAIGN_ID}"
print(f"URL: {url2.replace(API_KEY, 'HIDDEN')}")

response2 = requests.get(url2)
print(f"Status: {response2.status_code}")

if response2.status_code == 200:
    data2 = response2.json()
    
    # Find our campaign
    for campaign in data2 if isinstance(data2, list) else [data2]:
        if str(campaign.get('campaignId')) == CAMPAIGN_ID:
            print(f"\n=== FOUND CAMPAIGN {CAMPAIGN_ID} ===")
            print(f"campaignName: {campaign.get('campaignName')}")
            print(f"bids count: {len(campaign.get('bids', []))}")
            
            # Show first 3 bids
            bids = campaign.get('bids', [])
            print(f"\nFirst 3 bids:")
            for bid in bids[:3]:
                print(json.dumps(bid, indent=2))
            
            # Collect countries from bids
            countries = set()
            for bid in bids:
                country = bid.get('countryCode') or bid.get('country')
                if country:
                    countries.add(country)
            print(f"\nCountries in bids: {countries}")
            break
else:
    print(f"Error: {response2.text}")
