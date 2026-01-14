import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013205721"

# Test campaigns/bids/stats endpoint
url = f"https://api.trafficjunky.com/api/campaigns/bids/stats.json?api_key={API_KEY}&campaignIds={CAMPAIGN_ID}"
print(f"Testing campaigns/bids/stats.json for campaign {CAMPAIGN_ID}")
print()

response = requests.get(url)
print(f"Status: {response.status_code}")

if response.status_code == 200:
    data = response.json()
    
    # It should be a list or dict
    if isinstance(data, list):
        print(f"Got list with {len(data)} items")
        for campaign in data:
            if str(campaign.get('campaignId')) == CAMPAIGN_ID:
                print(f"\n=== CAMPAIGN {CAMPAIGN_ID} ===")
                print(f"campaignName: {campaign.get('campaignName')}")
                
                bids = campaign.get('bids', [])
                print(f"Total bids: {len(bids)}")
                
                # Show first 5 bids
                print("\nFirst 5 bids:")
                for i, bid in enumerate(bids[:5]):
                    print(f"\nBid {i+1}:")
                    print(json.dumps(bid, indent=2))
                
                # Collect all countries
                countries = {}
                for bid in bids:
                    country = bid.get('countryCode') or bid.get('countryName') or bid.get('country')
                    if country:
                        countries[country] = countries.get(country, 0) + 1
                
                print(f"\n=== COUNTRIES FOUND ===")
                print(f"Countries: {countries}")
                break
    else:
        print("Response is not a list:")
        print(json.dumps(data, indent=2)[:2000])
else:
    print(f"Error: {response.text[:500]}")

# Also check what geoId 1055650131 means by looking at a different endpoint
print("\n" + "="*60)
print("Checking if geoId can be decoded...")

# The geoId might be a compound ID - let's see if there's a pattern
# AU geoId from campaign 1013225801 was 1055760191
# AU geoId from campaign 1013205721 is 1055650131
# US geoId from campaign 1013232471 was 1055790301

print(f"Campaign 1013232471 (US only): geoId 1055790301")
print(f"Campaign 1013225801 (AU only): geoId 1055760191")  
print(f"Campaign 1013205721 (5 countries): geoId 1055650131")
print("\nThese are different geoIds - the one for 5 countries might be a 'group' ID")
