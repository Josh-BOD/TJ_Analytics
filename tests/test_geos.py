import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013205721"

# Test the bids endpoint
url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}"

print(f"Fetching bids for campaign {CAMPAIGN_ID}...")
print(f"URL: {url.replace(API_KEY, 'HIDDEN')}")
print()

response = requests.get(url)
print(f"Status: {response.status_code}")

if response.status_code == 200:
    data = response.json()
    
    print(f"Total bids: {len(data)}")
    print()
    
    # Collect all unique countries
    all_countries = set()
    country_counts = {}
    
    # Print first 5 bids with full geos
    print("=== FIRST 5 BIDS WITH FULL GEOS ===")
    for i, (bid_id, bid) in enumerate(data.items()):
        if i < 5:
            print(f"\nBid {bid_id}:")
            print(f"  spot_name: {bid.get('spot_name')}")
            print(f"  bid: ${bid.get('bid')}")
            print(f"  geos: {json.dumps(bid.get('geos'), indent=4)}")
        
        # Collect country info
        geos = bid.get('geos', {})
        for geo_id, geo in geos.items():
            country = geo.get('countryCode', 'Unknown')
            all_countries.add(country)
            country_counts[country] = country_counts.get(country, 0) + 1
    
    print("\n" + "="*50)
    print("=== COUNTRY SUMMARY ===")
    print(f"Unique countries found: {all_countries}")
    print(f"Country counts: {country_counts}")
    
    # Print a bid from each unique country
    print("\n=== ONE BID PER COUNTRY ===")
    seen_countries = set()
    for bid_id, bid in data.items():
        geos = bid.get('geos', {})
        for geo_id, geo in geos.items():
            country = geo.get('countryCode')
            if country and country not in seen_countries:
                seen_countries.add(country)
                print(f"\n{country} - Bid {bid_id}:")
                print(f"  spot_name: {bid.get('spot_name')}")
                print(f"  geos: {json.dumps(geos, indent=4)}")
else:
    print(f"Error: {response.text}")
