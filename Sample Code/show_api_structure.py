import requests
import json
from datetime import datetime, timedelta

# Your API key
API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
API_URL = "https://api.trafficjunky.com/api/campaigns/bids/stats.json"

print("=" * 100)
print("TRAFFICJUNKY API DATA STRUCTURE")
print("=" * 100)

# Get last 3 days of data
end_date = datetime.now() - timedelta(days=1)
start_date = end_date - timedelta(days=2)

params = {
    'api_key': API_KEY,
    'startDate': start_date.strftime('%d/%m/%Y'),
    'endDate': end_date.strftime('%d/%m/%Y'),
    'limit': 3,
    'offset': 1
}

print(f"\nFetching sample data from {params['startDate']} to {params['endDate']}...")

response = requests.get(API_URL, params=params)
data = response.json()

# Get first campaign
if isinstance(data, dict):
    first_campaign = list(data.values())[0]
elif isinstance(data, list):
    first_campaign = data[0]
else:
    print("ERROR: Unexpected data format")
    exit()

print(f"Total campaigns returned: {len(data)}")

print("\n" + "=" * 100)
print("FULL JSON STRUCTURE OF FIRST CAMPAIGN")
print("=" * 100)
print(json.dumps(first_campaign, indent=2))

print("\n" + "=" * 100)
print("ALL FIELD NAMES")
print("=" * 100)
print(", ".join(sorted(first_campaign.keys())))

print("\n" + "=" * 100)
print("MAIN CAMPAIGN FIELDS")
print("=" * 100)
print(f"{'Field Name':<25} {'Sample Value':<40} {'Type':<15}")
print("-" * 100)

for key, value in sorted(first_campaign.items()):
    if key in ['bids', 'spots']:
        continue  # Handle separately
    value_str = str(value)[:40] if value else ""
    print(f"{key:<25} {value_str:<40} {type(value).__name__:<15}")

print("\n" + "=" * 100)
print("BIDS ARRAY (Country/Geo Targeting)")
print("=" * 100)

if 'bids' in first_campaign and first_campaign['bids']:
    print(f"Number of bids: {len(first_campaign['bids'])}")
    print("\nFirst bid structure:")
    print(json.dumps(first_campaign['bids'][0], indent=2))
    
    print("\nAll countries targeted by this campaign:")
    print(f"{'Country Code':<15} {'Country Name':<30} {'Bid Amount':<15} {'Region':<20}")
    print("-" * 100)
    for bid in first_campaign['bids']:
        country_code = bid.get('countryCode', 'N/A')
        country_name = bid.get('countryName', 'N/A')
        bid_amount = bid.get('bid', 'N/A')
        region = bid.get('regionName', 'N/A') or bid.get('regionCode', 'N/A') or ''
        print(f"{country_code:<15} {country_name:<30} {bid_amount:<15} {region:<20}")
else:
    print("No bids data available")

print("\n" + "=" * 100)
print("SPOTS ARRAY (Ad Placements)")
print("=" * 100)

if 'spots' in first_campaign and first_campaign['spots']:
    print(f"Number of spots: {len(first_campaign['spots'])}")
    print(f"\n{'Spot ID':<15} {'Spot Name':<60}")
    print("-" * 100)
    for spot in first_campaign['spots']:
        spot_id = spot.get('id', 'N/A')
        spot_name = spot.get('name', 'N/A')
        print(f"{spot_id:<15} {spot_name:<60}")
else:
    print("No spots data available")

print("\n" + "=" * 100)
print("KEY FINDINGS")
print("=" * 100)
findings = [
    "✓ Country data IS available in 'bids' array",
    "✓ Each bid contains: countryCode, countryName, regionCode, regionName, city",
    "✗ Performance stats (clicks, impressions, cost) are AGGREGATED (not per-country)",
    "✗ To get per-country performance, you'd need Selenium web scraping",
    "",
    "WHAT THE API PROVIDES:",
    "  • Total impressions, clicks, cost, conversions per CAMPAIGN",
    "  • Which countries the campaign TARGETS",
    "  • Bid amounts per country",
    "",
    "WHAT THE API DOESN'T PROVIDE:",
    "  • Impressions by country",
    "  • Clicks by country", 
    "  • Cost by country",
    "  • Conversions by country"
]

for finding in findings:
    print(finding)

print("\n" + "=" * 100)




