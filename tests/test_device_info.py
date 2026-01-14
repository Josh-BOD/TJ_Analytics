import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013232471"

# Get bids
url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}"
response = requests.get(url)
data = response.json()

print("="*70)
print("CHECKING FOR DEVICE/OS INFORMATION IN API")
print("="*70)

# Get first bid and show ALL fields
first_bid = list(data.values())[0]
print("\nAll fields in bid object:")
for key, value in first_bid.items():
    print(f"  {key}: {json.dumps(value) if isinstance(value, (dict, list)) else value}")

# Check all unique spot names
print("\n\nAll unique spot names:")
spot_names = set()
for bid in data.values():
    spot_names.add(bid.get('spot_name', 'N/A'))
for name in sorted(spot_names):
    print(f"  - {name}")

# Search for any field containing 'device', 'os', 'android', 'ios', 'platform'
print("\n\nSearching for device/OS related fields:")
bid_str = json.dumps(first_bid).lower()
keywords = ['device', 'android', 'ios', 'platform', 'operating', 'os']
for kw in keywords:
    if kw in bid_str:
        print(f"  Found '{kw}' in bid data!")
    else:
        print(f"  '{kw}' NOT found")
