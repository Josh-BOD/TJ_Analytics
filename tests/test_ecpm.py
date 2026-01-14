import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013232471"

url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}"
response = requests.get(url)
data = response.json()

print("="*80)
print("COMPARING YOUR CPM (bid) vs eCPM (stats.ecpm)")
print("="*80)
print(f"\n{'Bid ID':<15} {'Your CPM':<12} {'eCPM':<12} {'Same?':<8} {'Spot Name'}")
print("-"*80)

for bid_id, bid in data.items():
    your_cpm = bid.get('bid', 'N/A')
    stats = bid.get('stats', {})
    ecpm = stats.get('ecpm', 'N/A')
    spot = bid.get('spot_name', 'N/A')[:35]
    
    # Check if they're the same
    same = "YES" if str(your_cpm) == str(ecpm) else "NO"
    
    print(f"{bid_id:<15} ${your_cpm:<11} ${ecpm:<11} {same:<8} {spot}")

print("\n" + "="*80)
print("RAW DATA for first bid:")
print("="*80)
first_bid = list(data.values())[0]
print(f"bid field: {first_bid.get('bid')} (type: {type(first_bid.get('bid'))})")
print(f"stats.ecpm: {first_bid.get('stats', {}).get('ecpm')} (type: {type(first_bid.get('stats', {}).get('ecpm'))})")
