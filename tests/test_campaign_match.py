import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"

# Campaign IDs from Legend sheet
CAMPAIGN_IDS = ["1013232471", "1013225801", "1013205721"]

print("="*80)
print("CHECKING IF CAMPAIGNS EXIST IN /api/campaigns/stats.json")
print("="*80)

# Fetch campaign stats
stats_url = f"https://api.trafficjunky.com/api/campaigns/stats.json?api_key={API_KEY}"
response = requests.get(stats_url)
stats_data = response.json()

print(f"\nStats endpoint returned {len(stats_data)} campaigns")
print(f"Campaign IDs in stats: {list(stats_data.keys())[:20]}...")

print("\n" + "-"*80)
print("CHECKING EACH CAMPAIGN FROM LEGEND:")
print("-"*80)

for cid in CAMPAIGN_IDS:
    if cid in stats_data:
        stats = stats_data[cid]
        print(f"\n✅ Campaign {cid} FOUND in stats:")
        print(f"   Name: {stats.get('campaign_name', 'N/A')}")
        print(f"   Avg eCPM: ${stats.get('ecpm', 'N/A')}")
        print(f"   Cost: ${stats.get('cost', 'N/A')}")
        print(f"   Impressions: {stats.get('impressions', 'N/A')}")
    else:
        print(f"\n❌ Campaign {cid} NOT FOUND in stats endpoint!")
        
        # Fetch bid data to see what we'd get
        bids_url = f"https://api.trafficjunky.com/api/bids/{cid}.json?api_key={API_KEY}"
        bids_response = requests.get(bids_url)
        if bids_response.status_code == 200:
            bids = bids_response.json()
            if bids:
                first_bid = list(bids.values())[0]
                print(f"   (Bid endpoint works - per-bid eCPM: ${first_bid.get('stats', {}).get('ecpm', 'N/A')})")
