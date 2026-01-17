import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
CAMPAIGN_ID = "1013232471"

url = f"https://api.trafficjunky.com/api/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}"
response = requests.get(url)
data = response.json()

print("="*100)
print("COMPARING: Your CPM vs API eCPM vs CALCULATED eCPM (Cost/Impressions*1000)")
print("="*100)
print(f"\n{'Spot Name':<35} {'Your CPM':<10} {'API eCPM':<10} {'Calc eCPM':<12} {'Cost':<10} {'Impr':<10}")
print("-"*100)

for bid_id, bid in data.items():
    your_cpm = float(bid.get('bid', 0))
    stats = bid.get('stats', {})
    api_ecpm = float(stats.get('ecpm', 0))
    cost = float(stats.get('revenue', 0))  # "revenue" is your cost/spend
    impressions = int(stats.get('impressions', 0))
    spot = bid.get('spot_name', 'N/A')[:33]
    
    # Calculate eCPM ourselves
    calc_ecpm = (cost / impressions * 1000) if impressions > 0 else 0
    
    print(f"{spot:<35} ${your_cpm:<9.3f} ${api_ecpm:<9.3f} ${calc_ecpm:<11.3f} ${cost:<9.2f} {impressions:<10}")

print("\n" + "="*100)
print("CONCLUSION:")
print("="*100)
print("API's 'ecpm' field = basically your bid price (not useful)")
print("Calculated eCPM = Cost/Impressions*1000 (what you actually paid per 1000 impressions)")
