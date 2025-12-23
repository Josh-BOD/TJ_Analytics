"""
Test TrafficJunky API endpoints to discover all available campaign settings
"""
import requests
import json
from datetime import datetime, timedelta

# Your API key
API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
BASE_URL = "https://api.trafficjunky.com/api"

# Get a sample campaign ID first
end_date = datetime.now() - timedelta(days=1)
start_date = end_date - timedelta(days=7)

date_params = {
    'api_key': API_KEY,
    'startDate': start_date.strftime('%d/%m/%Y'),
    'endDate': end_date.strftime('%d/%m/%Y'),
    'limit': 1,
    'offset': 1
}

print("=" * 100)
print("TESTING TRAFFICJUNKY API FOR COMPLETE CAMPAIGN SETTINGS")
print("=" * 100)

# First, get a campaign ID to test with
campaigns_response = requests.get(f"{BASE_URL}/campaigns/bids/stats.json", params=date_params)
campaigns_data = campaigns_response.json()

if isinstance(campaigns_data, dict):
    first_campaign_id = list(campaigns_data.keys())[0]
    first_campaign = campaigns_data[first_campaign_id]
elif isinstance(campaigns_data, list) and campaigns_data:
    first_campaign = campaigns_data[0]
    first_campaign_id = first_campaign.get('campaignId')
else:
    print("No campaigns found!")
    exit()

print(f"\nUsing Campaign ID: {first_campaign_id}")
print(f"Campaign Name: {first_campaign.get('campaignName', 'Unknown')}")

# List of potential endpoints to test for campaign settings
endpoints_to_test = [
    # Campaign detail endpoints
    f"/campaigns/{first_campaign_id}.json",
    f"/campaigns/{first_campaign_id}/settings.json",
    f"/campaigns/{first_campaign_id}/targeting.json",
    f"/campaigns/{first_campaign_id}/bids.json",
    f"/campaigns/{first_campaign_id}/placements.json",
    f"/campaigns/{first_campaign_id}/creatives.json",
    f"/campaigns/{first_campaign_id}/schedule.json",
    f"/campaigns/{first_campaign_id}/time-targeting.json",
    
    # List endpoints
    "/campaigns.json",
    "/campaigns/list.json",
    "/campaigns/all.json",
    "/member/campaigns.json",
    
    # Campaign settings endpoints
    "/campaigns/targeting.json",
    "/campaigns/settings.json",
    
    # Other possible endpoints
    "/spots.json",
    "/placements.json",
    "/countries.json",
    "/time-targets.json",
    "/device-targeting.json",
    "/os-targeting.json",
    "/browser-targeting.json",
]

print("\n" + "=" * 100)
print("TESTING POTENTIAL ENDPOINTS")
print("=" * 100)

results = {}
working_endpoints = []

for endpoint in endpoints_to_test:
    url = BASE_URL + endpoint
    print(f"\nTesting: {url}")
    
    try:
        response = requests.get(url, params={'api_key': API_KEY}, timeout=10)
        status = response.status_code
        
        print(f"  Status: {status}")
        
        if status == 200:
            try:
                data = response.json()
                print(f"  ✓ SUCCESS! Returns JSON data")
                
                # Show structure
                if isinstance(data, dict):
                    keys = list(data.keys())[:10]
                    print(f"  Keys: {keys}")
                    
                    # If dict has nested data, show first item's keys
                    if keys:
                        first_val = data[keys[0]]
                        if isinstance(first_val, dict):
                            print(f"  First item keys: {list(first_val.keys())[:15]}")
                            
                elif isinstance(data, list):
                    print(f"  Items: {len(data)}")
                    if data and isinstance(data[0], dict):
                        print(f"  First item keys: {list(data[0].keys())[:15]}")
                
                results[endpoint] = {
                    'status': 'SUCCESS',
                    'data': data
                }
                working_endpoints.append(endpoint)
                
            except json.JSONDecodeError:
                print(f"  Response is not valid JSON")
                print(f"  Response text: {response.text[:300]}")
                results[endpoint] = {'status': 'NOT JSON', 'response': response.text[:300]}
                
        elif status == 404:
            print(f"  ✗ Endpoint not found (404)")
        elif status == 403:
            print(f"  ✗ Access forbidden (403)")
        elif status == 405:
            print(f"  ✗ Method not allowed (405) - try POST?")
        else:
            print(f"  ✗ Error: {status}")
            print(f"  Response: {response.text[:200]}")
            
    except requests.exceptions.Timeout:
        print(f"  ✗ Timeout")
    except Exception as e:
        print(f"  ✗ Exception: {str(e)}")

print("\n" + "=" * 100)
print("WORKING ENDPOINTS SUMMARY")
print("=" * 100)

if working_endpoints:
    for endpoint in working_endpoints:
        print(f"\n✓ {endpoint}")
        if endpoint in results:
            data = results[endpoint].get('data', {})
            if isinstance(data, dict) and data:
                first_key = list(data.keys())[0]
                first_val = data[first_key]
                if isinstance(first_val, dict):
                    print(f"  Available fields: {', '.join(first_val.keys())}")
else:
    print("No additional endpoints found beyond /campaigns/bids/stats.json")

# Show what we HAVE from the known working endpoint
print("\n" + "=" * 100)
print("CURRENT CAMPAIGN SETTINGS FROM /campaigns/bids/stats.json")
print("=" * 100)

print("\nMain Campaign Fields:")
main_fields = {k: v for k, v in first_campaign.items() if k not in ['bids', 'spots']}
for field, value in main_fields.items():
    print(f"  • {field}: {value}")

print("\nAd Placements (spots):")
for spot in first_campaign.get('spots', [])[:5]:
    print(f"  • {spot.get('name')} (ID: {spot.get('id')})")
if len(first_campaign.get('spots', [])) > 5:
    print(f"  ... and {len(first_campaign.get('spots', [])) - 5} more")

print("\nGeo Targeting (bids):")
countries = set()
for bid in first_campaign.get('bids', []):
    countries.add(bid.get('countryName', 'Unknown'))
print(f"  Countries: {', '.join(countries)}")

print("\nBid Details:")
for bid in first_campaign.get('bids', [])[:3]:
    print(f"  • Placement {bid.get('placementId')}: ${bid.get('bid')} - {bid.get('countryName')}")
if len(first_campaign.get('bids', [])) > 3:
    print(f"  ... and {len(first_campaign.get('bids', [])) - 3} more bids")

# Save results
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
filename = f'campaign_settings_test_{timestamp}.json'
with open(filename, 'w') as f:
    json.dump({
        'working_endpoints': working_endpoints,
        'first_campaign_full_data': first_campaign,
        'results': {k: v for k, v in results.items() if v.get('status') == 'SUCCESS'}
    }, f, indent=2, default=str)

print(f"\n\n✓ Results saved to: {filename}")

print("\n" + "=" * 100)
print("SETTINGS NOT AVAILABLE VIA API (would need web scraping)")
print("=" * 100)
settings_not_available = [
    "Device targeting (iOS vs Android vs Desktop)",
    "OS targeting (specific OS versions)",
    "Browser targeting",
    "Time/day parting schedules (beyond numberOfTimeTargets count)",
    "Frequency capping settings",
    "Campaign start/end dates",
    "Conversion tracking URL/pixels",
    "Creative details (images, videos, URLs)",
    "Category/content targeting",
    "Per-country/per-placement performance breakdown"
]
for setting in settings_not_available:
    print(f"  ✗ {setting}")



