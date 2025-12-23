import requests
import json
from datetime import datetime, timedelta

# Your API key
API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
BASE_URL = "https://api.trafficjunky.com/api"

print("=" * 100)
print("TESTING FOR CREATIVE DATA ENDPOINTS")
print("=" * 100)

# Calculate dates
end_date = datetime.now() - timedelta(days=1)
start_date = end_date - timedelta(days=2)

date_params = {
    'api_key': API_KEY,
    'startDate': start_date.strftime('%d/%m/%Y'),
    'endDate': end_date.strftime('%d/%m/%Y'),
    'limit': 5,
    'offset': 1
}

# Test different possible endpoints
endpoints_to_test = [
    "/campaigns/bids/stats.json",  # We know this works
    "/creatives/stats.json",
    "/campaigns/creatives/stats.json",
    "/campaigns/creatives.json",
    "/creatives.json",
    "/ads/stats.json",
    "/banners/stats.json",
]

results = []

for endpoint in endpoints_to_test:
    url = BASE_URL + endpoint
    print(f"\nTesting: {url}")
    
    try:
        response = requests.get(url, params=date_params, timeout=10)
        status = response.status_code
        
        print(f"  Status: {status}")
        
        if status == 200:
            try:
                data = response.json()
                print(f"  ✓ SUCCESS! Returns JSON data")
                print(f"  Response type: {type(data)}")
                if isinstance(data, dict):
                    print(f"  Keys: {len(data)} items")
                    if data:
                        first_key = list(data.keys())[0]
                        print(f"  First item keys: {list(data[first_key].keys())[:10]}")
                elif isinstance(data, list):
                    print(f"  Items: {len(data)}")
                    if data:
                        print(f"  First item keys: {list(data[0].keys())[:10]}")
                
                results.append({
                    'endpoint': endpoint,
                    'status': 'SUCCESS',
                    'data_sample': json.dumps(data, indent=2)[:500]
                })
            except:
                print(f"  ✗ Response is not JSON")
                print(f"  Response text: {response.text[:200]}")
                results.append({
                    'endpoint': endpoint,
                    'status': 'NOT JSON',
                    'response': response.text[:200]
                })
        elif status == 404:
            print(f"  ✗ Endpoint not found")
            results.append({
                'endpoint': endpoint,
                'status': '404 Not Found'
            })
        elif status == 403:
            print(f"  ✗ Access forbidden")
            results.append({
                'endpoint': endpoint,
                'status': '403 Forbidden'
            })
        else:
            print(f"  ✗ Error: {status}")
            print(f"  Response: {response.text[:200]}")
            results.append({
                'endpoint': endpoint,
                'status': f'Error {status}',
                'response': response.text[:200]
            })
            
    except requests.exceptions.Timeout:
        print(f"  ✗ Timeout")
        results.append({
            'endpoint': endpoint,
            'status': 'TIMEOUT'
        })
    except Exception as e:
        print(f"  ✗ Exception: {str(e)[:100]}")
        results.append({
            'endpoint': endpoint,
            'status': 'EXCEPTION',
            'error': str(e)[:100]
        })

print("\n" + "=" * 100)
print("SUMMARY")
print("=" * 100)

for result in results:
    print(f"\n{result['endpoint']}")
    print(f"  Status: {result['status']}")
    if result['status'] == 'SUCCESS':
        print("  ✓ This endpoint works!")

# Save results
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
filename = f'creative_endpoints_test_{timestamp}.json'
with open(filename, 'w') as f:
    json.dump(results, f, indent=2)

print(f"\n✓ Results saved to: {filename}")

print("\n" + "=" * 100)
print("CHECKING CURRENT CAMPAIGN RESPONSE FOR CREATIVE DATA")
print("=" * 100)

# Get a campaign and check for creative references
campaign_url = BASE_URL + "/campaigns/bids/stats.json"
response = requests.get(campaign_url, params=date_params)
data = response.json()

if isinstance(data, dict):
    first_campaign = list(data.values())[0]
elif isinstance(data, list):
    first_campaign = data[0]

print("\nFields that might relate to creatives:")
for key, value in first_campaign.items():
    if 'creative' in key.lower() or 'ad' in key.lower() or 'banner' in key.lower():
        print(f"  • {key}: {value}")

print("\n" + "=" * 100)




