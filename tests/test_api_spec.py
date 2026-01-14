import requests
import json

# Fetch the API spec
url = "https://api.trafficjunky.com/docs/api-docs.json"
response = requests.get(url)
spec = response.json()

print("="*70)
print("FULL API SPECIFICATION ANALYSIS")
print("="*70)

# Find bid-related endpoints
print("\n1. All paths containing 'bid':")
for path in spec.get('paths', {}):
    if 'bid' in path.lower():
        print(f"   {path}")

# Look for schemas/definitions related to bids
print("\n2. Schemas/definitions related to bids:")
schemas = spec.get('components', {}).get('schemas', {})
if not schemas:
    schemas = spec.get('definitions', {})

for name, schema in schemas.items():
    if 'bid' in name.lower():
        print(f"\n   === {name} ===")
        print(f"   {json.dumps(schema, indent=4)[:2000]}")

# Print the full /api/bids/{campaignId} endpoint details
print("\n3. /api/bids/{campaignId}.{format} endpoint details:")
bid_campaign_path = spec.get('paths', {}).get('/api/bids/{campaignId}.{format}', {})
if bid_campaign_path:
    print(json.dumps(bid_campaign_path, indent=2)[:3000])
else:
    print("   Path not found, trying other variations...")
    for path, details in spec.get('paths', {}).items():
        if 'campaignId' in path and 'bids' in path:
            print(f"\n   Found: {path}")
            print(json.dumps(details, indent=2)[:3000])
            break

# Check if there's a countryCode in the response
print("\n4. Searching for 'country' in the spec:")
spec_str = json.dumps(spec)
if 'countryCode' in spec_str:
    print("   countryCode FOUND in spec")
if 'country' in spec_str.lower():
    print("   'country' FOUND in spec (case insensitive)")
