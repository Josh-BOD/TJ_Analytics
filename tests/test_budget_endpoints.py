"""
Test script to explore TrafficJunky API budget endpoints

This script will:
1. Fetch the API spec to find campaign/budget related endpoints
2. Get current campaign data including budget fields
3. Test potential budget update endpoints
"""

import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
BASE_URL = "https://api.trafficjunky.com/api"

# Use a test campaign ID - replace with a real one
CAMPAIGN_ID = "1013205721"

def print_section(title):
    print("\n" + "=" * 70)
    print(title)
    print("=" * 70)

# =============================================================================
# PART 1: Check API Spec for campaign/budget endpoints
# =============================================================================
print_section("PART 1: API SPEC - Campaign/Budget Endpoints")

try:
    spec_url = "https://api.trafficjunky.com/docs/api-docs.json"
    spec_response = requests.get(spec_url)
    
    if spec_response.status_code == 200:
        spec = spec_response.json()
        
        print("\nEndpoints containing 'campaign':")
        for path in spec.get('paths', {}):
            if 'campaign' in path.lower():
                methods = list(spec['paths'][path].keys())
                print(f"  {path}")
                print(f"    Methods: {methods}")
        
        print("\nEndpoints containing 'budget':")
        for path in spec.get('paths', {}):
            if 'budget' in path.lower():
                methods = list(spec['paths'][path].keys())
                print(f"  {path}")
                print(f"    Methods: {methods}")
        
        print("\nEndpoints containing 'settings':")
        for path in spec.get('paths', {}):
            if 'setting' in path.lower():
                methods = list(spec['paths'][path].keys())
                print(f"  {path}")
                print(f"    Methods: {methods}")
                
        # Look for PUT/POST methods on campaign endpoints
        print("\nPUT/POST endpoints for campaigns:")
        for path, details in spec.get('paths', {}).items():
            if 'campaign' in path.lower():
                for method in ['put', 'post', 'patch']:
                    if method in details:
                        print(f"  {method.upper()} {path}")
                        # Show parameters if available
                        params = details[method].get('parameters', [])
                        if params:
                            print(f"    Parameters: {[p.get('name') for p in params]}")
    else:
        print(f"Failed to fetch API spec: {spec_response.status_code}")
except Exception as e:
    print(f"Error fetching API spec: {e}")

# =============================================================================
# PART 2: Get current campaign data with budget fields
# =============================================================================
print_section("PART 2: Current Campaign Data (Budget Fields)")

# Try campaigns/stats endpoint
url = f"{BASE_URL}/campaigns/stats.json?api_key={API_KEY}&limit=10&offset=1"
print(f"\nFetching from /campaigns/stats.json...")

response = requests.get(url)
if response.status_code == 200:
    data = response.json()
    
    if isinstance(data, list) and len(data) > 0:
        # Show budget-related fields from first campaign
        campaign = data[0]
        print(f"\nFirst campaign fields:")
        
        budget_fields = ['dailyBudget', 'dailyBudgetLeft', 'budget', 'totalBudget', 
                        'daily_budget', 'daily_budget_left', 'campaignBudget']
        
        for field in budget_fields:
            if field in campaign:
                print(f"  {field}: {campaign[field]}")
        
        print(f"\nAll fields in campaign response:")
        for key in sorted(campaign.keys()):
            print(f"  {key}: {campaign[key]}")
else:
    print(f"Error: {response.status_code} - {response.text[:200]}")

# Try single campaign endpoint
print(f"\nFetching single campaign {CAMPAIGN_ID}...")
url2 = f"{BASE_URL}/campaigns/{CAMPAIGN_ID}.json?api_key={API_KEY}"
response2 = requests.get(url2)

if response2.status_code == 200:
    campaign_data = response2.json()
    print(f"\nSingle campaign response fields:")
    for key in sorted(campaign_data.keys()):
        value = campaign_data[key]
        if isinstance(value, (dict, list)) and len(str(value)) > 100:
            print(f"  {key}: <{type(value).__name__} with {len(value)} items>")
        else:
            print(f"  {key}: {value}")
else:
    print(f"Error: {response2.status_code} - {response2.text[:200]}")

# =============================================================================
# PART 3: Test potential budget update endpoints (READ ONLY - just checking)
# =============================================================================
print_section("PART 3: Testing Potential Update Endpoints (OPTIONS/HEAD only)")

# These are potential endpoints that MIGHT support budget updates
# We'll test with OPTIONS request first to see what methods are allowed

potential_endpoints = [
    f"/campaigns/{CAMPAIGN_ID}.json",
    f"/campaigns/{CAMPAIGN_ID}/settings.json",
    f"/campaigns/{CAMPAIGN_ID}/budget.json",
    f"/campaigns/{CAMPAIGN_ID}/set.json",
    f"/campaigns/{CAMPAIGN_ID}/update.json",
]

for endpoint in potential_endpoints:
    url = f"{BASE_URL}{endpoint}?api_key={API_KEY}"
    print(f"\n{endpoint}")
    
    # Try OPTIONS to see allowed methods
    try:
        options_resp = requests.options(url)
        allow_header = options_resp.headers.get('Allow', 'Not specified')
        print(f"  OPTIONS status: {options_resp.status_code}")
        print(f"  Allowed methods: {allow_header}")
    except Exception as e:
        print(f"  OPTIONS error: {e}")
    
    # Try GET to see if endpoint exists
    try:
        get_resp = requests.get(url)
        print(f"  GET status: {get_resp.status_code}")
        if get_resp.status_code == 200:
            data = get_resp.json()
            if isinstance(data, dict):
                print(f"  GET returns: {list(data.keys())[:10]}")
    except Exception as e:
        print(f"  GET error: {e}")

# =============================================================================
# PART 4: Check campaigns/bids/stats for budget info
# =============================================================================
print_section("PART 4: Check campaigns/bids/stats for budget")

url = f"{BASE_URL}/campaigns/bids/stats.json?api_key={API_KEY}&limit=5&offset=1"
print(f"\nFetching from /campaigns/bids/stats.json...")

response = requests.get(url)
if response.status_code == 200:
    data = response.json()
    
    if isinstance(data, list) and len(data) > 0:
        campaign = data[0]
        print(f"\nCampaign-level fields:")
        for key in sorted(campaign.keys()):
            if key != 'bids':
                print(f"  {key}: {campaign[key]}")
        
        # Check if there's budget info
        if 'dailyBudget' in campaign or 'budget' in campaign:
            print("\n  ✓ Budget data IS available in this endpoint!")
        else:
            print("\n  ✗ No budget data in this endpoint")
else:
    print(f"Error: {response.status_code}")

# =============================================================================
# SUMMARY
# =============================================================================
print_section("SUMMARY")
print("""
Based on the tests above:

1. READING BUDGET:
   - Check if dailyBudget/dailyBudgetLeft fields are available
   - Which endpoint provides this data

2. UPDATING BUDGET:
   - Check if any PUT/POST endpoints exist for campaigns
   - Check what methods are allowed on campaign endpoints

Review the output above to determine:
- Can we read budget data? (likely YES)
- Can we update budget via API? (need to verify)
""")
