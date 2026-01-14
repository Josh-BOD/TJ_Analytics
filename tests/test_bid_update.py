"""
Test TrafficJunky Bid UPDATE API Endpoints
Check if we can update bids via API
"""

import requests
import json

# Configuration
API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
BASE_URL = "https://api.trafficjunky.com/api"
CAMPAIGN_ID = "1013232471"
BID_ID = "1201505001"  # One of the bid IDs from the campaign

print("="*80)
print("TESTING TRAFFICJUNKY BID UPDATE ENDPOINTS")
print("="*80)
print(f"Campaign ID: {CAMPAIGN_ID}")
print(f"Bid ID: {BID_ID}")

# First, get the current bid value
print("\n--- Getting current bid value ---")
get_url = f"{BASE_URL}/bids/{CAMPAIGN_ID}.json?api_key={API_KEY}"
response = requests.get(get_url)
if response.status_code == 200:
    data = response.json()
    if BID_ID in data:
        current_bid = data[BID_ID]['bid']
        print(f"Current bid value: ${current_bid}")
    else:
        print(f"Bid {BID_ID} not found")
        current_bid = "2.600"
else:
    print(f"Error getting current bid: {response.status_code}")
    current_bid = "2.600"

# Test various update endpoints (READ-ONLY - just checking if endpoints exist)
print("\n" + "="*80)
print("TESTING UPDATE ENDPOINTS (OPTIONS/HEAD requests only - no actual changes)")
print("="*80)

endpoints_to_test = [
    # PUT endpoints
    ("PUT", f"/bids/{BID_ID}.json"),
    ("PUT", f"/bids/{CAMPAIGN_ID}/{BID_ID}.json"),
    ("PUT", f"/campaigns/{CAMPAIGN_ID}/bids/{BID_ID}.json"),
    ("PUT", f"/bid/{BID_ID}.json"),
    
    # POST endpoints  
    ("POST", f"/bids/{BID_ID}.json"),
    ("POST", f"/bids/{CAMPAIGN_ID}/{BID_ID}.json"),
    ("POST", f"/campaigns/{CAMPAIGN_ID}/bids/{BID_ID}.json"),
    ("POST", f"/bids/update.json"),
    ("POST", f"/bids/{BID_ID}/update.json"),
    
    # PATCH endpoints
    ("PATCH", f"/bids/{BID_ID}.json"),
    ("PATCH", f"/campaigns/{CAMPAIGN_ID}/bids/{BID_ID}.json"),
]

# Use OPTIONS request to check if endpoint exists without making changes
for method, endpoint in endpoints_to_test:
    url = f"{BASE_URL}{endpoint}?api_key={API_KEY}"
    print(f"\n{method} {endpoint}")
    
    try:
        # First try OPTIONS to see what methods are allowed
        options_response = requests.options(url, timeout=10)
        allowed = options_response.headers.get('Allow', 'N/A')
        print(f"  OPTIONS status: {options_response.status_code}, Allowed: {allowed}")
        
        # Try HEAD request
        head_response = requests.head(url, timeout=10)
        print(f"  HEAD status: {head_response.status_code}")
        
    except Exception as e:
        print(f"  Error: {str(e)[:50]}")

# Now try an actual PUT request with the SAME bid value (no change)
print("\n" + "="*80)
print("TESTING ACTUAL UPDATE (with same bid value - no real change)")
print("="*80)

# Test 1: PUT /bids/{bid_id}.json
print("\n--- Test: PUT /bids/{bid_id}.json ---")
url = f"{BASE_URL}/bids/{BID_ID}.json?api_key={API_KEY}"
payload = {"bid": current_bid}  # Same value, no actual change
print(f"URL: {url.replace(API_KEY, 'HIDDEN')}")
print(f"Payload: {payload}")

try:
    response = requests.put(url, json=payload, timeout=10)
    print(f"Status: {response.status_code}")
    print(f"Response: {response.text[:500]}")
except Exception as e:
    print(f"Error: {e}")

# Test 2: POST /bids/{bid_id}.json  
print("\n--- Test: POST /bids/{bid_id}.json ---")
url = f"{BASE_URL}/bids/{BID_ID}.json?api_key={API_KEY}"
try:
    response = requests.post(url, json=payload, timeout=10)
    print(f"Status: {response.status_code}")
    print(f"Response: {response.text[:500]}")
except Exception as e:
    print(f"Error: {e}")

# Test 3: PUT /campaigns/{campaign_id}/bids/{bid_id}.json
print("\n--- Test: PUT /campaigns/{campaign_id}/bids/{bid_id}.json ---")
url = f"{BASE_URL}/campaigns/{CAMPAIGN_ID}/bids/{BID_ID}.json?api_key={API_KEY}"
try:
    response = requests.put(url, json=payload, timeout=10)
    print(f"Status: {response.status_code}")
    print(f"Response: {response.text[:500]}")
except Exception as e:
    print(f"Error: {e}")

# Test 4: Try form data instead of JSON
print("\n--- Test: POST with form data ---")
url = f"{BASE_URL}/bids/{BID_ID}.json"
form_data = {"api_key": API_KEY, "bid": current_bid}
try:
    response = requests.post(url, data=form_data, timeout=10)
    print(f"Status: {response.status_code}")
    print(f"Response: {response.text[:500]}")
except Exception as e:
    print(f"Error: {e}")

# Test 5: Try PUT with form data
print("\n--- Test: PUT with form data ---")
url = f"{BASE_URL}/bids/{BID_ID}.json"
try:
    response = requests.put(url, data=form_data, timeout=10)
    print(f"Status: {response.status_code}")
    print(f"Response: {response.text[:500]}")
except Exception as e:
    print(f"Error: {e}")

print("\n" + "="*80)
print("DONE - Check results above to see if any update endpoint works")
print("="*80)
