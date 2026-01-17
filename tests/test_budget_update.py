"""
Test script to verify budget UPDATE capability

WARNING: This will make REAL changes to your campaign budget!
Only run this if you understand the implications.
"""

import requests
import json

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
BASE_URL = "https://api.trafficjunky.com/api"

# Use a test campaign ID
CAMPAIGN_ID = "1013205721"

def get_current_budget(campaign_id):
    """Get current budget for a campaign"""
    url = f"{BASE_URL}/campaigns/{campaign_id}.json?api_key={API_KEY}"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        return {
            'campaign_name': data.get('campaign_name'),
            'daily_budget': data.get('campaign_daily_budget'),
            'budget_left': data.get('daily_budget_left')
        }
    return None

def update_budget(campaign_id, new_budget):
    """Update campaign budget"""
    url = f"{BASE_URL}/campaigns/{campaign_id}.json?api_key={API_KEY}"
    
    # Try with dailyBudget parameter
    payload = {
        'dailyBudget': str(new_budget)
    }
    
    print(f"\nAttempting PUT to: {url.replace(API_KEY, 'HIDDEN')}")
    print(f"Payload: {payload}")
    
    response = requests.put(url, json=payload)
    
    print(f"Response status: {response.status_code}")
    print(f"Response headers: {dict(response.headers)}")
    print(f"Response body: {response.text[:500]}")
    
    return response.status_code, response.text

# =============================================================================
# Main Test
# =============================================================================
print("=" * 70)
print("BUDGET UPDATE TEST")
print("=" * 70)

# Step 1: Get current budget
print("\n1. Getting current budget...")
current = get_current_budget(CAMPAIGN_ID)
if current:
    print(f"   Campaign: {current['campaign_name']}")
    print(f"   Current daily budget: ${current['daily_budget']}")
    print(f"   Budget left: ${current['budget_left']}")
else:
    print("   Failed to get current budget")
    exit(1)

# Step 2: Test update (we'll set it to the SAME value to be safe)
print("\n2. Testing budget update (setting to same value for safety)...")
print(f"   Setting budget to: ${current['daily_budget']}")

status, response_text = update_budget(CAMPAIGN_ID, current['daily_budget'])

if status == 200:
    print("\n   ✓ SUCCESS! Budget update endpoint works!")
    
    # Verify the change
    print("\n3. Verifying budget after update...")
    after = get_current_budget(CAMPAIGN_ID)
    if after:
        print(f"   Daily budget after: ${after['daily_budget']}")
else:
    print(f"\n   ✗ Update failed with status {status}")
    
    # Try alternative methods
    print("\n   Trying alternative payload format...")
    
    url = f"{BASE_URL}/campaigns/{CAMPAIGN_ID}.json?api_key={API_KEY}&dailyBudget={current['daily_budget']}"
    response = requests.put(url)
    print(f"   URL params method - Status: {response.status_code}")
    print(f"   Response: {response.text[:300]}")

print("\n" + "=" * 70)
print("TEST COMPLETE")
print("=" * 70)
