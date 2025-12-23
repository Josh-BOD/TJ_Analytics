import requests
import json
import pandas as pd
from datetime import datetime, timedelta

# Your API key
API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
API_URL = "https://api.trafficjunky.com/api/campaigns/bids/stats.json"

print("=" * 100)
print("EXPORTING TRAFFICJUNKY API DATA STRUCTURE")
print("=" * 100)

# Get last 3 days of data
end_date = datetime.now() - timedelta(days=1)
start_date = end_date - timedelta(days=2)

params = {
    'api_key': API_KEY,
    'startDate': start_date.strftime('%d/%m/%Y'),
    'endDate': end_date.strftime('%d/%m/%Y'),
    'limit': 10,  # Get more campaigns for better sample
    'offset': 1
}

print(f"\nFetching sample data from {params['startDate']} to {params['endDate']}...")

response = requests.get(API_URL, params=params)
data = response.json()

timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

# 1. Export full JSON response
json_filename = f'api_full_response_{timestamp}.json'
with open(json_filename, 'w') as f:
    json.dump(data, f, indent=2)
print(f"✓ Saved full JSON response: {json_filename}")

# 2. Export first campaign structure
if isinstance(data, dict):
    first_campaign = list(data.values())[0]
elif isinstance(data, list):
    first_campaign = data[0]
else:
    print("ERROR: Unexpected data format")
    exit()

sample_filename = f'api_sample_campaign_{timestamp}.json'
with open(sample_filename, 'w') as f:
    json.dump(first_campaign, f, indent=2)
print(f"✓ Saved sample campaign: {sample_filename}")

# 3. Export field structure as CSV
fields_data = []
for key, value in sorted(first_campaign.items()):
    if key not in ['bids', 'spots']:
        fields_data.append({
            'Field Name': key,
            'Sample Value': str(value)[:50],
            'Data Type': type(value).__name__,
            'Description': {
                'campaignId': 'Unique campaign identifier',
                'campaignName': 'Campaign name',
                'campaignType': 'Type of campaign (bid, deal)',
                'status': 'Campaign status',
                'clicks': 'Total clicks',
                'impressions': 'Total impressions',
                'conversions': 'Total conversions',
                'cost': 'Total cost spent',
                'CTR': 'Click-through rate (%)',
                'CPM': 'Cost per thousand impressions',
                'dailyBudget': 'Daily budget limit',
                'dailyBudgetLeft': 'Remaining daily budget',
            }.get(key, '')
        })

fields_df = pd.DataFrame(fields_data)
fields_csv = f'api_field_structure_{timestamp}.csv'
fields_df.to_csv(fields_csv, index=False)
print(f"✓ Saved field structure: {fields_csv}")

# 4. Export country targeting data
if 'bids' in first_campaign and first_campaign['bids']:
    bids_data = []
    for bid in first_campaign['bids']:
        bids_data.append({
            'Placement ID': bid.get('placementId'),
            'Country Code': bid.get('countryCode'),
            'Country Name': bid.get('countryName'),
            'Region Code': bid.get('regionCode'),
            'Region Name': bid.get('regionName'),
            'City': bid.get('city'),
            'Bid Amount': bid.get('bid')
        })
    
    bids_df = pd.DataFrame(bids_data)
    bids_csv = f'api_country_targeting_{timestamp}.csv'
    bids_df.to_csv(bids_csv, index=False)
    print(f"✓ Saved country targeting data: {bids_csv}")

# 5. Export spots/placements data
if 'spots' in first_campaign and first_campaign['spots']:
    spots_data = []
    for spot in first_campaign['spots']:
        spots_data.append({
            'Spot ID': spot.get('id'),
            'Spot Name': spot.get('name')
        })
    
    spots_df = pd.DataFrame(spots_data)
    spots_csv = f'api_ad_placements_{timestamp}.csv'
    spots_df.to_csv(spots_csv, index=False)
    print(f"✓ Saved ad placements: {spots_csv}")

# 6. Export summary report
summary_filename = f'api_data_summary_{timestamp}.txt'
with open(summary_filename, 'w') as f:
    f.write("=" * 100 + "\n")
    f.write("TRAFFICJUNKY API DATA STRUCTURE SUMMARY\n")
    f.write("=" * 100 + "\n\n")
    
    f.write(f"Date Range: {params['startDate']} to {params['endDate']}\n")
    f.write(f"Total Campaigns Returned: {len(data)}\n\n")
    
    f.write("MAIN CAMPAIGN FIELDS:\n")
    f.write("-" * 100 + "\n")
    for key in sorted(first_campaign.keys()):
        if key not in ['bids', 'spots']:
            f.write(f"  • {key}\n")
    
    f.write("\nCOUNTRY/GEO DATA (from 'bids' array):\n")
    f.write("-" * 100 + "\n")
    f.write("  • placementId\n")
    f.write("  • countryCode (e.g., US, AU, UK)\n")
    f.write("  • countryName (e.g., United States)\n")
    f.write("  • regionCode\n")
    f.write("  • regionName\n")
    f.write("  • city\n")
    f.write("  • bid (amount)\n")
    
    f.write("\nAD PLACEMENTS (from 'spots' array):\n")
    f.write("-" * 100 + "\n")
    f.write("  • id (spot ID)\n")
    f.write("  • name (e.g., Pornhub PC - Preroll)\n")
    
    f.write("\n" + "=" * 100 + "\n")
    f.write("KEY FINDINGS\n")
    f.write("=" * 100 + "\n\n")
    
    findings = [
        "✓ Country data IS available in 'bids' array",
        "✓ Each bid contains: countryCode, countryName, regionCode, regionName, city",
        "✗ Performance stats (clicks, impressions, cost) are AGGREGATED (not per-country)",
        "✗ To get per-country performance, you'd need Selenium web scraping",
        "",
        "WHAT THE API PROVIDES:",
        "  • Total impressions, clicks, cost, conversions per CAMPAIGN",
        "  • Which countries the campaign TARGETS",
        "  • Bid amounts per country",
        "",
        "WHAT THE API DOESN'T PROVIDE:",
        "  • Impressions by country",
        "  • Clicks by country",
        "  • Cost by country",
        "  • Conversions by country",
        "",
        "CONCLUSION:",
        "The API provides campaign-level aggregated statistics and shows which",
        "countries are being targeted, but does NOT break down performance metrics",
        "by country. For per-country performance data, web scraping (Selenium)",
        "would be required."
    ]
    
    for finding in findings:
        f.write(finding + "\n")

print(f"✓ Saved summary report: {summary_filename}")

print("\n" + "=" * 100)
print("EXPORT COMPLETE!")
print("=" * 100)
print("\nFiles created:")
print(f"  1. {json_filename} - Full API response (JSON)")
print(f"  2. {sample_filename} - Sample campaign structure (JSON)")
print(f"  3. {fields_csv} - Field definitions (CSV)")
print(f"  4. {bids_csv} - Country targeting data (CSV)")
print(f"  5. {spots_csv} - Ad placements (CSV)")
print(f"  6. {summary_filename} - Summary report (TXT)")
print("\nYou can open the CSV files in Excel/Google Sheets!")
print("=" * 100)




