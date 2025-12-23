import requests
import pandas as pd
from datetime import datetime, timedelta

# Your API key
API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"

def get_all_campaign_stats(days_back=30, start_date=None, end_date=None, quiet=False):
    """Get all campaigns with their stats in one API call"""
    url = "https://api.trafficjunky.com/api/campaigns/bids/stats.json"
    
    # Calculate date range - end date must be yesterday or earlier
    if start_date and end_date:
        # Use provided dates
        if isinstance(start_date, str):
            start_date = datetime.strptime(start_date, '%Y-%m-%d')
        if isinstance(end_date, str):
            end_date = datetime.strptime(end_date, '%Y-%m-%d')
    else:
        # Use default logic
        end_date = datetime.now() - timedelta(days=1)  # Yesterday
        start_date = end_date - timedelta(days=days_back)  # 30 days before yesterday
    
    params = {
        'api_key': API_KEY,
        'startDate': start_date.strftime('%d/%m/%Y'),
        'endDate': end_date.strftime('%d/%m/%Y'),
        'limit': 1000,  # Get more results
        'offset': 1
    }
    
    if not quiet:
        print(f"Getting campaign stats from {params['startDate']} to {params['endDate']}")
    
    response = requests.get(url, params=params)
    response.raise_for_status()
    
    return response.json()

def extract_campaign_data(start_date=None, end_date=None, quiet=False):
    """Extract campaign performance and cost data"""
    # Get all campaign stats
    stats_data = get_all_campaign_stats(30, start_date, end_date, quiet)
    
    if not quiet:
        print(f"Total campaigns: {len(stats_data)}")
    
    data = []
    
    # Handle both list and dict responses from API
    if isinstance(stats_data, list):
        # API returned a list of campaigns
        campaigns_to_process = stats_data
    elif isinstance(stats_data, dict):
        # API returned a dict with campaign IDs as keys
        campaigns_to_process = stats_data.values()
    else:
        if not quiet:
            print("Unexpected API response format")
        return pd.DataFrame()
    
    # Loop through each campaign in the response
    for campaign_stats in campaigns_to_process:
        if campaign_stats and isinstance(campaign_stats, dict):
            # Helper function to convert to numeric safely
            def to_numeric(value, default=0):
                try:
                    if value is None or value == '':
                        return default
                    return float(value) if '.' in str(value) else int(value)
                except (ValueError, TypeError):
                    return default
            
            # Use the correct field names from the API with proper type conversion
            record = {
                'campaign_id': campaign_stats.get('campaignId', campaign_stats.get('id', 'unknown')),
                'campaign_name': campaign_stats.get('campaignName', ''),
                'campaign_type': campaign_stats.get('campaignType', ''),
                'status': campaign_stats.get('status', ''),
                'daily_budget': to_numeric(campaign_stats.get('dailyBudget', 0)),
                'daily_budget_left': to_numeric(campaign_stats.get('dailyBudgetLeft', 0)),
                'ads_paused': to_numeric(campaign_stats.get('adsPaused', 0)),
                'number_of_bids': to_numeric(campaign_stats.get('numberOfBids', 0)),
                'number_of_creatives': to_numeric(campaign_stats.get('numberOfCreative', 0)),
                'impressions': to_numeric(campaign_stats.get('impressions', 0)),
                'clicks': to_numeric(campaign_stats.get('clicks', 0)),
                'conversions': to_numeric(campaign_stats.get('conversions', 0)),
                'cost': to_numeric(campaign_stats.get('cost', 0)),  # This is your cost data
                'ctr': to_numeric(campaign_stats.get('CTR', 0)),
                'cpm': to_numeric(campaign_stats.get('CPM', 0)),
            }
            
            data.append(record)
    
    final_df = pd.DataFrame(data)
    
    # Simple summary for consolidator
    if quiet and not final_df.empty:
        # Check if cost column exists
        if 'cost' in final_df.columns:
            total_cost = final_df['cost'].sum()
            print(f"TrafficJunky: {len(final_df)} campaigns, ${total_cost:.2f}")
        else:
            print(f"TrafficJunky: {len(final_df)} campaigns, $0.00")
    
    return final_df

if __name__ == "__main__":
    # TEST: Check what fields the API returns
    print("=" * 80)
    print("TESTING: Checking API response for country/geo data")
    print("=" * 80)
    
    # Get raw API response
    stats_data = get_all_campaign_stats(7, quiet=False)  # Last 7 days
    
    # Show first campaign with ALL fields
    if isinstance(stats_data, list):
        first_campaign = stats_data[0] if stats_data else None
    elif isinstance(stats_data, dict):
        first_campaign = list(stats_data.values())[0] if stats_data else None
    else:
        first_campaign = None
    
    if first_campaign:
        print("\n" + "=" * 80)
        print("FIRST CAMPAIGN - ALL AVAILABLE FIELDS:")
        print("=" * 80)
        import json
        print(json.dumps(first_campaign, indent=2, default=str))
        
        print("\n" + "=" * 80)
        print("FIELD NAMES ONLY:")
        print("=" * 80)
        print(", ".join(sorted(first_campaign.keys())))
        
        # Check for geo/country related fields
        geo_fields = [key for key in first_campaign.keys() if any(
            term in key.lower() for term in ['geo', 'country', 'location', 'region', 'city', 'state']
        )]
        
        print("\n" + "=" * 80)
        print("GEO/COUNTRY RELATED FIELDS:")
        print("=" * 80)
        if geo_fields:
            print(f"✅ FOUND: {', '.join(geo_fields)}")
            for field in geo_fields:
                print(f"   {field}: {first_campaign.get(field)}")
        else:
            print("❌ NO geo/country/location fields found in API response")
    
    print("\n" + "=" * 80)
    print("Now extracting full campaign data...")
    print("=" * 80)
    
    # Extract data
    df = extract_campaign_data()
    
    # Show summary
    print(f"\nSummary:")
    print(f"Total campaigns: {len(df)}")
    print(f"Total cost: ${df['cost'].sum():,.2f}")
    print(f"Total clicks: {df['clicks'].sum():,}")
    print(f"Total impressions: {df['impressions'].sum():,}")
    
    # Save to CSV
    filename = f"campaign_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    df.to_csv(filename, index=False)
    print(f"Data saved to: {filename}")
    
    # Show all campaigns
    print(f"\nCampaign Performance:")
    print(df[['campaign_name', 'campaign_type', 'impressions', 'clicks', 'cost', 'conversions']].to_string(index=False)) 