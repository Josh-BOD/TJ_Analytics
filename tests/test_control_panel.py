"""
TJ Control Panel - Python Test Script

This script validates the TrafficJunky API calls for multi-period stats
(Today, Yesterday, 7-Day) before building the Google Apps Script.

Usage:
    python test_control_panel.py
"""

import requests
from datetime import datetime, timedelta
import pytz
import json
from collections import defaultdict


def to_float(value, default=0.0):
    """Safely convert value to float"""
    if value is None or value == '':
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default

# ============================================================================
# CONFIGURATION
# ============================================================================

API_KEY = "9c77fe112485aff1fc1266f21994de5fabcf8df5f45886c0504494bc4ea0479156092332fc020799a6183b154b467f2bb3e2b169e4886504d83e63fe08c4f039"
API_BASE_URL = "https://api.trafficjunky.com/api"
API_TIMEZONE = "America/New_York"

# Test with these campaign IDs (from Legend sheet)
TEST_CAMPAIGN_IDS = ["1013232471"]

# ============================================================================
# DATE HELPERS
# ============================================================================

def get_est_now():
    """Get current datetime in EST timezone"""
    est = pytz.timezone(API_TIMEZONE)
    return datetime.now(est)

def format_date_for_api(date):
    """Format date as DD/MM/YYYY for TrafficJunky API"""
    return date.strftime("%d/%m/%Y")

def get_date_ranges():
    """
    Calculate date ranges for Today, Yesterday, and 7-Day periods.
    All dates are in EST timezone.
    """
    now = get_est_now()
    today = now.date()
    yesterday = today - timedelta(days=1)
    seven_days_ago = today - timedelta(days=7)
    
    return {
        'today': {
            'start': format_date_for_api(today),
            'end': format_date_for_api(today),
            'label': 'Today'
        },
        'yesterday': {
            'start': format_date_for_api(yesterday),
            'end': format_date_for_api(yesterday),
            'label': 'Yesterday'
        },
        'seven_day': {
            'start': format_date_for_api(seven_days_ago),
            'end': format_date_for_api(yesterday),  # Excludes today
            'label': '7-Day (excl. today)'
        }
    }

# ============================================================================
# API FETCH FUNCTIONS
# ============================================================================

def fetch_current_bids(campaign_ids):
    """
    Fetch current bid values and metadata from /api/bids/{campaignId}.json
    
    Returns:
        dict: {bid_id: {bid_data}} for all bids across all campaigns
    """
    all_bids = {}
    
    for campaign_id in campaign_ids:
        print(f"\n📥 Fetching bids for campaign {campaign_id}...")
        
        # First get campaign name
        campaign_url = f"{API_BASE_URL}/campaigns/{campaign_id}.json?api_key={API_KEY}"
        try:
            resp = requests.get(campaign_url)
            if resp.status_code == 200:
                campaign_data = resp.json()
                campaign_name = campaign_data.get('campaign_name', '')
                print(f"   Campaign: {campaign_name}")
            else:
                campaign_name = ''
                print(f"   ⚠️ Could not fetch campaign name: {resp.status_code}")
        except Exception as e:
            campaign_name = ''
            print(f"   ⚠️ Error fetching campaign: {e}")
        
        # Now get bids
        bids_url = f"{API_BASE_URL}/bids/{campaign_id}.json?api_key={API_KEY}"
        try:
            resp = requests.get(bids_url)
            if resp.status_code != 200:
                print(f"   ❌ Error {resp.status_code}: {resp.text[:200]}")
                continue
            
            data = resp.json()
            
            # Response is object with bid_ids as keys
            if isinstance(data, dict):
                bid_count = len(data)
                print(f"   ✅ Found {bid_count} bids")
                
                for bid_id, bid_data in data.items():
                    # Add campaign info to bid data
                    bid_data['campaign_id'] = campaign_id
                    bid_data['campaign_name'] = campaign_name
                    all_bids[bid_id] = bid_data
                    
        except Exception as e:
            print(f"   ❌ Exception: {e}")
    
    return all_bids


def fetch_bid_stats_for_period(campaign_ids, start_date, end_date, period_label):
    """
    Fetch BID-LEVEL stats for a specific date range.
    
    Uses /api/bids/{campaignId}.json with startDate/endDate parameters.
    This endpoint supports date filtering at the bid level!
    
    Returns:
        dict: {bid_id: {stats}} for all bids
    """
    print(f"\n📊 Fetching BID stats for {period_label} ({start_date} to {end_date})...")
    
    all_stats = {}
    
    for campaign_id in campaign_ids:
        url = f"{API_BASE_URL}/bids/{campaign_id}.json"
        params = {
            'api_key': API_KEY,
            'startDate': start_date,
            'endDate': end_date
        }
        
        try:
            resp = requests.get(url, params=params)
            
            if resp.status_code != 200:
                print(f"   ⚠️ Campaign {campaign_id}: {resp.status_code}")
                continue
            
            data = resp.json()
            
            if isinstance(data, dict):
                for bid_id, bid_data in data.items():
                    stats = bid_data.get('stats', {})
                    all_stats[bid_id] = {
                        'impressions': to_float(stats.get('impressions', 0)),
                        'clicks': to_float(stats.get('clicks', 0)),
                        'conversions': to_float(stats.get('conversions', 0)),
                        'cost': to_float(stats.get('revenue', 0)),  # API calls it 'revenue'
                        'ecpm': to_float(stats.get('ecpm', 0)),
                        'ctr': to_float(stats.get('ctr', 0)),
                    }
                    
        except Exception as e:
            print(f"   ❌ Campaign {campaign_id} error: {e}")
    
    print(f"   ✅ Got stats for {len(all_stats)} bids")
    return all_stats


# ============================================================================
# DATA PROCESSING
# ============================================================================

def extract_device_os(spot_name, campaign_name):
    """
    Extract device and OS from spot name and campaign name.
    
    Device: from spot_name (Mobile, Tablet, PC/Desktop)
    OS: from campaign_name (_AND_, _IOS_, etc.)
    
    Returns: "Mob - iOS", "Tab - Android", "Desk - All", etc.
    """
    # Device from spot name
    device = 'Desk'
    if 'Mobile' in spot_name:
        device = 'Mob'
    elif 'Tablet' in spot_name:
        device = 'Tab'
    
    # OS from campaign name
    os_type = 'All'
    campaign_upper = campaign_name.upper()
    if '_IOS_' in campaign_upper or '_IOS' in campaign_upper or '-IOS_' in campaign_upper or '-IOS-' in campaign_upper:
        os_type = 'iOS'
    elif '_AND_' in campaign_upper or '_AND' in campaign_upper or '-AND_' in campaign_upper or '-AND-' in campaign_upper:
        os_type = 'Android'
    
    return f"{device} - {os_type}"


def extract_countries(geos):
    """
    Extract country codes from geos object.
    
    Returns: comma-separated country codes or count if too many
    """
    if not geos:
        return ''
    
    countries = []
    for geo_key, geo_data in geos.items():
        if isinstance(geo_data, dict):
            cc = geo_data.get('countryCode', '')
            if cc and cc not in countries:
                countries.append(cc)
    
    if len(countries) <= 5:
        return ', '.join(countries)
    else:
        return f"{', '.join(countries[:3])} (+{len(countries) - 3} more)"


def merge_bid_data(current_bids, today_stats, yesterday_stats, seven_day_stats):
    """
    Merge all data sources into the final row structure.
    
    - current_bids: keyed by bid_id (bid-level detail)
    - today/yesterday/seven_day_stats: keyed by bid_id (BID-level stats with date filtering!)
    
    Returns: list of dicts, each representing a row in the Control Panel
    """
    rows = []
    
    for bid_id, bid_data in current_bids.items():
        # Get BID-level stats for each period (keyed by bid_id)
        t_stats = today_stats.get(bid_id, {})
        y_stats = yesterday_stats.get(bid_id, {})
        sd_stats = seven_day_stats.get(bid_id, {})
        
        # Extract geo info
        geos = bid_data.get('geos', {})
        geo_ids = list(geos.keys()) if geos else []
        
        # Calculate 7D daily averages (total / 7)
        sd_spend = sd_stats.get('cost', 0)
        sd_conv = sd_stats.get('conversions', 0)
        sd_impr = sd_stats.get('impressions', 0)
        sd_clicks = sd_stats.get('clicks', 0)
        
        sd_spend_avg = sd_spend / 7 if sd_spend else 0
        sd_conv_avg = sd_conv / 7 if sd_conv else 0
        
        # Calculate CPA (Cost per Acquisition)
        t_conv = t_stats.get('conversions', 0)
        t_spend = t_stats.get('cost', 0)
        y_conv = y_stats.get('conversions', 0)
        y_spend = y_stats.get('cost', 0)
        
        t_cpa = t_spend / t_conv if t_conv > 0 else 0
        y_cpa = y_spend / y_conv if y_conv > 0 else 0
        sd_cpa = sd_spend / sd_conv if sd_conv > 0 else 0
        
        row = {
            # A-G: Identity columns
            'tier1_strategy': '',  # From Legend sheet
            'sub_strategy': '',    # From Legend sheet
            'campaign_name': bid_data.get('campaign_name', ''),
            'campaign_id': bid_data.get('campaign_id', ''),
            'format': '',          # From Legend sheet
            'country': extract_countries(geos),
            'device_os': extract_device_os(
                bid_data.get('spot_name', ''),
                bid_data.get('campaign_name', '')
            ),
            
            # H-L: Bid management columns
            'current_ecpm_bid': to_float(bid_data.get('bid', 0)),
            'new_cpm_bid': '',     # User editable
            'change_pct': '',      # Formula
            't_bid_adjust': '',    # Formula - check Bid Logs
            'date_last_bid_adjust': '',  # Formula - VLOOKUP Bid Logs
            
            # M-O: eCPM (Today, Yesterday, 7D avg)
            't_ecpm': t_stats.get('ecpm', 0),
            'y_ecpm': y_stats.get('ecpm', 0),
            '7d_ecpm': sd_stats.get('ecpm', 0),  # API returns avg already
            
            # P-R: Spend (Today, Yesterday, 7D daily avg)
            't_spend': t_spend,
            'y_spend': y_spend,
            '7d_spend': sd_spend_avg,
            
            # S-U: CPA (Today, Yesterday, 7D)
            't_cpa': t_cpa,
            'y_cpa': y_cpa,
            '7d_cpa': sd_cpa,
            
            # V-X: Conversions (Today, Yesterday, 7D daily avg)
            't_conv': t_conv,
            'y_conv': y_conv,
            '7d_conv': sd_conv_avg,
            
            # Y-AA: CTR (Today, Yesterday, 7D avg)
            't_ctr': t_stats.get('ctr', 0),
            'y_ctr': y_stats.get('ctr', 0),
            '7d_ctr': sd_stats.get('ctr', 0),  # API returns avg already
            
            # AB-AE: Reference IDs
            'spot_id': bid_data.get('spot_id', ''),
            'bid_id': bid_id,
            'geo_id': geo_ids[0] if len(geo_ids) == 1 else f"{len(geo_ids)} geos" if geo_ids else '',
            'last_updated': datetime.now().isoformat(),
        }
        
        rows.append(row)
    
    return rows


# ============================================================================
# MAIN
# ============================================================================

def main():
    print("=" * 60)
    print("TJ Control Panel - API Test Script")
    print("=" * 60)
    
    # Get date ranges
    date_ranges = get_date_ranges()
    print(f"\n📅 Date Ranges (EST):")
    for period, dates in date_ranges.items():
        print(f"   {dates['label']}: {dates['start']} to {dates['end']}")
    
    # Step 1: Fetch current bids
    print("\n" + "=" * 60)
    print("STEP 1: Fetch Current Bids")
    print("=" * 60)
    current_bids = fetch_current_bids(TEST_CAMPAIGN_IDS)
    print(f"\n📋 Total bids fetched: {len(current_bids)}")
    
    if current_bids:
        # Show sample bid structure
        sample_bid_id = list(current_bids.keys())[0]
        sample_bid = current_bids[sample_bid_id]
        print(f"\n📝 Sample bid structure (bid_id: {sample_bid_id}):")
        for key, value in sample_bid.items():
            if key == 'geos':
                print(f"   {key}: {len(value)} geo targets")
            elif key == 'stats':
                print(f"   {key}: {value}")
            else:
                print(f"   {key}: {value}")
    
    # Step 2: Fetch BID-LEVEL stats for each period
    # (The /api/bids/{campaignId}.json endpoint DOES support date filtering!)
    print("\n" + "=" * 60)
    print("STEP 2: Fetch BID Stats (T/Y/7D)")
    print("=" * 60)
    
    today_stats = fetch_bid_stats_for_period(
        TEST_CAMPAIGN_IDS,
        date_ranges['today']['start'],
        date_ranges['today']['end'],
        date_ranges['today']['label']
    )
    
    yesterday_stats = fetch_bid_stats_for_period(
        TEST_CAMPAIGN_IDS,
        date_ranges['yesterday']['start'],
        date_ranges['yesterday']['end'],
        date_ranges['yesterday']['label']
    )
    
    seven_day_stats = fetch_bid_stats_for_period(
        TEST_CAMPAIGN_IDS,
        date_ranges['seven_day']['start'],
        date_ranges['seven_day']['end'],
        date_ranges['seven_day']['label']
    )
    
    # Step 3: Merge data
    print("\n" + "=" * 60)
    print("STEP 3: Merge Data")
    print("=" * 60)
    
    rows = merge_bid_data(current_bids, today_stats, yesterday_stats, seven_day_stats)
    print(f"\n📊 Merged {len(rows)} rows")
    
    # Step 4: Display sample output
    print("\n" + "=" * 60)
    print("STEP 4: Sample Output")
    print("=" * 60)
    
    if rows:
        print("\n📋 Column Headers:")
        headers = list(rows[0].keys())
        for i, h in enumerate(headers):
            print(f"   {chr(65 + i)}: {h}")
        
        print(f"\n📋 First {min(3, len(rows))} rows:")
        for i, row in enumerate(rows[:3]):
            print(f"\n--- Row {i + 1} ---")
            print(f"   Campaign: {row['campaign_name']}")
            print(f"   Device/OS: {row['device_os']}")
            print(f"   Country: {row['country']}")
            print(f"   Current Bid: ${row['current_ecpm_bid']}")
            print(f"   T eCPM: ${row['t_ecpm']:.3f}, Y eCPM: ${row['y_ecpm']:.3f}, 7D eCPM: ${row['7d_ecpm']:.3f}")
            print(f"   T Spend: ${row['t_spend']:.2f}, Y Spend: ${row['y_spend']:.2f}, 7D Spend (avg): ${row['7d_spend']:.2f}")
            print(f"   T Conv: {row['t_conv']}, Y Conv: {row['y_conv']}, 7D Conv (avg): {row['7d_conv']:.1f}")
            print(f"   Spot ID: {row['spot_id']}, Bid ID: {row['bid_id']}")
    
    # Save to JSON for reference
    output_file = "control_panel_test_output.json"
    with open(output_file, 'w') as f:
        json.dump({
            'date_ranges': date_ranges,
            'total_bids': len(current_bids),
            'rows': rows
        }, f, indent=2, default=str)
    print(f"\n💾 Full output saved to: {output_file}")
    
    print("\n" + "=" * 60)
    print("✅ Test Complete")
    print("=" * 60)


if __name__ == "__main__":
    main()
