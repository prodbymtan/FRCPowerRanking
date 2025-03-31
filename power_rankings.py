import requests
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import os
import time  # Added import for sleep functionality
import gspread  # Added for Google Sheets integration
from oauth2client.service_account import ServiceAccountCredentials  # Added for Google Sheets auth

# Enable interactive mode for matplotlib
plt.ion()

# 🔹 API Configuration
TBA_AUTH_KEY = "UiaVF0OcJcT3Sp8eKr1uExmiwYMqeMey8DmudPL4AgHfROhBY9fiHNi55FaoECfD"  # Replace with your TBA API key
EVENT_KEY = "2025vtbur"
CURRENT_SEASON = 2025  # Update for each season
HEADERS = {"X-TBA-Auth-Key": TBA_AUTH_KEY}

# 🔹 Google Sheets Configuration
# Use a service account with appropriate permissions
# Save your credentials JSON file in the same directory or specify path
SHEETS_CREDENTIALS_FILE = 'frc-scouting-credentials.json'  # Update with your file name
SPREADSHEET_NAME = f'FRC {CURRENT_SEASON} Power Rankings - {EVENT_KEY}'  # Name of your spreadsheet

# 🔹 Connect to Google Sheets
def connect_to_sheets():
    try:
        # Define the scope
        scope = ['https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive']
        
        # Authenticate using service account credentials
        credentials = ServiceAccountCredentials.from_json_keyfile_name(SHEETS_CREDENTIALS_FILE, scope)
        
        # Authorize the client
        client = gspread.authorize(credentials)
        
        # Try to open existing spreadsheet, create if it doesn't exist
        try:
            spreadsheet = client.open(SPREADSHEET_NAME)
            print(f"✅ Connected to existing Google Sheet: {SPREADSHEET_NAME}")
        except gspread.SpreadsheetNotFound:
            spreadsheet = client.create(SPREADSHEET_NAME)
            # Make it accessible to anyone with the link
            spreadsheet.share(None, perm_type='anyone', role='reader')
            print(f"✅ Created new Google Sheet: {SPREADSHEET_NAME}")
            print(f"📊 Link: {spreadsheet.url}")
            
            # Initialize the sheets
            rankings_sheet = spreadsheet.add_worksheet(title="Power Rankings", rows=100, cols=20)
            detailed_sheet = spreadsheet.add_worksheet(title="Detailed Metrics", rows=100, cols=20)
            historical_sheet = spreadsheet.add_worksheet(title="Historical Data", rows=100, cols=20)
            
            # Delete default Sheet1
            try:
                sheet1 = spreadsheet.worksheet("Sheet1")
                spreadsheet.del_worksheet(sheet1)
            except:
                pass
            
        return spreadsheet
        
    except Exception as e:
        print(f"⚠️ Error connecting to Google Sheets: {e}")
        return None

# 🔹 Update Google Sheets with Power Rankings
def update_google_sheets(spreadsheet, power_rankings):
    if not spreadsheet:
        print("⚠️ No spreadsheet connection available. Skipping Google Sheets update.")
        return
    
    try:
        # Update main Power Rankings sheet
        rankings_sheet = spreadsheet.worksheet("Power Rankings")
        
        # Clear existing data
        rankings_sheet.clear()
        
        # Prepare header and data rows for main rankings
        header = ["Rank", "Team", "Score", "OPR", "DPR", "CCWM", "Avg Score", "Win Rate", 
                  "Avg Auto", "Avg Barge", "Matches Played", "Last Updated"]
        
        # Current timestamp for update tracking
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Create rows for the rankings data
        rows = [header]
        for rank, (team, score, stats) in enumerate(power_rankings, 1):
            matches = stats["matches_played"] if stats["matches_played"] > 0 else 1
            match_avg = stats["score_avg"] / matches
            
            row = [
                rank,
                f"Team {team}",
                round(score, 2),
                round(stats.get('OPR', 0), 2),
                round(stats.get('DPR', 0), 2),
                round(stats.get('CCWM', 0), 2),
                round(match_avg, 2),
                f"{stats.get('win_rate', 0)*100:.1f}%",
                round(stats.get('avg_auto', 0), 2),
                round(stats.get('avg_barge', 0), 2),
                stats['matches_played'],
                timestamp
            ]
            rows.append(row)
        
        # Update the sheet with all data at once (more efficient)
        rankings_sheet.update('A1', rows)
        
        # Format the header row
        rankings_sheet.format('A1:L1', {
            'backgroundColor': {'red': 0.2, 'green': 0.2, 'blue': 0.2},
            'textFormat': {'bold': True, 'foregroundColor': {'red': 1.0, 'green': 1.0, 'blue': 1.0}}
        })
        
        # Auto-resize columns to fit content
        try:
            for i in range(1, len(header) + 1):
                rankings_sheet.columns_auto_resize(i, i)
        except:
            pass  # Auto-resize is not critical, continue if it fails
        
        print(f"✅ Google Sheet updated at {timestamp}")
        print(f"📊 Sheet URL: {spreadsheet.url}")
        
    except Exception as e:
        print(f"⚠️ Error updating Google Sheets: {e}")

# 🔹 Fetch Data with Error Handling
def fetch_data(endpoint):
    url = f"https://www.thebluealliance.com/api/v3/{endpoint}"
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"⚠️ API Error: {e}")
        return None

# 🔹 Get Teams Attending the Event
def get_event_teams():
    teams = fetch_data(f"event/{EVENT_KEY}/teams/keys") or []
    if not teams:
        print("⚠️ No team data available. Exiting.")
        exit()
    return teams

# 🔹 Get Historical Data for a Team
def get_team_history(team_key):
    # Get team events from current season
    season_events = fetch_data(f"team/{team_key}/events/{CURRENT_SEASON}") or []
    
    # Sort events by date (most recent last)
    season_events.sort(key=lambda e: e.get('start_date', ''))
    
    # Filter out future events and the current event
    past_events = [e for e in season_events if e.get('key') != EVENT_KEY and 
                  datetime.fromisoformat(e.get('end_date', '2099-01-01')) < datetime.now()]
    
    if not past_events:
        return None
    
    # Track metrics across events
    metrics = {
        'OPR': [], 'DPR': [], 'CCWM': [], 'RP': [], 
        'events': [], 'rank': [], 'record': [], 
        'avg_auto': [], 'avg_barge': [],
        'total_teams': []
    }
    
    for event in past_events:
        event_key = event.get('key')
        event_name = event.get('name', event_key)
        metrics['events'].append(event_name)
        
        # Get OPR/DPR/CCWM
        event_stats = fetch_data(f"event/{event_key}/oprs")
        if event_stats:
            metrics['OPR'].append(event_stats.get('oprs', {}).get(team_key, 0))
            metrics['DPR'].append(event_stats.get('dprs', {}).get(team_key, 0))
            metrics['CCWM'].append(event_stats.get('ccwms', {}).get(team_key, 0))
        else:
            metrics['OPR'].append(0)
            metrics['DPR'].append(0)
            metrics['CCWM'].append(0)
        
        # Get Rankings and Record
        rankings = fetch_data(f"event/{event_key}/rankings")
        team_rank = None
        team_record = {"wins": 0, "losses": 0, "ties": 0}
        rp = 0
        
        if rankings and 'rankings' in rankings:
            metrics['total_teams'].append(len(rankings['rankings']))
            for row in rankings['rankings']:
                if row.get('team_key') == team_key:
                    team_rank = row.get('rank')
                    rp_values = row.get('sort_orders', [])
                    rp = rp_values[0] if rp_values else 0
                    
                    # Get team record
                    record = row.get('record', {})
                    if record:
                        team_record = {
                            "wins": record.get('wins', 0),
                            "losses": record.get('losses', 0),
                            "ties": record.get('ties', 0)
                        }
                    break
            
            metrics['rank'].append(team_rank if team_rank else len(rankings['rankings']))
            metrics['RP'].append(rp)
            metrics['record'].append(team_record)
        else:
            metrics['rank'].append(999)  # Default high rank if no data
            metrics['RP'].append(0)
            metrics['record'].append(team_record)
            metrics['total_teams'].append(40)  # Default reasonable team count
        
        # Get Match Data for Auto and Barge Points
        matches = fetch_data(f"event/{event_key}/matches")
        
        total_auto_points = 0
        total_barge_points = 0
        match_count = 0
        
        # Use specific field names based on user input
        auto_field_name = "autoPoints"       # Default auto field name
        barge_field_name = "endGameBargePoints"  # Specific barge field name the user wants
        
        if matches:
            # Check if endGameBargePoints exists in the data
            field_check_completed = False
            
            for match in matches[:5]:
                if not field_check_completed and match.get('comp_level') == 'qm' and 'score_breakdown' in match and match['score_breakdown']:
                    for color in ['red', 'blue']:
                        if color in match['score_breakdown']:
                            breakdown_keys = list(match['score_breakdown'][color].keys())
                            print(f"    Available fields in {event_key}: {', '.join(breakdown_keys)}")
                            
                            # Check if endGameBargePoints exists
                            if barge_field_name in breakdown_keys:
                                print(f"    Found '{barge_field_name}' in match data")
                            else:
                                print(f"    '{barge_field_name}' not found, checking for alternatives...")
                                # Look for other barge-related fields
                                barge_fields = [k for k in breakdown_keys if 'barge' in k.lower()]
                                if barge_fields:
                                    barge_field_name = barge_fields[0]
                                    print(f"    Using alternative barge field: '{barge_field_name}'")
                                else:
                                    # If no barge fields, use endgame fields as fallback
                                    endgame_fields = [k for k in breakdown_keys if 'endgame' in k.lower()]
                                    if endgame_fields:
                                        barge_field_name = endgame_fields[0]
                                        print(f"    Using fallback field: '{barge_field_name}'")
                                    else:
                                        print(f"    No barge or endgame fields found")
                            
                            field_check_completed = True
                            break
                    if field_check_completed:
                        break
                        
            # Process all matches with the identified field names
            for match in matches:
                if match.get('comp_level') != 'qm':  # Only qualification matches
                    continue
                
                # Find which alliance the team is on
                alliance = None
                for color in ['red', 'blue']:
                    if team_key in match.get('alliances', {}).get(color, {}).get('team_keys', []):
                        alliance = color
                        break
                
                if not alliance or 'score_breakdown' not in match or not match['score_breakdown']:
                    continue
                
                # Extract auto and barge points from match breakdown
                if alliance in match['score_breakdown']:
                    match_count += 1
                    
                    # Get auto points
                    auto_points = match['score_breakdown'][alliance].get(auto_field_name, 0)
                    total_auto_points += auto_points
                    
                    # Get barge points - numerical value if available
                    barge_points = match['score_breakdown'][alliance].get(barge_field_name, 0)
                    
                    # Handle case where barge field might be a boolean (like bargeBonusAchieved)
                    if isinstance(barge_points, bool) or (isinstance(barge_points, str) and barge_points.lower() in ['true', 'false', 'yes', 'no']):
                        # Convert boolean to points (typically bonus points for achieving something)
                        if barge_points in [True, 'True', 'true', 'YES', 'Yes', 'yes']:
                            barge_points = 5  # Assume 5 points for a bonus/achievement
                        else:
                            barge_points = 0
                    
                    total_barge_points += barge_points
        
        # Print debug info about field detection
        if match_count > 0:
            print(f"    Found {match_count} matches for team {team_key[3:]} at {event_key}")
            print(f"    Auto field used: '{auto_field_name}'")
            print(f"    Barge field used: '{barge_field_name}'")
        
        # Calculate averages
        avg_auto = total_auto_points / max(match_count, 1)
        avg_barge = total_barge_points / max(match_count, 1)
        
        metrics['avg_auto'].append(avg_auto)
        metrics['avg_barge'].append(avg_barge)
    
    return metrics

# 🔹 Calculate historical ranking score
def calc_historical_score(history):
    if not history or not history.get('OPR'):
        return 0
    
    # Get the metrics
    oprs = history.get('OPR', [])
    dprs = history.get('DPR', [])
    ccwms = history.get('CCWM', [])
    ranks = history.get('rank', [])
    total_teams = history.get('total_teams', [])
    records = history.get('record', [])
    avg_auto = history.get('avg_auto', [])  # Changed from auto_success_rate
    avg_barge = history.get('avg_barge', [])  # Changed from endgame_success_rate
    
    # Weight metrics (more weight to recent events)
    event_count = len(oprs)
    weights = np.linspace(0.5, 1.0, event_count)
    
    # Calculate normalized rank percentiles (lower rank = higher percentile)
    rank_percentiles = []
    for i, (rank, team_count) in enumerate(zip(ranks, total_teams)):
        if rank and team_count:
            # Convert rank to percentile (1st place = 1.0, last place = 0.0)
            percentile = 1.0 - ((rank - 1) / max(team_count - 1, 1))
            rank_percentiles.append(percentile)
        else:
            rank_percentiles.append(0.5)  # Middle of the pack if no data
    
    # Calculate win rates from records
    win_rates = []
    for record in records:
        total_matches = record.get('wins', 0) + record.get('losses', 0) + record.get('ties', 0)
        if total_matches > 0:
            # Give half credit for ties
            win_rate = (record.get('wins', 0) + 0.5 * record.get('ties', 0)) / total_matches
            win_rates.append(win_rate)
        else:
            win_rates.append(0.5)  # Default 50% win rate if no data
    
    # Calculate weighted averages
    avg_opr = np.average(oprs, weights=weights) if oprs else 0
    avg_dpr = np.average(dprs, weights=weights) if dprs else 0
    avg_ccwm = np.average(ccwms, weights=weights) if ccwms else 0
    avg_rank_pct = np.average(rank_percentiles, weights=weights) if rank_percentiles else 0.5
    avg_win_rate = np.average(win_rates, weights=weights) if win_rates else 0.5
    avg_auto_pts = np.average(avg_auto, weights=weights) if avg_auto else 0  # Changed name
    avg_barge_pts = np.average(avg_barge, weights=weights) if avg_barge else 0  # Changed name
    
    # Normalize auto and barge points to a 0-100 scale for consistent weighting
    # Assuming reasonable max values for scaling
    norm_auto = min(avg_auto_pts / 15 * 100, 100)  # Scale assuming 15 is a "perfect" auto score
    norm_barge = min(avg_barge_pts / 15 * 100, 100)  # Scale assuming 15 is a "perfect" barge score
    
    # Improved historical score formula with redistributed weights
    historical_score = (
        0.35 * avg_opr +          # Offensive capability
        0.15 * (- avg_dpr) +  # Defensive capability (inverted: lower DPR is better)
        0.20 * avg_ccwm +         # Overall contribution to winning margin
        0.10 * avg_rank_pct +     # Team's percentile ranking at events
        0.10 * avg_win_rate +     # Win rate across all matches
        0.05 * norm_auto +        # Average auto points (normalized to 0-100)
        0.05 * norm_barge         # Average barge/endgame points (normalized to 0-100)
    )
    
    return historical_score

# 🔹 Main function to run the entire process
def generate_power_rankings():
    # Get teams
    teams = get_event_teams()
    team_stats = {team[3:]: {
        'OPR': 0, 'DPR': 0, 'CCWM': 0, 'RP': 0, 
        'score_avg': 0, 'matches_played': 0,
        'historical_score': 0, 'rank': 0, 
        'win_rate': 0.5, 'avg_auto': 0, 'avg_barge': 0
    } for team in teams}
    
    print("🔄 Fetching historical team data...")
    # Get historical rankings for each team
    for team in teams:
        print(f"  Processing historical data for {team}...")
        history = get_team_history(team)
        if history:
            team_num = team[3:]
            historical_score = calc_historical_score(history)
            team_stats[team_num]['historical_score'] = historical_score
            
            # Store the last event metrics as defaults if current event has no data
            if history['OPR']:
                team_stats[team_num]['OPR_default'] = history['OPR'][-1]
            if history['CCWM']:
                team_stats[team_num]['CCWM_default'] = history['CCWM'][-1]
            if history['DPR']:
                team_stats[team_num]['DPR_default'] = history['DPR'][-1]
            if history['rank']:
                team_stats[team_num]['rank_default'] = history['rank'][-1]
            if history['avg_auto']:
                team_stats[team_num]['avg_auto_default'] = history['avg_auto'][-1]
            if history['avg_barge']:
                team_stats[team_num]['avg_barge_default'] = history['avg_barge'][-1]
            if history['record'] and history['record'][-1]:
                rec = history['record'][-1]
                matches = rec.get('wins', 0) + rec.get('losses', 0) + rec.get('ties', 0)
                if matches > 0:
                    team_stats[team_num]['win_rate_default'] = (rec.get('wins', 0) + 0.5 * rec.get('ties', 0)) / matches
    
    # 🔹 Get Current Event Stats
    print("🔄 Processing current event data...")
    
    # Get OPR, DPR, CCWM
    event_stats = fetch_data(f"event/{EVENT_KEY}/oprs")
    if event_stats:
        for team in teams:
            team_num = team[3:]
            team_stats[team_num]['OPR'] = event_stats.get('oprs', {}).get(team, 0) or \
                                         team_stats[team_num].get('OPR_default', 0)
            team_stats[team_num]['DPR'] = event_stats.get('dprs', {}).get(team, 0) or \
                                         team_stats[team_num].get('DPR_default', 0)
            team_stats[team_num]['CCWM'] = event_stats.get('ccwms', {}).get(team, 0) or \
                                          team_stats[team_num].get('CCWM_default', 0)
    
    # Get Ranking Points (RP) and Ranks from Standings
    team_count = 0
    rankings = fetch_data(f"event/{EVENT_KEY}/rankings")
    if rankings and 'rankings' in rankings:
        team_count = len(rankings['rankings'])
        for row in rankings['rankings']:
            team_key = row.get('team_key', '')
            team_num = str(team_key)[3:]
            if team_num in team_stats:
                team_stats[team_num]['rank'] = row.get('rank', 999)
                rp_values = row.get('sort_orders', [])
                if len(rp_values) > 0:
                    team_stats[team_num]['RP'] = rp_values[0]
                
                # Extract record
                record = row.get('record', {})
                if record:
                    wins = record.get('wins', 0)
                    losses = record.get('losses', 0)
                    ties = record.get('ties', 0)
                    total = wins + losses + ties
                    if total > 0:
                        team_stats[team_num]['win_rate'] = (wins + 0.5 * ties) / total
    
    # 🔹 Analyze Match Performance including Auto and Barge
    print("🔄 Analyzing match performance with auto and barge points...")
    matches = fetch_data(f"event/{EVENT_KEY}/matches") or []
    analyze_match_performance(matches, team_stats, teams)
    
    # 🔹 Compute Final Power Rankings using Enhanced Ensemble Algorithm
    power_rankings = []
    for team, stats in team_stats.items():
        matches = stats["matches_played"] if stats["matches_played"] > 0 else 1
        score_avg = stats["score_avg"] / matches  # Average match contribution
        
        # Calculate rank percentile (1st place = 1.0, last place = 0.0)
        rank_percentile = 0.5  # Default middle rank
        if stats["rank"] > 0 and team_count > 0:
            rank_percentile = 1.0 - ((stats["rank"] - 1) / max(team_count - 1, 1))
        
        # Normalize auto and barge to 0-100 scale
        norm_auto = min(stats["avg_auto"] / 15 * 100, 100)  # Scale assuming 15 is a "perfect" auto score
        norm_barge = min(stats["avg_barge"] / 15 * 100, 100)  # Scale assuming 15 is a "perfect" barge score
        
        # Enhanced Current Score calculation with the new metrics
        current_score = (
            0.45 * stats["OPR"] +         # Heavy focus on offense
            0.05 * (100 - stats["DPR"]) + # Less concern for defense
            0.20 * stats["CCWM"] +        # Still considers contribution to winning
            0.10 * score_avg +            # Match consistency
            0.05 * rank_percentile +      # Rank consideration
            0.10 * stats["win_rate"] * 100 + # Team's history of winning
            0.025 * norm_auto +           # Autonomous contribution
            0.025 * norm_barge            # Endgame contribution
        )
        
        # Blend current score with historical (if matches exist, rely more on current)
        # More sophisticated blending based on match count and data quality
        match_confidence = min(stats["matches_played"] / 8, 1.0)  # Max confidence after 8 matches
        
        # If we have good current data (4+ matches), weight it more heavily
        if stats["matches_played"] >= 4:
            blend_factor = 0.7 + (0.2 * match_confidence)  # 70-90% weight on current data
        else:
            blend_factor = 0.3 + (0.4 * match_confidence)  # 30-70% weight on current data
        
        historical_factor = 1 - blend_factor  # Historical weight
        
        # Final score calculation with improved blending
        ranking_score = (blend_factor * current_score) + (historical_factor * stats["historical_score"])
        
        power_rankings.append((team, ranking_score, stats))
    
    # 🔹 Sort by Power Ranking Score
    power_rankings.sort(key=lambda x: x[1], reverse=True)
    
    return power_rankings

# 🔹 Visualization Function
def create_ranking_graph(power_rankings, top_n=10):
    # Close all existing figures to prevent memory leaks and clutter
    plt.close('all')
    
    # Create figure with two subplots
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 12), gridspec_kw={'height_ratios': [2, 1]})
    
    # Get the top N teams
    top_teams = power_rankings[:top_n]
    
    # Extract team numbers and scores
    teams = [f"Team {team}" for team, _, _ in top_teams]
    scores = [score for _, score, _ in top_teams]
    
    # Create color gradient (green to yellow to red)
    colors = plt.cm.RdYlGn(np.linspace(0.8, 0.2, len(teams)))
    
    # Create the main bar chart
    bars = ax1.bar(teams, scores, color=colors)
    
    # Add score labels on top of bars
    for bar, score in zip(bars, scores):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                f'{score:.1f}', ha='center', fontweight='bold')
    
    # Add labels and title to main chart
    ax1.set_xlabel('Team', fontsize=12)
    ax1.set_ylabel('Power Ranking Score', fontsize=12)
    ax1.set_title(f'FRC {CURRENT_SEASON} Top {top_n} Power Rankings - UVM', fontsize=14)
    ax1.tick_params(axis='x', rotation=45)
    
    # Create stacked bar chart for score components
    components = []
    # Updated component names for auto and barge
    component_names = ['OPR', 'DPR', 'CCWM', 'Match Avg', 'Rank', 'Auto Pts', 'Barge Pts']
    component_colors = ['#ff9999', '#66b3ff', '#99ff99', '#ffcc99', '#c2c2f0', '#ffb3e6', '#c4e17f']
    
    # Extract component data
    for _, _, stats in top_teams:
        matches = stats["matches_played"] if stats["matches_played"] > 0 else 1
        score_avg = stats["score_avg"] / matches
        
        # Normalize values to make them comparable on same scale
        opr_norm = stats.get("OPR", 0) / 3  # Normalize to approximate 0-33 range
        dpr_norm = (100 - stats.get("DPR", 0)) / 3  # Inverse and normalize
        ccwm_norm = stats.get("CCWM", 0) / 3  # Normalize
        score_norm = score_avg / 4  # Normalize
        
        # Get rank percentile
        team_count = 40  # Reasonable default
        rank = stats.get("rank", team_count)
        if rank > 0:
            rank_norm = 30 * (1.0 - ((rank - 1) / max(team_count - 1, 1)))  # 0-30 scale
        else:
            rank_norm = 15  # Middle value
            
        # Auto and barge points (0-15 scale)
        auto_norm = stats.get("avg_auto", 0)  # Changed from auto_success
        barge_norm = stats.get("avg_barge", 0)  # Changed from endgame_success
        
        components.append([opr_norm, dpr_norm, ccwm_norm, score_norm, rank_norm, auto_norm, barge_norm])
    
    # Create a stacked bar for each team
    bottom = np.zeros(len(teams))
    for i, component_values in enumerate(zip(*components)):
        ax2.bar(teams, component_values, bottom=bottom, label=component_names[i], color=component_colors[i])
        bottom += np.array(component_values)
    
    # Add legend and labels to component chart
    ax2.set_xlabel('Team', fontsize=12)
    ax2.set_ylabel('Score Components', fontsize=12)
    ax2.set_title('Breakdown of Ranking Score Components', fontsize=14)
    ax2.tick_params(axis='x', rotation=45)
    ax2.legend(loc='upper right')
    
    plt.tight_layout()
    
    # Save figure
    save_dir = 'rankings'
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)
    plt.savefig(f'{save_dir}/power_rankings_{EVENT_KEY}.png', dpi=300)
    print(f"✅ Graph saved to {save_dir}/power_rankings_{EVENT_KEY}.png")
    
    # Also create a detailed CSV file with all metrics
    csv_file = f'{save_dir}/power_rankings_{EVENT_KEY}_detailed.csv'
    with open(csv_file, 'w') as f:
        # Updated header with avg_auto and avg_barge
        f.write("Rank,Team,Overall Score,OPR,DPR,CCWM,Match Avg,Win Rate,Avg Auto,Avg Barge,Historical Score,Matches Played\n")
        
        # Write data for all teams
        for rank, (team, score, stats) in enumerate(power_rankings, 1):
            matches = stats["matches_played"] if stats["matches_played"] > 0 else 1
            match_avg = stats["score_avg"] / matches
            
            f.write(f"{rank},{team},{score:.2f},{stats.get('OPR', 0):.2f},{stats.get('DPR', 0):.2f}," + 
                    f"{stats.get('CCWM', 0):.2f},{match_avg:.2f}," +
                    f"{stats.get('win_rate', 0)*100:.1f}%,{stats.get('avg_auto', 0):.2f}," +
                    f"{stats.get('avg_barge', 0):.2f},{stats.get('historical_score', 0):.2f}," +
                    f"{stats['matches_played']}\n")
    
    print(f"✅ Detailed CSV saved to {csv_file}")
    
    # Update the plot display without blocking
    plt.draw()
    plt.pause(0.001)  # Small pause to allow the GUI to update

# 🔹 Analyze Match Performance including Auto and Barge
def analyze_match_performance(matches, team_stats, teams):
    if not matches:
        return
    
    # Set the exact field name for barge points as specified by the user
    barge_field_name = "endGameBargePoints"
    auto_field_name = "autoPoints"  # Default field name
    
    print(f"🔍 Using '{barge_field_name}' as specified for Barge points")
    
    # Look for field names in the first few matches to confirm they exist
    for match in matches[:5]:
        if match.get('comp_level') != 'qm':
            continue
            
        if 'score_breakdown' in match and match['score_breakdown']:
            # Check in each alliance
            for color in ['red', 'blue']:
                if color in match['score_breakdown']:
                    # Get all fields in the score breakdown
                    breakdown_keys = list(match['score_breakdown'][color].keys())
                    
                    # Print all available keys for debugging
                    print(f"🔍 Available score breakdown fields: {', '.join(breakdown_keys)}")
                    
                    # Confirm the barge_field_name exists in the data
                    if barge_field_name in breakdown_keys:
                        print(f"✓ Confirmed '{barge_field_name}' exists in match data")
                    else:
                        print(f"⚠️ Warning: '{barge_field_name}' not found in match data! Will check for alternatives.")
                        # If the specific field doesn't exist, look for alternatives
                        barge_specific = [k for k in breakdown_keys if 'barge' in k.lower()]
                        if barge_specific:
                            barge_field_name = barge_specific[0]
                            print(f"✓ Using alternative barge field: '{barge_field_name}'")
                        else:
                            # If no barge fields exist, use endgame points as fallback
                            endgame_fields = [k for k in breakdown_keys if 'endgame' in k.lower()]
                            if endgame_fields:
                                barge_field_name = endgame_fields[0]
                                print(f"✓ Using fallback endgame field: '{barge_field_name}'")
                    
                    # Look for auto-related fields
                    auto_candidates = [k for k in breakdown_keys if 'auto' in k.lower()]
                    if auto_candidates:
                        if 'autoPoints' in auto_candidates:
                            auto_field_name = 'autoPoints'
                        else:
                            auto_field_name = auto_candidates[0]
                        print(f"✓ Using '{auto_field_name}' for Auto points")
                    
                    # If we've checked the field names, break out
                    break
            
            # If we've confirmed field names, break out
            break
    
    # Extract a single match breakdown to analyze manually
    print("🔄 Examining a sample match score breakdown:")
    for match in matches[:1]:
        if 'score_breakdown' in match and match['score_breakdown']:
            alliance = 'red' if 'red' in match['score_breakdown'] else 'blue'
            if alliance:
                bd = match['score_breakdown'][alliance]
                print(f"Sample match data for alliance {alliance}:")
                for key, value in bd.items():
                    if 'barge' in key.lower() or 'endgame' in key.lower() or 'auto' in key.lower():
                        print(f"  {key}: {value}  <-- Relevant scoring field")
                    else:
                        print(f"  {key}: {value}")
    
    # Now process all the matches with the identified field names
    print(f"🔄 Processing matches using auto_field='{auto_field_name}' and barge_field='{barge_field_name}'")
    
    for match in matches:
        if match["comp_level"] != "qm":  # Ignore playoffs, only use qualification matches
            continue
        
        # Skip matches without score breakdown
        if 'score_breakdown' not in match or not match['score_breakdown']:
            continue
        
        for alliance in ["blue", "red"]:
            score = match["alliances"][alliance]["score"]
            
            # Extract auto and barge data if available
            auto_points = 0
            barge_points = 0
            
            if alliance in match['score_breakdown']:
                # Extract actual auto points using detected field name
                auto_points = match['score_breakdown'][alliance].get(auto_field_name, 0)
                
                # Extract actual barge points using the specified field name
                barge_points = match['score_breakdown'][alliance].get(barge_field_name, 0)
                
                # If value is not a number but a string status, try to convert
                if isinstance(barge_points, str):
                    try:
                        if barge_points.lower() in ['yes', 'docked', 'engaged']:
                            barge_points = 10  # Typical high score for successful endgame
                        elif barge_points.lower() in ['partial', 'parked']:
                            barge_points = 5   # Typical medium score
                        else:
                            barge_points = 0
                    except:
                        barge_points = 0
            
            for team_key in match["alliances"][alliance]["team_keys"]:
                team_num = team_key[3:]
                if team_num in team_stats:
                    # Track total score
                    team_stats[team_num]["score_avg"] += score
                    team_stats[team_num]["matches_played"] += 1
                    
                    # Update average auto points
                    current_matches = team_stats[team_num]["matches_played"]
                    prev_total_auto = team_stats[team_num]["avg_auto"] * (current_matches - 1)
                    team_stats[team_num]["avg_auto"] = (prev_total_auto + auto_points) / current_matches
                    
                    # Update average barge/endgame points
                    prev_total_barge = team_stats[team_num]["avg_barge"] * (current_matches - 1)
                    team_stats[team_num]["avg_barge"] = (prev_total_barge + barge_points) / current_matches
    
    # Print summary of metrics after processing
    print("✅ Match analysis complete")
    # Sample teams to verify data
    for team_num in list(team_stats.keys())[:5]:
        if team_stats[team_num]["matches_played"] > 0:
            print(f"  Team {team_num}: {team_stats[team_num]['matches_played']} matches, " +
                  f"Auto avg: {team_stats[team_num]['avg_auto']:.1f}, " +
                  f"Barge avg: {team_stats[team_num]['avg_barge']:.1f}")
    
    return

# 🔹 Run the program
if __name__ == "__main__":
    print("🚀 Starting FRC Power Rankings Loop (updates every 3 minutes)...")
    
    # Connect to Google Sheets
    spreadsheet = connect_to_sheets()
    
    # Counter for determining when to create visualizations
    update_counter = 0
    
    try:
        while True:
            update_counter += 1
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            print(f"\n⏱️ Update #{update_counter} started at {current_time}")
            
            # Generate rankings
            rankings = generate_power_rankings()
            
            # Print just the top 32 teams for quick reference
            print("\n🏆 TOP 32 POWER RANKINGS 🏆")
            for rank, (team, score, stats) in enumerate(rankings[:32], 1):
                matches = stats["matches_played"] if stats["matches_played"] > 0 else 1
                print(f"{rank}. Team {team} → Score: {score:.2f} | OPR: {stats['OPR']:.1f} | Matches: {stats['matches_played']}")
            
            # Update Google Sheets
            print("\n📊 Updating Google Sheets...")
            update_google_sheets(spreadsheet, rankings)
            
            # Create visualization and show (won't block due to interactive mode)
            print("\n📊 Creating visualization and CSV...")
            create_ranking_graph(rankings, top_n=10)
            
            # Every 5 updates, print full rankings
            if update_counter % 5 == 0:
                print("\n🏆 FULL POWER RANKINGS 🏆")
                for rank, (team, score, stats) in enumerate(rankings[:38], 1):
                    matches = stats["matches_played"] if stats["matches_played"] > 0 else 1
                    match_avg = stats["score_avg"] / matches
                    
                    print(f"{rank}. Team {team} → Score: {score:.2f}")
                    print(f"   OPR: {stats['OPR']:.1f} | DPR: {stats['DPR']:.1f} | CCWM: {stats['CCWM']:.1f}")
                    print(f"   Rank: {stats['rank']} | Win Rate: {stats['win_rate']*100:.1f}% | Match Avg: {match_avg:.1f}")
                    print(f"   Avg Auto: {stats['avg_auto']:.1f} | Avg Barge: {stats['avg_barge']:.1f}")
                    print(f"   Historical Score: {stats['historical_score']:.1f} | Matches: {stats['matches_played']}")
                    print("-" * 50)
            
            print(f"✅ Update completed at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"⏳ Waiting for 3 minutes before next update...")
            
            # Sleep for 3 minutes (180 seconds)
            time.sleep(180)
            
    except KeyboardInterrupt:
        print("\n🛑 Program stopped by user. Exiting...")
    except Exception as e:
        print(f"\n❌ Error occurred: {e}")
        raise
