import requests
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# 🔹 API Configuration
TBA_AUTH_KEY = "YOUR_API_KEY"  # Replace with your API key
EVENT_KEY = "YOUR _EVENT_KEY"  # Replace with your event key
HEADERS = {"X-TBA-Auth-Key": TBA_AUTH_KEY}

# 🔹 Fetch Data Function
def fetch_data(endpoint):
    url = f"https://www.thebluealliance.com/api/v3/{endpoint}"
    response = requests.get(url, headers=HEADERS)
    
    if response.status_code != 200:
        print(f"⚠️ Error {response.status_code}: {response.text}")  # Debugging info
        return None
    return response.json()

# 🔹 Get Teams Attending
teams = fetch_data(f"event/{EVENT_KEY}/teams/keys") or []
if not teams:
    print("⚠️ No team data available. Exiting.")
    exit()

# 🔹 Initialize Team Stats
team_stats = {team[3:]: {'OPR': 0, 'DPR': 0, 'CCWM': 0, 'RP': 0, 'score_avg': 0, 'matches_played': 0} for team in teams}

# 🔹 Get OPR, DPR, CCWM
event_stats = fetch_data(f"event/{EVENT_KEY}/oprs")
if event_stats:
    for team in teams:
        team_num = team[3:]
        team_stats[team_num]['OPR'] = event_stats.get('oprs', {}).get(team, 0)
        team_stats[team_num]['DPR'] = event_stats.get('dprs', {}).get(team, 0)
        team_stats[team_num]['CCWM'] = event_stats.get('ccwms', {}).get(team, 0)

# 🔹 Get Ranking Points (RP)
rankings = fetch_data(f"event/{EVENT_KEY}/rankings")
if rankings and 'rankings' in rankings:
    for row in rankings['rankings']:
        team_num = str(row.get('team_key', ''))[3:]
        if team_num in team_stats:
            rp_values = row.get('sort_orders', [])
            if len(rp_values) > 0:
                team_stats[team_num]['RP'] = rp_values[0]
            else:
                print(f"⚠️ No RP data for team {team_num}")

# 🔹 Analyze Match Performance
matches = fetch_data(f"event/{EVENT_KEY}/matches") or []
for match in matches:
    if match["comp_level"] != "qm":  # Ignore playoffs, only use qualification matches
        continue

    for alliance in ["blue", "red"]:
        score = match["alliances"][alliance]["score"]
        for team in match["alliances"][alliance]["team_keys"]:
            team_num = team[3:]
            if team_num in team_stats:
                team_stats[team_num]["score_avg"] += score
                team_stats[team_num]["matches_played"] += 1

# 🔹 Compute Power Rankings (Ensemble Model)
power_rankings = []
for team, stats in team_stats.items():
    matches = stats["matches_played"] if stats["matches_played"] > 0 else 1
    score_avg = stats["score_avg"] / matches  # Average match contribution
    
    # Weighted Score of Multiple Metrics
    ranking_score = (
        0.4 * stats["OPR"] +   # Offensive Power
        0.3 * stats["CCWM"] +  # Contribution to Winning
        0.2 * score_avg +      # Actual match performance
        0.1 * stats["RP"]      # Ranking Points
    )

    power_rankings.append((team, ranking_score))

# 🔹 Sort by Power Ranking Score
power_rankings.sort(key=lambda x: x[1], reverse=True)

# 🔹 Extract Top Teams for Graphing
top_teams = [team for team, _ in power_rankings[:10]]
top_scores = [score for _, score in power_rankings[:10]]

# 🔹 Plot Power Rankings
sns.set(style="whitegrid")
plt.figure(figsize=(10, 6))
ax = sns.barplot(x=top_scores, y=top_teams, palette="viridis")

plt.xlabel("Power Ranking Score")
plt.ylabel("Team Number")
plt.title("🏆 FRC Power Rankings - EventName 🏆")
plt.gca().invert_yaxis()  # Highest-ranked team on top
plt.grid(axis="x", linestyle="--", alpha=0.7)

# 🔹 Show Values on Bars
for i, v in enumerate(top_scores):
    ax.text(v + 0.5, i, f"{v:.2f}", va="center", fontsize=12)

plt.show()
