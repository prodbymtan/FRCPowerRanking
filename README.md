# FRC Power Rankings Generator

## Overview
This script retrieves and analyzes team performance data for an FRC event using **The Blue Alliance (TBA) API**. It calculates power rankings based on a weighted ensemble model incorporating **Offensive Power Rating (OPR), Defensive Power Rating (DPR), Calculated Contribution to Winning Margin (CCWM), Ranking Points (RP), and Average Match Score**.

## Features
- **Fetches team lists** attending the specified event.
- **Retrieves OPR, DPR, and CCWM** statistics.
- **Extracts ranking points (RP)** from event standings.
- **Analyzes qualification matches** to compute average team performance.
- **Calculates a composite power ranking score** using an ensemble model.
- **Sorts teams** by overall performance and displays the top rankings.

## Power Ranking Formula
The script assigns weights to different performance metrics to compute an overall **ranking score**:

\[ \text{Ranking Score} = (0.4 \times \text{OPR}) + (0.3 \times \text{CCWM}) + (0.2 \times \text{Score Avg}) + (0.1 \times \text{RP}) \]

## Installation & Setup
### Prerequisites
- Python 3.x
- `requests` package (install using `pip install requests`)
- A valid **The Blue Alliance API key**

### Setup Instructions
1. Clone the repository or copy the script.
2. Replace `TBA_AUTH_KEY` with your valid **TBA API key**.
3. Set the event key (e.g., `2025mawor`) for the competition you want to analyze.
4. Run the script:
   ```bash
   python script.py
   ```

## Output
The script prints a **sorted list of teams** ranked by performance metrics, displaying the top teams along with their computed ranking scores.

Example Output:
```
🏆 TOP 10 POWER RANKINGS 🏆
1. Team 254 → Score: 89.45
2. Team 1678 → Score: 85.32
3. Team 2056 → Score: 82.67
...
```

## Customization
- Adjust the weight coefficients in the **ranking formula** to modify how different stats contribute to the final rankings.
- Modify the script to include **playoff matches** instead of just qualification rounds.
- Expand the data collection by integrating **additional metrics** from TBA API.

## License
This project is open-source and available for educational and competitive robotics analysis purposes.

---
*Powered by The Blue Alliance API*
