# Monte Carlo Projection Tool for Jira

A lightweight, local Monte Carlo simulation tool for agile teams using Jira. Produces probabilistic forecasts for backlog completion — with confidence-level date projections — using your team's own historical data exported directly from Jira.

**No cloud. No accounts. No data leaves your machine.**

![Monte Carlo Projection Tool Screenshot](screenshot.png)

---

## Why This Tool?

Most agile forecasting tools either require cloud access, send your data to a third-party service, or are buried inside expensive portfolio management platforms. This tool runs entirely on your local machine using CSV exports from Jira — making it safe to use with sensitive or proprietary project data on corporate networks.

It supports both **Scrum** (velocity-based) and **Kanban** (throughput-based) teams.

---

## What It Does

- Runs 10,000 Monte Carlo simulations (configurable) against your historical Jira data
- Forecasts how many sprints or weeks are needed to clear a backlog
- Shows results at 50%, 70%, 85%, and 95% confidence levels
- Calculates forecast finish **dates** from a given start date
- Displays a weekly throughput breakdown for Kanban teams
- Saves charts as PNG for use in presentations or reports

---

## Requirements

### Python
Python 3.7 or higher is required. Download from [python.org](https://python.org/downloads).

> **Windows users**: During installation, tick **"Add Python to PATH"** before clicking Install.

### Python packages
Three packages are required. Install them with:

```
pip install pandas numpy matplotlib
```

If you are on a corporate network with SSL inspection, use:

```
pip install pandas numpy matplotlib --trusted-host pypi.org --trusted-host files.pythonhosted.org --user
```

`tkinter` (the GUI framework) is **built into Python** — no installation needed.

### Platform support
| Platform | Supported |
|----------|-----------|
| Windows  | ✅ |
| macOS    | ✅ |
| Linux    | ✅ |

---

## Installation

1. Download or clone this repository
2. Install the required packages (see above)
3. Run the tool:

```
python monte_carlo_jira.py
```

No configuration files, no setup scripts, no environment variables.

---

## Exporting Data from Jira

### Scrum teams — Velocity CSV

1. Go to your Jira board
2. Click **Reports** → **Velocity Chart**
3. Click **Export** (top right of the chart)
4. Save the CSV file

The tool will automatically detect the completed points column.

### Kanban teams — Issue Export CSV

1. Go to **Issues** → **Search for Issues** (or use an existing filter)
2. Filter for your team's completed issues over the last 12–16 weeks, for example:
   ```
   project = "YOUR PROJECT" AND statusCategory = Done AND resolved >= -16w ORDER BY resolved ASC
   ```
3. Click **Export** → **Export Excel CSV (all fields)**
4. Save the CSV file

The tool reads the **Resolved** date column and calculates weekly throughput automatically.

### Control Chart CSV (optional — both modes)

1. Go to your Jira board
2. Click **Reports** → **Control Chart**
3. Click **Export**
4. Save the CSV file

If provided, the tool will add cycle time percentile statistics to the summary.

---

## Usage

1. Run `python monte_carlo_jira.py`
2. Select your **Team Type** — Scrum or Kanban
3. Browse to your **CSV file**
4. Set your **backlog size** (story points for Scrum, number of items for Kanban)
5. Set your **start date** (defaults to today)
6. Click **Run Simulation**

Results appear across three tabs:

- **Chart** — distribution histograms with confidence band overlays
- **Summary** — full numeric breakdown including forecast finish dates
- **Weekly Throughput** — week-by-week item count with visual bar (Kanban mode)

Use **Save Chart as PNG** to export the chart for reports or stakeholder presentations.

---

## Understanding the Results

### Confidence levels

The tool reports results at four confidence levels:

| Level | Meaning |
|-------|---------|
| 50% | Half of all simulations finished by this point. Optimistic estimate. |
| 70% | A reasonable working estimate. |
| 85% | A safe commitment for most stakeholder conversations. |
| 95% | Near-certain. Use for hard deadlines or release planning. |

### Weeks/sprints chart (left)

Shows how many weeks or sprints were needed across all simulations. The dashed lines show where each confidence level falls. A wider distribution means more variability in your historical throughput.

### Throughput chart (right)

Shows how many items or points were completed in the median number of periods across all simulations. The `>=` figures are **lower bounds** — in X% of simulations, the team completed *at least* that many items.

### Why the 95% throughput figure is higher than the 50%

The throughput chart asks *"how much will we complete?"* — so higher is better. 95% confidence means the team exceeded that figure in nearly every simulation. This is the opposite direction to the weeks chart, which asks *"how long will it take?"* where lower is better.

---

## Tips

- **Minimum history**: At least 5–6 sprints or weeks of data for meaningful results. 10–16 is better.
- **Zero weeks**: Weeks with zero completions (holidays, blockers) are included in the simulation. If they are genuinely non-working periods you may want to reduce the weeks of history to exclude them, or note their effect on the confidence spread.
- **Mean vs median gap**: A large gap between mean and median throughput indicates spike-and-drought flow patterns — items being held and resolved in batches. This is worth addressing as a flow improvement.
- **Story points vs items**: For Kanban forecasting, item count is generally more reliable than story points because it does not depend on consistent estimation.

---

## Licence

MIT Licence — free to use, adapt, and distribute. See [LICENSE](LICENSE) for details.

---

## Contributing

Issues and pull requests welcome. If your Jira export uses column names not automatically detected by the tool, please open an issue with the column headers and they will be added to the auto-detection logic.
