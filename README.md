# Monte Carlo Projection Tool for Jira

A lightweight, local Monte Carlo simulation tool for agile teams using Jira. Produces probabilistic forecasts for backlog completion — with confidence-level date projections — using your team's own historical data exported directly from Jira.

**No cloud. No accounts. No data leaves your machine.**

![Monte Carlo Projection Tool Screenshot](simulation with dates.png)

---

## Why This Tool?

Most agile forecasting tools either require cloud access, send your data to a third-party service, or are buried inside expensive portfolio management platforms. This tool runs entirely on your local machine using CSV exports from Jira — making it safe to use with sensitive or proprietary project data on corporate networks.

It supports both **Scrum** (velocity-based) and **Kanban** (throughput-based) teams.

---

## What It Does

- Runs 10,000 Monte Carlo simulations (configurable) against your historical Jira data
- Supports **Scrum** (velocity/story points) and **Kanban** (throughput/items) modes
- Forecasts how many sprints or weeks are needed to clear a backlog at multiple confidence levels
- Applies a **capacity adjustment** to account for meetings, holidays, and overhead
- Calculates forecast finish **dates** from a given start date
- Automatically calculates **cycle time** from Created and Resolved dates
- Displays a weekly throughput or sprint velocity breakdown
- Exports results as a **PNG chart**, **HTML report**, or **PDF report**
- Includes **info tooltips** on every field explaining what to enter and why

---

## Platform Support

| Platform | Supported |
|----------|-----------|
| Windows  | ✅ |
| macOS    | ✅ |
| Linux    | ✅ |

---

## Running the Tool

### Windows — Run from Source

1. Install Python 3.7 or higher from [python.org](https://python.org/downloads)

> During installation, tick **"Add Python to PATH"** before clicking Install.

2. Install the required packages:

```
pip install pandas numpy matplotlib
```

If you are on a corporate network with SSL inspection, use:

```
pip install pandas numpy matplotlib --trusted-host pypi.org --trusted-host files.pythonhosted.org --user
```

3. Run the tool:

```
python monte_carlo_jira.py
```

### macOS — Launcher Script

1. Download `monte_carlo_jira.py` and `launch_monte_carlo.command` into the same folder
2. Install dependencies (one-time setup):

```
brew install python@3.12 python-tk@3.12
/opt/homebrew/bin/python3.12 -m venv ~/monte_carlo_env
source ~/monte_carlo_env/bin/activate
pip install pandas numpy matplotlib
```

3. Make the launcher executable (one-time setup):

```
chmod +x /path/to/launch_monte_carlo.command
```

4. Double-click `launch_monte_carlo.command` in Finder to launch the app

> If Homebrew is not installed, run this first:
> ```
> /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
> ```

### Linux — Run from Source

```
pip3 install pandas numpy matplotlib
python3 monte_carlo_jira.py
```

---

## Exporting Data from Jira

### Kanban mode — Issue Export CSV

1. Go to **Issues** → **Search for Issues** (or use an existing filter)
2. Filter for your team's completed issues over the last 12–16 weeks, for example:
   ```
   project = "YOUR PROJECT" AND statusCategory = Done AND resolved >= -16w ORDER BY resolved ASC
   ```
3. Click **Export** → **Export Excel CSV (all fields)**
4. Save the CSV file

The tool reads the **Resolved** date column and calculates weekly throughput automatically. Cycle time is calculated automatically from the **Created** and **Resolved** columns if both are present.

### Scrum mode — Velocity CSV

Jira's velocity chart cannot be exported directly. Create a simple CSV manually by reading the completed points off the Jira velocity chart screen:

```
Sprint,Completed
Sprint 1,42
Sprint 2,38
Sprint 3,51
```

- **Sprint** — sprint name or number (label only, not used in calculations)
- **Completed** — story points completed in that sprint

Save as a `.csv` file and load it in Scrum mode. Aim for at least 6–10 sprints of history.

---

## Usage

1. Launch the tool using the appropriate method for your platform
2. Select your **Team Type** — Kanban or Scrum
3. Browse to your **CSV file**
4. Set your **backlog size** (items for Kanban, story points for Scrum)
5. Set your **weeks or sprints of history** to use
6. Set your **start date** for the forecast (defaults to today)
7. Set your **team availability %** (default 80%)
8. Click **Run Simulation**

Results appear across three tabs:

- **Chart** — distribution histograms with confidence band overlays
- **Summary** — full numeric breakdown including forecast finish dates and cycle time
- **Weekly Throughput** — week-by-week or sprint-by-sprint breakdown

After running, three export options appear at the bottom of the left panel:

- **Save Chart as PNG** — exports the chart image
- **Export Report as HTML** — single self-contained file; opens in any browser, easily emailed
- **Export Report as PDF** — three-page PDF; no additional dependencies required

Hover over the **ⓘ** icons next to each field for guidance on what to enter.

---

## Capacity Adjustment

The **Team availability %** field scales the historical throughput before simulation to reflect realistic team capacity. The default is **80%**, which is the standard agile recommendation accounting for meetings, ceremonies, holidays, and general overhead.

| Setting | Use case |
|---------|----------|
| 100% | No adjustment — raw historical throughput only |
| 80% | Standard recommendation for most teams |
| 70% | High-overhead periods (e.g. PI planning, major releases) |
| 60% | Significant planned absence or reduced team size |

---

## Understanding the Results

### Confidence levels

| Level | Meaning |
|-------|---------|
| 50% | Half of all simulations finished by this point. Optimistic estimate. |
| 70% | A reasonable working estimate. |
| 85% | A safe commitment for most stakeholder conversations. |
| 95% | Near-certain. Use for hard deadlines or release planning. |

### Weeks/sprints chart (left)

Shows how many weeks or sprints were needed across all simulations. A wider distribution means more variability in your historical data.

### Throughput chart (right)

Shows how many items or points were completed in the median number of periods. The `>=` figures are **lower bounds** — in X% of simulations, the team completed *at least* that many items.

### Why the 95% throughput figure is higher than the 50%

The throughput chart asks *"how much will we complete?"* — so higher is better. This is the opposite direction to the weeks chart, which asks *"how long will it take?"* where lower is better.

### A note on accuracy

Monte Carlo simulation produces a probability distribution, not a precise prediction. The results reflect the range of outcomes that history suggests are plausible. The 85% confidence level is generally appropriate for stakeholder conversations; the 95% level for hard external deadlines.

---

## Changelog

| Version | Changes |
|---------|---------|
| v2.6 | Scrum mode re-added with velocity CSV support, sprint length field |
| v2.5 | Splash screen, info tooltips, auto cycle time from issue export |
| v2.4 | Capacity adjustment (%), HTML report export, PDF report export |
| v2.3 | Cross-platform button fix for macOS compatibility |
| v2.2 | Forecast finish dates, scrollable left panel |
| v2.1 | Kanban throughput mode, weekly throughput tab |
| v2.0 | GUI with tkinter |

---

## Tips

- **Minimum history**: At least 5–6 weeks or sprints of data for meaningful results. 10–16 is better.
- **Kanban vs Scrum**: For teams with inconsistent sprint data in Jira, Kanban mode using the Resolved date is more reliable than Scrum velocity — even for teams running Scrum.
- **Fewer weeks = faster forecast**: If the team has improved recently, reducing weeks of history weights the simulation towards recent performance. Use judgement about whether the older data still represents the team.
- **Zero weeks**: Weeks with zero completions are included in the simulation — they widen the upper confidence bands, which is appropriate.
- **Capacity default**: The 80% default is a planning assumption. Adjust it to reflect your team's actual situation.

---

## Licence

MIT Licence — free to use, adapt, and distribute. See [LICENSE](LICENSE) for details.

---

## Contributing

Issues and pull requests welcome. If your Jira export uses column names not automatically detected by the tool, please open an issue with the column headers and they will be added to the auto-detection logic.
