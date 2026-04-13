#!/bin/bash
# Monte Carlo Projection Tool — Mac Launcher
# Activates the virtual environment and runs the app.

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ENV_PATH="$HOME/monte_carlo_env"
SCRIPT="$SCRIPT_DIR/monte_carlo_jira_v2.3.py"

# Check virtual environment exists
if [ ! -f "$ENV_PATH/bin/activate" ]; then
    osascript -e 'display alert "Setup Required" message "Virtual environment not found at ~/monte_carlo_env\n\nPlease run the setup instructions in the README." as critical'
    exit 1
fi

# Check script exists
if [ ! -f "$SCRIPT" ]; then
    osascript -e 'display alert "File Not Found" message "monte_carlo_jira_v2.3.py not found in the same folder as this launcher." as critical'
    exit 1
fi

# Activate environment and run
source "$ENV_PATH/bin/activate"
python "$SCRIPT"
