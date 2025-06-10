#!/bin/bash

cd "$(dirname "$0")"

# Ensure Python 3 is installed
if ! command -v python3 &> /dev/null
then
    echo "Python 3 not found. Please install Python 3."
    exit 1
fi

# Install dependencies
echo "📦 Installing Python dependencies..."
python3 -m pip install -r requirements.txt

# Run the KPI report generator
echo "🔄 Generating sprint report..."
python3 jira_kpi_report.py

# Check if the report file was created
REPORT_FILE="sprint_report.xlsx"
if [ -f "$REPORT_FILE" ]; then
    echo "✅ Report created: $REPORT_FILE"

    # Run the pie chart enhancer
    echo "📊 Adding pie charts..."
    python3 jira_kpi_report_pie_gen.py

    if [ $? -eq 0 ]; then
        echo "✅ Pie charts added successfully."

        # Auto-open on macOS
        if [[ "$OSTYPE" == "darwin"* ]]; then
            open "$REPORT_FILE"
        fi
    else
        echo "❌ Pie chart generation failed."
    fi
else
    echo "❌ Error: Report file not found. Pie charts were not added."
fi
