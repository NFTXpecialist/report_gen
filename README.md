# Business Establishment Report Generator

This Streamlit app generates detailed Excel reports from your local business establishment data. It features advanced filtering, summary cards, interactive charts, and a dedicated Inspectors Output section for team/inspector analytics.

## Features
- Upload Excel files (expects a sheet named `DATABASE`)
- Filter by STATUS, date returned, AOR, CATEGORY, Occupancy Type, and more
- Generate reports by AOR, STATUS, Occupancy Type, NEW/RENEW
- Dedicated Inspectors Output: team-level summaries, drilldowns, and Excel export
- Download all reports and filtered data as Excel
- Interactive charts and summary cards
- Data validation and issues report

## Deploy on Streamlit Cloud
1. Push this repo to your GitHub account
2. Go to [streamlit.io/cloud](https://streamlit.io/cloud) and connect your repo
3. Set the main file as `app.py`
4. Deploy and enjoy!

## Requirements
See `requirements.txt` for dependencies.

## Usage
- Upload your Excel file
- Select columns as prompted (if auto-detection fails)
- Use the filters and download/export features as needed
