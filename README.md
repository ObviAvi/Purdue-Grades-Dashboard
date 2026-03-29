# Purdue Grade Distribution Dashboard

An interactive dashboard for exploring Purdue course grade distributions across terms, built from publicly available Purdue grade distribution exports.

## Features

- Filter by subject, course, instructor, and semester
- Compare grade trends over time
- View grade buckets (A/B/C/D/F/W/Other) across sections
- Fast client-side dashboard using prebuilt JSON data

## Data Source

This project uses **publicly available Purdue grade distribution data** exported to Excel and stored locally as `grades.xlsx`.

- Input data file: `grades.xlsx`
- Generated data file: `grades_data.json`
- Builder script: `build_dashboard_data.py`

## Project Files

- `index.html` - Dashboard UI and filtering/chart logic
- `build_dashboard_data.py` - Parses all sheets in `grades.xlsx` and builds compact JSON data
- `grades_data.json` - Dashboard data consumed by the frontend
- `package.json` - Node package metadata/dependencies

## Requirements

- Python 3.10+
- Node.js 18+ (optional, only needed if serving locally with a dev server)

Python packages used by the builder script:

- `pandas`
- `openpyxl`

Install Python dependencies:

```bash
pip install pandas openpyxl
```

## Usage

### 1. Rebuild dashboard data

After updating `grades.xlsx`, regenerate the JSON:

```bash
python build_dashboard_data.py
```

### 2. Run the dashboard

You can open `index.html` directly in a browser, or serve the folder locally.

Example local static server:

```bash
npx serve .
```

Then open the local URL shown in your terminal.