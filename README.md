# LSC Buyer Journey Intelligence Engine

This repository contains a Streamlit app that processes Life Science Connect (LSC) exports and produces buyer-journey intelligence and multi-tab Excel reports.

**Features**
- Upload XLSX/CSV exports from LSC
- Validate required columns and normalize headers
- Process large data volumes (~40,000+ rows)
- Determine buyer-journey stages (Awareness → Problem Definition → Solution Exploration → Vendor Evaluation)
- Aggregate to one row per individual and generate company-level summaries
- Generate the following sheets: `Company_Combined`, `Company_Summary`, `Individuals`, `Product_Map`, `Sales_Hot_List`
- Filter out low-engagement individuals (default minimum 5 engagements)
- Optional enrichment stub for high-intent companies
- Generates a multi-tab Excel file with color-coding and a download link

## Requirements
- Python 3.10
- See `requirements.txt` (Streamlit, pandas, numpy, openpyxl)

## Run locally
1. Create a virtual environment (recommended):

```bash
python3.10 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

2. Start the app:

```bash
streamlit run app.py --server.address 0.0.0.0 --server.port 8000
```

3. Open your browser at `http://localhost:8000`.

## Run in GitHub Codespaces / Dev Container
This repository includes a `.devcontainer` configuration. Open the repo in Codespaces or VS Code Remote - Containers and the container will install dependencies and run the app automatically. The app will be available at port `8000`.

## Files
- `app.py` — main Streamlit application
- `index.html` — landing page with instructions and "Open App" button (opens in new tab)
- `requirements.txt` — pinned Python dependencies
- `.devcontainer/` — devcontainer config and postStart script

## Commands
- Install deps: `pip install -r requirements.txt`
- Run locally: `streamlit run app.py --server.address 0.0.0.0 --server.port 8000`
- In Codespaces: open repo -> forward port 8000 -> click preview or open in browser

## Screenshot placeholders
- `screenshots/dashboard.png` — (placeholder) app dashboard
- `screenshots/excel_preview.png` — (placeholder) Excel report preview

---
If you want, I can:
- Add unit tests for the processing functions
- Add more advanced enrichment integration (e.g., webhooks or 3rd-party APIs)
- Add an option to export CSV versions of sheets
