# Social Media Scorecard Automator 📊

A premium Python tool for automating the generation of visually stunning social media scorecards. 

## Features
- **Premium Design**: Generates 1000x1200 PNG scorecards with smooth Outfit typography.
- **Smart Matching**: Automatically links order data with scorecard records using fuzzy string matching.
- **Dynamic Charts**: Visualizes daily order distribution (excluding biteME orders) with color-coded status.
- **SVG Integration**: Supports vector branding assets (B n B logo included).
- **Streamlit Interface**: User-friendly web UI for uploading CSV/Excel data and batch-downloading results.

## Setup
1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   playwright install chromium
   ```
2. Run the application:
   ```bash
   streamlit run app.py
   ```

## Files
- `app.py`: Main Streamlit application and core rendering logic.
- `generate_scorecards.py`: Logic for parsing data and generating bulk PNGs.
- `B n B.svg`: Branding logo used in scorecard headers.
- `requirements.txt`: Python package list.
- `packages.txt`: System-level dependencies for Playwright.
