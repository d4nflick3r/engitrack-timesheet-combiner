# EngiTrack Timesheet Combiner

A Streamlit app that combines up to 30 EngiTrack engineer timesheet CSV files into a single formatted Excel workbook.

## Features
- Upload up to 30 CSV timesheet files at once
- Preview each timesheet before combining
- Styled Excel output with 6 sheets: Instructions, Engineer_List, Week_List, Weekly_Summary, WeeklyData, DailyData
- Auto-filter, alternating row colours, frozen header rows
- Download the finished `.xlsx` workbook directly

## Stack
- Python 3.11
- Streamlit (UI)
- pandas (CSV reading)
- openpyxl (Excel workbook generation)

## Running (web)
```bash
streamlit run app.py --server.port 5000
```

## Structure
- `app.py` — main Streamlit application
- `patch_pwa.py` — patches Streamlit's index.html to include PWA manifest link
- `windows_launcher.py` — entry point for the Windows desktop build
- `windows/` — Windows batch scripts and .ico file
- `.github/workflows/build-windows.yml` — GitHub Actions workflow to build Windows .exe
- `static/manifest.json` — PWA web app manifest
- `static/icons/` — app icons (192×192, 512×512 PNG)
- `static/screenshot-1.png` — PWA store screenshot
- `.streamlit/config.toml` — Streamlit server configuration

## PWA (web install)
The hosted app at `timesheet-merger.replit.app` is installable as a PWA via Edge/Chrome.
Manifest link is injected directly into Streamlit's `index.html` by `patch_pwa.py` at startup.

## Windows Desktop App
To build a standalone Windows `.exe`:
1. Push the repository to GitHub
2. Go to Actions → "Build Windows App" → Run workflow
3. Download `EngiTrack.exe` from the Artifacts section
4. Run it on any Windows 10/11 machine (no Python required)

Or to run locally with Python:
1. Run `windows/install.bat` once
2. Run `windows/run.bat` to open the app
