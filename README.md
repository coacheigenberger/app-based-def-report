# DEF Offensive Report PowerPoint Generator v2.2

This Streamlit web app has one job:

**Upload Hudl Excel/CSV files and download an athlete-ready offensive scouting PowerPoint.**

## v2.2 improvements

- Percentages continue to show counts: `75% (6/8)`
- Predictive tendencies now include sample reliability labels
- Added a **Top Defensive Alerts** slide
- Added a **Predictive Tells** slide
- Added a **Game Day Call Sheet** slide
- Added a **Data Quality Check** slide
- Predictions now show top two likely play calls when available
- Tendencies are ranked by usefulness, not just raw percentage
- The app includes an advanced minimum-sample slider

## Weekly workflow

1. Open the Streamlit app.
2. Enter the opponent name.
3. Upload one or more Hudl Excel/CSV files.
4. Click **Generate PowerPoint Presentation**.
5. Download the completed `.pptx`.

## Required GitHub files

Make sure these are in the repository:

- `app.py`
- `def_report_engine.py`
- `requirements.txt`
- `assets/MASTER Offensive Breakdown Template.pptx`

## Updating your existing Streamlit app

Replace these files in your GitHub repository:

- `app.py`
- `def_report_engine.py`
- `README.md`

Keep:

- `requirements.txt`
- `assets/MASTER Offensive Breakdown Template.pptx`

Streamlit should automatically redeploy after the commit. If it does not, open Streamlit Cloud and click **Reboot app** or **Redeploy**.
