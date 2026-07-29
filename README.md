# DEF Offensive Report PowerPoint Generator

This Streamlit web app has one job:

**Upload Hudl Excel/CSV files and download an athlete-ready offensive scouting PowerPoint.**

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

The master template should be inside the `assets` folder. The app also supports the template in the repository root as a backup.

## Updating your existing Streamlit app

Replace these files in your GitHub repository:

- `app.py`
- `README.md`

Keep the existing:

- `def_report_engine.py`
- `requirements.txt`
- `assets/MASTER Offensive Breakdown Template.pptx`

Streamlit should automatically redeploy after the commit. If it does not, open Streamlit Cloud and click **Reboot app** or **Redeploy**.

## Purpose

The PowerPoint is the final product. Dashboards and extra analysis screens are intentionally avoided so the weekly workflow stays simple for coaches and staff.


## Version 2.1 update

This update changes the report output in two important ways:

1. Duplicated slide titles preserve the original PowerPoint template formatting.
2. Every percentage now shows its numerator and denominator, for example:
   `25% (2/8)`

To update an existing Streamlit deployment, replace `def_report_engine.py` in GitHub and commit the change. Streamlit will redeploy automatically.
