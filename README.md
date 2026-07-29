# DEF Football Report Generator — Web App

This browser-based application generates an offensive tendency PowerPoint from one or more Hudl Excel/CSV exports.

## Weekly workflow

1. Open the web app.
2. Enter the opponent name.
3. Upload one or more weekly Excel/CSV exports.
4. Click **Generate Offensive Breakdown**.
5. Download the completed `.pptx`.

No Python installation is required on the school computer after the app is deployed.

## Easiest deployment: Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload the contents of this folder to the repository root.
3. Sign in to Streamlit Community Cloud using GitHub.
4. Create a new app from the repository.
5. Set the main file to `app.py`.
6. Deploy.
7. Bookmark the generated web address.

The exact button names may change over time, but the required settings are the repository, branch, and `app.py`.

## Important privacy note

Opponent data will be processed by whichever hosting service runs the app. Before using school or team data, confirm that the hosting service is acceptable under your school, district, conference, and team policies. For stricter privacy, deploy the same package to a school-approved private server or cloud account.

## Included files

- `app.py` — browser interface
- `def_report_engine.py` — analytics and PowerPoint engine
- `assets/MASTER Offensive Breakdown Template.pptx` — default template
- `requirements.txt` — cloud dependencies
- `DEF_Football_Analytics_Engine_v2_1.docx` — analytics rules reference

## Updating the PowerPoint template

Replace:

`assets/MASTER Offensive Breakdown Template.pptx`

with the revised template, keeping the same filename. The app also allows a template to be uploaded for a single report run.
