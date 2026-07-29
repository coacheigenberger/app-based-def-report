# DEF Analyst v3.0

DEF Analyst v3.0 turns the existing PowerPoint report generator into a reusable offensive scouting platform.

## New modules

- **Report** — generates the existing athlete-ready PowerPoint.
- **Ask the Offense** — answers plain-language football questions using only the uploaded opponent data.
- **Tendencies** — filters by personnel, formation, down-and-distance, field zone, motion, backfield, and hash.
- **Predictions** — ranks the strongest run/pass tells and shows sample-size confidence.
- **Game Plan** — summarizes offensive identity and priority alerts.

All percentages retain their supporting totals, such as `75% (6/8)`.

## GitHub update

Replace or add:

- `app.py`
- `analytics_core.py`
- `def_report_engine.py`
- `README.md`

Keep:

- `requirements.txt`
- `assets/MASTER Offensive Breakdown Template.pptx`

Commit the changes. Streamlit Cloud should redeploy automatically.

## Weekly use

1. Enter the opponent.
2. Upload one or more Hudl files.
3. Click **Load and Analyze Data**.
4. Use the interactive tabs throughout the week.
5. Generate and download the PowerPoint from the Report tab.

## Important scope

The platform uses only the current uploaded files. It does not import general football tendencies or outside opponent information.
