from __future__ import annotations

import re
import tempfile
from pathlib import Path

import streamlit as st

from def_report_engine import generate

APP_TITLE = "DEF Offensive Report PowerPoint Generator"

ROOT = Path(__file__).parent
DEFAULT_TEMPLATE_OPTIONS = [
    ROOT / "assets" / "MASTER Offensive Breakdown Template.pptx",
    ROOT / "MASTER Offensive Breakdown Template.pptx",
]
DEFAULT_TEMPLATE = next((p for p in DEFAULT_TEMPLATE_OPTIONS if p.exists()), DEFAULT_TEMPLATE_OPTIONS[0])

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🏈",
    layout="centered",
)

st.title("🏈 DEF Offensive Report Generator")
st.subheader("Create the athlete-ready PowerPoint presentation")
st.write(
    "Upload the Hudl Excel/CSV files, enter the opponent name, and download the completed "
    "PowerPoint built from the master offensive breakdown template."
)

with st.expander("What this app does", expanded=False):
    st.markdown(
        """
        - Filters to **ODK = O**
        - Runs the offensive tendency analytics
        - Builds formation, situation, explosive, touch, prediction, and predictability sections
        - Populates the PowerPoint report
        - Returns a downloadable `.pptx` presentation for your team
        """
    )

st.divider()

opponent = st.text_input(
    "Opponent name",
    placeholder="Example: Eau Claire Memorial",
)

uploaded_files = st.file_uploader(
    "Upload Hudl Excel/CSV file(s)",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True,
    help="You may upload one game file or multiple game files for the same opponent.",
)

with st.expander("Advanced options", expanded=False):
    odk = st.selectbox("ODK filter", ["O"], index=0)
    st.caption("Default is O for offensive snaps.")
    min_sample = st.slider(
        "Minimum snaps for predictive tendencies",
        min_value=3,
        max_value=15,
        value=5,
        step=1,
        help="Controls which tells, predictions, call-sheet rows, and alert slices appear. Percentages still show counts like 75% (6/8).",
    )
    use_uploaded_template = st.checkbox("Use a different PowerPoint template for this run")
    custom_template = None
    if use_uploaded_template:
        custom_template = st.file_uploader(
            "Upload replacement .pptx template",
            type=["pptx"],
            accept_multiple_files=False,
        )
    st.caption(
        f"Default template status: {'found' if DEFAULT_TEMPLATE.exists() else 'missing'} — {DEFAULT_TEMPLATE}"
    )

def safe_filename(text: str) -> str:
    value = re.sub(r"[^A-Za-z0-9 _-]+", "_", text).strip()
    value = re.sub(r"\s+", "_", value)
    return value or "Opponent"

generate_clicked = st.button(
    "Generate PowerPoint Presentation",
    type="primary",
    use_container_width=True,
)

if generate_clicked:
    if not opponent.strip():
        st.error("Enter the opponent name.")
        st.stop()
    if not uploaded_files:
        st.error("Upload at least one Hudl Excel or CSV file.")
        st.stop()
    if not DEFAULT_TEMPLATE.exists() and not custom_template:
        st.error(
            "The master template was not found. Upload a template under Advanced options "
            "or add MASTER Offensive Breakdown Template.pptx to the repository."
        )
        st.stop()
    if use_uploaded_template and custom_template is None:
        st.error("Upload the replacement PowerPoint template or turn off the template option.")
        st.stop()

    with st.status("Generating PowerPoint presentation…", expanded=True) as status:
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                work = Path(temp_dir)
                input_paths = []

                st.write("Reading uploaded Hudl file(s)…")
                for upload in uploaded_files:
                    path = work / Path(upload.name).name
                    path.write_bytes(upload.getbuffer())
                    input_paths.append(str(path))

                if custom_template is not None:
                    template_path = work / "uploaded_template.pptx"
                    template_path.write_bytes(custom_template.getbuffer())
                else:
                    template_path = DEFAULT_TEMPLATE

                output_name = f"{safe_filename(opponent)}_Offensive_Breakdown.pptx"
                output_path = work / output_name

                st.write("Filtering to ODK = O…")
                st.write("Running offensive tendency analytics…")
                st.write("Populating the PowerPoint template…")

                result = generate(
                    input_paths,
                    str(template_path),
                    str(output_path),
                    opponent.strip(),
                    odk,
                    min_sample=min_sample,
                )

                report_bytes = output_path.read_bytes()

            status.update(label="PowerPoint completed.", state="complete", expanded=False)

            st.success("Your athlete-ready PowerPoint is ready.")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("O Plays", result["plays"])
            c2.metric("Run", result["run_pct"])
            c3.metric("Pass", result["pass_pct"])
            c4.metric("Explosives", result["explosives"])

            st.download_button(
                "⬇️ Download PowerPoint Presentation",
                data=report_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True,
            )

            issues = result.get("issues") or []
            if issues:
                with st.expander("Validation notes for staff review"):
                    for issue in issues:
                        st.warning(issue)

        except Exception as exc:
            status.update(label="PowerPoint generation failed.", state="error", expanded=True)
            st.error("The report could not be generated. Share the error below so it can be fixed.")
            st.exception(exc)

st.divider()
st.caption(
    "Designed for one weekly output: a completed offensive scouting PowerPoint presentation."
)
