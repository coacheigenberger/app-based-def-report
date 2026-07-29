from __future__ import annotations

import re
import tempfile
from pathlib import Path

import streamlit as st

from def_report_engine import generate

APP_TITLE = "DEF Football Report Generator"
DEFAULT_TEMPLATE = Path(__file__).parent / "assets" / "MASTER Offensive Breakdown Template.pptx"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🏈",
    layout="wide",
)

st.title("🏈 DEF Football Report Generator")
st.caption("Upload Hudl Excel/CSV exports and generate a template-driven offensive breakdown PowerPoint.")

with st.sidebar:
    st.header("Report Setup")
    opponent = st.text_input("Opponent name", placeholder="Example: West")
    odk = st.selectbox("ODK filter", ["O"], index=0)
    use_uploaded_template = st.checkbox("Use a different PowerPoint template")
    custom_template = None
    if use_uploaded_template:
        custom_template = st.file_uploader(
            "Upload .pptx template",
            type=["pptx"],
            accept_multiple_files=False,
        )

st.subheader("1. Upload Opponent Data")
uploaded_files = st.file_uploader(
    "Select one or more Hudl Excel/CSV files",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True,
)

st.subheader("2. Generate Report")
st.write(
    "The app filters to **ODK = O**, validates the files, runs the analytics engine, "
    "and populates the PowerPoint template."
)

def safe_filename(text: str) -> str:
    value = re.sub(r"[^A-Za-z0-9 _-]+", "_", text).strip()
    value = re.sub(r"\s+", "_", value)
    return value or "Opponent"

generate_clicked = st.button(
    "Generate Offensive Breakdown",
    type="primary",
    use_container_width=True,
)

if generate_clicked:
    if not opponent.strip():
        st.error("Enter the opponent name.")
        st.stop()
    if not uploaded_files:
        st.error("Upload at least one Excel or CSV file.")
        st.stop()
    if use_uploaded_template and custom_template is None:
        st.error("Upload the replacement PowerPoint template or turn off the template option.")
        st.stop()

    with st.status("Building report…", expanded=True) as status:
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                work = Path(temp_dir)
                input_paths = []

                st.write("Saving uploaded files securely for this report run…")
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

                st.write("Validating and analyzing offensive snaps…")
                result = generate(
                    input_paths,
                    str(template_path),
                    str(output_path),
                    opponent.strip(),
                    odk,
                )

                st.write("Finalizing PowerPoint…")
                report_bytes = output_path.read_bytes()

            status.update(label="Report completed.", state="complete", expanded=False)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Offensive Plays", result["plays"])
            c2.metric("Run", result["run_pct"])
            c3.metric("Pass", result["pass_pct"])
            c4.metric("Explosives", result["explosives"])

            issues = result.get("issues") or []
            if issues:
                with st.expander("Validation notes"):
                    for issue in issues:
                        st.warning(issue)

            st.download_button(
                "Download Completed PowerPoint",
                data=report_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True,
            )

        except Exception as exc:
            status.update(label="Report generation failed.", state="error", expanded=True)
            st.exception(exc)

st.divider()
st.caption(
    "Uploaded files are used only for the current report run by the application code. "
    "Hosting-provider retention and access policies depend on where the app is deployed."
)
