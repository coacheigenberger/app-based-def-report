from __future__ import annotations

import re
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from analytics_core import FootballAnalyticsEngine
from def_report_engine import generate

APP_TITLE = "DEF Analyst"
ROOT = Path(__file__).parent
DEFAULT_TEMPLATE_OPTIONS = [
    ROOT / "assets" / "MASTER Offensive Breakdown Template.pptx",
    ROOT / "MASTER Offensive Breakdown Template.pptx",
]
DEFAULT_TEMPLATE = next((p for p in DEFAULT_TEMPLATE_OPTIONS if p.exists()), DEFAULT_TEMPLATE_OPTIONS[0])

st.set_page_config(page_title=APP_TITLE, page_icon="🏈", layout="wide")
st.title("🏈 DEF Analyst")
st.caption("PowerPoint report generator and interactive opponent-offense analyst")

def safe_filename(text: str) -> str:
    value = re.sub(r"[^A-Za-z0-9 _-]+", "_", text).strip()
    return re.sub(r"\s+", "_", value) or "Opponent"

def save_uploads(files, folder: Path):
    paths = []
    for upload in files:
        path = folder / Path(upload.name).name
        path.write_bytes(upload.getbuffer())
        paths.append(str(path))
    return paths

def render_result(result):
    st.markdown(result.summary)
    m = result.metrics
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Snaps", m["total"])
    c2.metric("Run", m["run_pct"])
    c3.metric("Pass", m["pass_pct"])
    c4.metric("Efficiency", m["eff_pct"])
    c5.metric("Explosives", m["explosive"])
    a, b = st.columns(2)
    with a:
        st.markdown("#### Top runs")
        st.dataframe(result.top_runs, use_container_width=True, hide_index=True)
        st.markdown("#### Formations")
        st.dataframe(result.formations, use_container_width=True, hide_index=True)
    with b:
        st.markdown("#### Top passes")
        st.dataframe(result.top_passes, use_container_width=True, hide_index=True)
        st.markdown("#### Personnel")
        st.dataframe(result.personnel, use_container_width=True, hide_index=True)

with st.sidebar:
    st.header("Current opponent")
    opponent = st.text_input("Opponent name", placeholder="Example: Eau Claire Memorial")
    uploaded_files = st.file_uploader(
        "Hudl Excel/CSV file(s)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
    )
    min_sample = st.slider("Minimum predictive sample", 3, 15, 5)
    load_clicked = st.button("Load and Analyze Data", type="primary", use_container_width=True)
    st.caption("All analysis uses only the files uploaded in this session.")

if load_clicked:
    if not uploaded_files:
        st.error("Upload at least one Hudl file.")
    else:
        try:
            with tempfile.TemporaryDirectory() as td:
                paths = save_uploads(uploaded_files, Path(td))
                engine = FootballAnalyticsEngine.from_files(paths, odk="O")
            st.session_state["engine"] = engine
            st.session_state["opponent"] = opponent.strip() or "Opponent"
            st.success(f"Loaded {len(engine.df)} offensive snaps.")
        except Exception as exc:
            st.exception(exc)

engine = st.session_state.get("engine")
tabs = st.tabs(["📊 Report", "🔍 Ask the Offense", "📈 Tendencies", "🎯 Predictions", "📋 Game Plan"])

with tabs[0]:
    st.subheader("Athlete-ready PowerPoint")
    st.write("Generate the full offensive breakdown from the same validated dataset used by the interactive analyst.")
    if engine is None:
        st.info("Upload files and click **Load and Analyze Data** in the sidebar.")
    else:
        identity = engine.identity()
        st.info(identity["narrative"])
        custom_template = st.file_uploader("Optional replacement PowerPoint template", type=["pptx"], key="template")
        if st.button("Generate PowerPoint Presentation", type="primary", use_container_width=True):
            if not DEFAULT_TEMPLATE.exists() and custom_template is None:
                st.error("Master PowerPoint template not found.")
            else:
                with st.status("Generating PowerPoint…", expanded=True) as status:
                    try:
                        with tempfile.TemporaryDirectory() as td:
                            work = Path(td)
                            # Recreate a temporary CSV from the validated offense dataset.
                            csv_path = work / "validated_offense.csv"
                            engine.df.to_csv(csv_path, index=False)
                            # The report loader expects ODK and original columns.
                            report_df = engine.df.copy()
                            report_df["ODK"] = "O"
                            report_df["PLAY TYPE"] = report_df["PLAY_TYPE_NORM"]
                            report_df["GN/LS"] = report_df["YARDS"]
                            report_df.to_csv(csv_path, index=False)

                            if custom_template is not None:
                                template_path = work / "uploaded_template.pptx"
                                template_path.write_bytes(custom_template.getbuffer())
                            else:
                                template_path = DEFAULT_TEMPLATE

                            name = st.session_state.get("opponent", "Opponent")
                            output_name = f"{safe_filename(name)}_Offensive_Breakdown.pptx"
                            output_path = work / output_name
                            result = generate(
                                [str(csv_path)],
                                str(template_path),
                                str(output_path),
                                name,
                                "O",
                                min_sample=min_sample,
                            )
                            report_bytes = output_path.read_bytes()

                        status.update(label="PowerPoint completed.", state="complete", expanded=False)
                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric("O Plays", result["plays"])
                        c2.metric("Run", result["run_pct"])
                        c3.metric("Pass", result["pass_pct"])
                        c4.metric("Explosives", result["explosives"])
                        st.download_button(
                            "⬇️ Download PowerPoint Presentation",
                            report_bytes,
                            output_name,
                            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            type="primary",
                            use_container_width=True,
                        )
                    except Exception as exc:
                        status.update(label="Generation failed.", state="error")
                        st.exception(exc)

with tabs[1]:
    st.subheader("Ask the Offense")
    st.write("Ask a football question. Answers are calculated only from the current uploaded data.")
    if engine is None:
        st.info("Load opponent files first.")
    else:
        examples = [
            "What do they do from 11P on 3rd and medium?",
            "When do they run Counter?",
            "What do they do in the low red zone?",
            "What is their tendency from Trips?",
        ]
        question = st.text_input("Question", placeholder=examples[0])
        st.caption("Examples: " + " • ".join(examples))
        if st.button("Analyze Question", type="primary") and question.strip():
            render_result(engine.answer(question))

with tabs[2]:
    st.subheader("Interactive Tendencies")
    if engine is None:
        st.info("Load opponent files first.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        personnel = c1.selectbox("Personnel", ["ALL"] + engine.available_values("personnel"))
        formation = c2.selectbox("Formation", ["ALL"] + engine.available_values("formation"))
        dnd = c3.selectbox("Down & Distance", ["ALL"] + engine.available_values("down_distance"))
        zone = c4.selectbox("Field Zone", ["ALL"] + engine.available_values("field_zone"))
        c5, c6, c7 = st.columns(3)
        motion = c5.selectbox("Motion", ["ALL"] + engine.available_values("motion"))
        backfield = c6.selectbox("Backfield", ["ALL"] + engine.available_values("backfield"))
        hash_value = c7.selectbox("Hash", ["ALL"] + engine.available_values("hash"))
        filters = {
            "personnel": personnel, "formation": formation, "down_distance": dnd,
            "field_zone": zone, "motion": motion, "backfield": backfield, "hash": hash_value,
        }
        render_result(engine.query(filters))

with tabs[3]:
    st.subheader("Predictions and Predictability")
    if engine is None:
        st.info("Load opponent files first.")
    else:
        st.write("Highest-value run/pass tells, weighted by tendency strength, sample size, and explosives.")
        table = engine.strongest_tendencies(min_sample=min_sample, n=20)
        st.dataframe(table, use_container_width=True, hide_index=True)
        if table.empty:
            st.warning("No tendencies met the current minimum sample and 70% predictability threshold.")

with tabs[4]:
    st.subheader("Game Plan")
    if engine is None:
        st.info("Load opponent files first.")
    else:
        identity = engine.identity()
        st.markdown("### Offensive Identity")
        st.write(identity["narrative"])
        cols = st.columns(6)
        cols[0].metric("Top Personnel", identity["top_personnel"])
        cols[1].metric("Top Formation", identity["top_formation"])
        cols[2].metric("Top Motion", identity["top_motion"])
        cols[3].metric("Formation Variety", identity["formation_diversity"])
        cols[4].metric("Run Concepts", identity["run_concept_diversity"])
        cols[5].metric("Pass Concepts", identity["pass_concept_diversity"])

        st.markdown("### Priority Alerts")
        alerts = engine.strongest_tendencies(min_sample=min_sample, n=8)
        st.dataframe(alerts, use_container_width=True, hide_index=True)
        st.caption(
            "These are opponent tendencies, not automatic defensive calls. Staff should pair them with film, "
            "front/coverage structure, personnel availability, and game-plan rules."
        )
