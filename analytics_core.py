from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd

from def_report_engine import (
    add_features,
    core_metrics,
    pct,
    read_files,
    validate,
)

FILTER_COLUMNS = {
    "personnel": "PERSONNEL",
    "formation": "OFF_FORM",
    "motion": "MOTION",
    "backfield": "BACKFIELD",
    "down_distance": "DND",
    "field_zone": "FIELD_ZONE",
    "hash": "HASH",
    "play": "OFF_PLAY",
    "play_type": "PLAY_TYPE_NORM",
    "player": "PLAYER",
    "field_route": "FIELD_ROUTE",
    "boundary_route": "BOUNDARY_ROUTE",
}

ALIASES = {
    "11 personnel": ("PERSONNEL", "11"),
    "11p": ("PERSONNEL", "11"),
    "12 personnel": ("PERSONNEL", "12"),
    "12p": ("PERSONNEL", "12"),
    "10 personnel": ("PERSONNEL", "10"),
    "10p": ("PERSONNEL", "10"),
    "13 personnel": ("PERSONNEL", "13"),
    "13p": ("PERSONNEL", "13"),
    "20 personnel": ("PERSONNEL", "20"),
    "21 personnel": ("PERSONNEL", "21"),
    "22 personnel": ("PERSONNEL", "22"),
    "1st and 10": ("DND", "1ST & 10"),
    "first and 10": ("DND", "1ST & 10"),
    "2nd and long": ("DND", "2ND & 8+"),
    "second and long": ("DND", "2ND & 8+"),
    "2nd and medium": ("DND", "2ND & 4-7"),
    "second and medium": ("DND", "2ND & 4-7"),
    "2nd and short": ("DND", "2ND & 1-3"),
    "second and short": ("DND", "2ND & 1-3"),
    "3rd and long": ("DND", "3RD & 8+"),
    "third and long": ("DND", "3RD & 8+"),
    "3rd and medium": ("DND", "3RD & 4-7"),
    "third and medium": ("DND", "3RD & 4-7"),
    "3rd and short": ("DND", "3RD & 1-3"),
    "third and short": ("DND", "3RD & 1-3"),
    "4th and long": ("DND", "4TH & 8+"),
    "4th and medium": ("DND", "4TH & 4-7"),
    "4th and short": ("DND", "4TH & 1-3"),
    "low red": ("FIELD_ZONE", "LOW RED ZONE"),
    "high red": ("FIELD_ZONE", "HIGH RED ZONE"),
    "goal line": ("FIELD_ZONE", "GOAL LINE"),
    "midfield": ("FIELD_ZONE", "MIDFIELD"),
    "backed up": ("FIELD_ZONE", "BACKED UP"),
    "coming out": ("FIELD_ZONE", "COMING OUT"),
}

def reliability(total: int) -> str:
    if total >= 20:
        return "HIGH CONFIDENCE"
    if total >= 10:
        return "MEDIUM CONFIDENCE"
    return "LOW SAMPLE"

def predictability_label(value: float) -> str:
    if value >= 90:
        return "EXTREMELY PREDICTABLE"
    if value >= 80:
        return "HIGHLY PREDICTABLE"
    if value >= 70:
        return "STRONG TENDENCY"
    if value >= 60:
        return "SLIGHT TENDENCY"
    return "BALANCED"

@dataclass
class QueryResult:
    filters: Dict[str, str]
    total: int
    metrics: Dict[str, Any]
    top_runs: pd.DataFrame
    top_passes: pd.DataFrame
    formations: pd.DataFrame
    personnel: pd.DataFrame
    motions: pd.DataFrame
    players: pd.DataFrame
    summary: str

class FootballAnalyticsEngine:
    """Single source of truth for report, query, and prediction analytics."""

    def __init__(self, df: pd.DataFrame, issues: Optional[List[str]] = None):
        self.df = df.copy()
        self.issues = issues or []

    @classmethod
    def from_files(cls, paths: List[str], odk: str = "O") -> "FootballAnalyticsEngine":
        raw = read_files(paths)
        issues = validate(raw)
        df = add_features(raw, odk=odk)
        if df.empty:
            raise ValueError("No valid ODK=O Run/Pass plays found after filtering.")
        return cls(df, issues)

    def available_values(self, key: str) -> List[str]:
        col = FILTER_COLUMNS[key]
        return sorted(v for v in self.df[col].dropna().astype(str).unique() if v != "-")

    def slice(self, filters: Optional[Dict[str, Any]] = None) -> pd.DataFrame:
        g = self.df
        for key, value in (filters or {}).items():
            if value in (None, "", "ALL", []):
                continue
            col = FILTER_COLUMNS.get(key, key)
            if col not in g.columns:
                continue
            values = value if isinstance(value, (list, tuple, set)) else [value]
            normalized = {str(v).strip().upper() for v in values}
            g = g[g[col].astype(str).str.strip().str.upper().isin(normalized)]
        return g.copy()

    @staticmethod
    def _ranking(g: pd.DataFrame, column: str, play_type: Optional[str] = None, n: int = 5) -> pd.DataFrame:
        if g.empty or column not in g.columns:
            return pd.DataFrame(columns=["Value", "Calls", "Frequency", "YPP", "EFF", "Explosives"])
        base = g
        if play_type:
            base = base[base["PLAY_TYPE_NORM"] == play_type]
        denom = len(base)
        rows = []
        for value, sub in base.groupby(column, dropna=False):
            calls = len(sub)
            rows.append({
                "Value": str(value),
                "Calls": calls,
                "Frequency": pct(calls, denom),
                "YPP": round(float(sub["YARDS"].mean()), 1) if calls else 0,
                "EFF": pct(int(sub["EFF_BOOL"].sum()), calls),
                "Explosives": int(sub["IS_EXPLOSIVE"].sum()),
            })

        # A valid query can return zero rows for one side of the breakdown
        # (for example, no pass plays in a fully run-only slice). Return an
        # empty table with the expected schema instead of sorting a DataFrame
        # that has no columns.
        columns = ["Value", "Calls", "Frequency", "YPP", "EFF", "Explosives"]
        if not rows:
            return pd.DataFrame(columns=columns)

        return (
            pd.DataFrame(rows, columns=columns)
            .sort_values(["Calls", "YPP"], ascending=[False, False])
            .head(n)
            .reset_index(drop=True)
        )

    def query(self, filters: Optional[Dict[str, Any]] = None) -> QueryResult:
        g = self.slice(filters)
        m = core_metrics(g)
        run_tendency = 100 * m["run"] / m["total"] if m["total"] else 0
        pass_tendency = 100 * m["pass"] / m["total"] if m["total"] else 0
        dominant = "Run" if run_tendency >= pass_tendency else "Pass"
        dominant_pct = max(run_tendency, pass_tendency)
        top_run = self._ranking(g, "OFF_PLAY", "Run", 3)
        top_pass = self._ranking(g, "OFF_PLAY", "Pass", 3)
        top_name = "-"
        source = top_run if dominant == "Run" else top_pass
        if not source.empty:
            top_name = str(source.iloc[0]["Value"])
        filter_text = " | ".join(f"{k.replace('_', ' ').title()}: {v}" for k, v in (filters or {}).items()) or "All offensive snaps"
        summary = (
            f"{filter_text} produced {m['total']} snaps. The offense was {m['run_pct']} run and "
            f"{m['pass_pct']} pass. The dominant tendency was {dominant} "
            f"({predictability_label(dominant_pct)}; {reliability(m['total'])}). "
            f"The leading {dominant.lower()} concept was {top_name}. Overall efficiency was "
            f"{m['eff_pct']} with {m['explosive']} explosive plays."
        )
        return QueryResult(
            filters=filters or {},
            total=m["total"],
            metrics=m,
            top_runs=top_run,
            top_passes=top_pass,
            formations=self._ranking(g, "OFF_FORM", n=5),
            personnel=self._ranking(g, "PERSONNEL", n=5),
            motions=self._ranking(g, "MOTION", n=5),
            players=self._ranking(g, "PLAYER", n=5),
            summary=summary,
        )

    def parse_question(self, question: str) -> Tuple[Dict[str, str], Optional[str]]:
        q = re.sub(r"\s+", " ", question.lower()).strip()
        filters: Dict[str, str] = {}
        requested_play: Optional[str] = None

        for phrase, (col, value) in sorted(ALIASES.items(), key=lambda x: -len(x[0])):
            if phrase in q:
                key = next((k for k, c in FILTER_COLUMNS.items() if c == col), col)
                filters[key] = value

        # Match exact values already present in the uploaded data.
        searchable = ["OFF_FORM", "MOTION", "BACKFIELD", "OFF_PLAY", "PLAYER", "HASH"]
        for col in searchable:
            values = sorted(self.df[col].dropna().astype(str).unique(), key=len, reverse=True)
            for value in values:
                if value == "-" or len(value) < 2:
                    continue
                if re.search(rf"(?<!\w){re.escape(value.lower())}(?!\w)", q):
                    key = next((k for k, c in FILTER_COLUMNS.items() if c == col), col)
                    if col == "OFF_PLAY" and any(x in q for x in ["when do", "how often", "show", "run "]):
                        requested_play = value
                    else:
                        filters[key] = value
                    break

        if "pass" in q and "run/pass" not in q:
            filters.setdefault("play_type", "Pass")
        elif re.search(r"\brun\b", q) and not any(x in q for x in ["run/pass", "when do they run"]):
            filters.setdefault("play_type", "Run")

        return filters, requested_play

    def answer(self, question: str) -> QueryResult:
        filters, requested_play = self.parse_question(question)
        if requested_play:
            filters["play"] = requested_play
        return self.query(filters)

    def identity(self) -> Dict[str, Any]:
        total = len(self.df)
        m = core_metrics(self.df)
        top_personnel = self._ranking(self.df, "PERSONNEL", n=1)
        top_form = self._ranking(self.df, "OFF_FORM", n=1)
        top_motion = self._ranking(self.df[self.df["MOTION"] != "-"], "MOTION", n=1)
        run_concepts = self.df[self.df["IS_RUN"]]["OFF_PLAY"].nunique()
        pass_concepts = self.df[self.df["IS_PASS"]]["OFF_PLAY"].nunique()
        form_count = self.df["OFF_FORM"].nunique()
        motion_rate = 100 * (self.df["MOTION"] != "-").sum() / total if total else 0

        def diversity(count: int) -> str:
            return "HIGH" if count >= 10 else "MODERATE" if count >= 5 else "LOW"

        dominant = "RUN" if m["run"] >= m["pass"] else "PASS"
        top_p = top_personnel.iloc[0]["Value"] if not top_personnel.empty else "-"
        top_f = top_form.iloc[0]["Value"] if not top_form.empty else "-"
        top_m = top_motion.iloc[0]["Value"] if not top_motion.empty else "NO MOTION"
        narrative = (
            f"This is a {dominant}-leaning offense ({m['run_pct']} run, {m['pass_pct']} pass). "
            f"Its primary personnel is {top_p}, and its most-used formation is {top_f}. "
            f"Motion appeared on {pct((self.df['MOTION'] != '-').sum(), total)} of snaps, led by {top_m}. "
            f"The offense showed {diversity(form_count).lower()} formation diversity and generated "
            f"{m['explosive']} explosives."
        )
        return {
            "metrics": m,
            "top_personnel": top_p,
            "top_formation": top_f,
            "top_motion": top_m,
            "formation_diversity": diversity(form_count),
            "run_concept_diversity": diversity(run_concepts),
            "pass_concept_diversity": diversity(pass_concepts),
            "motion_usage": "HEAVY" if motion_rate >= 40 else "MODERATE" if motion_rate >= 20 else "MINIMAL",
            "narrative": narrative,
        }

    def strongest_tendencies(self, min_sample: int = 5, n: int = 12) -> pd.DataFrame:
        candidates = []
        dimensions = [
            ("DND",), ("PERSONNEL",), ("OFF_FORM",),
            ("PERSONNEL", "OFF_FORM"), ("DND", "OFF_FORM"),
            ("PERSONNEL", "OFF_FORM", "MOTION"),
        ]
        for dims in dimensions:
            for keys, sub in self.df.groupby(list(dims), dropna=False):
                if len(sub) < min_sample:
                    continue
                keys = keys if isinstance(keys, tuple) else (keys,)
                m = core_metrics(sub)
                run_pct = 100 * m["run"] / m["total"]
                pass_pct = 100 * m["pass"] / m["total"]
                tendency = "RUN" if run_pct >= pass_pct else "PASS"
                strength = max(run_pct, pass_pct)
                if strength < 70:
                    continue
                label = " | ".join(f"{d}: {v}" for d, v in zip(dims, keys))
                score = strength * min(1.0, len(sub) / 20) + m["explosive"] * 1.5
                candidates.append({
                    "Situation": label,
                    "Snaps": len(sub),
                    "Tendency": tendency,
                    "Rate": pct(m["run"] if tendency == "RUN" else m["pass"], m["total"]),
                    "Predictability": predictability_label(strength),
                    "Confidence": reliability(len(sub)),
                    "EFF": m["eff_pct"],
                    "Explosives": m["explosive"],
                    "_score": score,
                })
        if not candidates:
            return pd.DataFrame(columns=["Situation","Snaps","Tendency","Rate","Predictability","Confidence","EFF","Explosives"])
        return pd.DataFrame(candidates).sort_values("_score", ascending=False).drop(columns="_score").head(n).reset_index(drop=True)
