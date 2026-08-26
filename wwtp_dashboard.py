# -*- coding: utf-8 -*-
"""
WWTP Performance & Diagnostic Platform
======================================

Data source:
    WWTP_Analytics_Database.xlsx

Expected sheets:
    Daily_KPI
    KPI_Definition
    KPI_Units
    KPI_Target

GitHub / Streamlit Cloud:
    Put this .py file, requirements.txt and WWTP_Analytics_Database.xlsx
    in the repository. Change DATA_FILE below if the workbook is stored in
    another relative path such as "data/WWTP_Analytics_Database.xlsx".

Design principles:
- Dashboard reads calculated KPIs from the analytics database; it does not
  rebuild process formulas.
- Every displayed KPI obtains its unit from KPI_Units / KPI_Definition.
- Engineering targets take priority over historical percentile status.
- Where no approved target exists, historical percentiles are used only as a
  statistical deviation indicator, NOT as a process/compliance limit.
- Process flow: PM1/2/3 -> Mixed PM -> AEQ -> D/N (includes Aeration I)
  -> Aeration Tank -> Secondary/Sludge -> EO.
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
import html

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# 0. CONFIGURATION
# ============================================================

DATA_FILE = "WWTP_Analytics_Database.xlsx"
# Example if stored in a GitHub data folder:
# DATA_FILE = "data/WWTP_Analytics_Database.xlsx"

st.set_page_config(
    page_title="WWTP Performance & Diagnostic Platform",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Status semantics. Colors are intentionally muted and used only for status.
STATUS_STYLE = {
    "normal": {"label": "Normal", "color": "#2E7D32", "bg": "#EDF7EE"},
    "attention": {"label": "Attention", "color": "#C77700", "bg": "#FFF7E6"},
    "abnormal": {"label": "Abnormal", "color": "#C62828", "bg": "#FDEEEE"},
    "no_status": {"label": "No status", "color": "#667085", "bg": "#F2F4F7"},
}

# Statistical direction for historical percentile status.
# This is NOT an engineering limit table. It only controls how unusual values
# are colored when no approved KPI target exists.
DIRECTION_TYPE = {
    # Higher generally means poorer effluent / settling condition
    "EO_COD": "higher_worse",
    "EO_TN": "higher_worse",
    "EO_NH3N": "higher_worse",
    "EO_NO3N": "higher_worse",
    "EO_TP": "higher_worse",
    "EO_SS": "higher_worse",
    "EO_COD_Load_kg_d": "higher_worse",
    "EO_TN_Load_kg_d": "higher_worse",
    "EO_NH3N_Load_kg_d": "higher_worse",
    "EO_TP_Load_kg_d": "higher_worse",
    "EO_SS_Load_kg_d": "higher_worse",
    "EOA_SVI": "higher_worse",
    "SVI_Recalc_mL_g": "higher_worse",
    "SC_Sludge_Blanket": "higher_worse",
    # Very high or very low can both be operationally noteworthy
    "PM1_pH": "two_sided",
    "PM2_pH": "two_sided",
    "PM3_pH": "two_sided",
    "PM_pH": "two_sided",
    "AEQ_pH": "two_sided",
    "DN_Inlet_pH": "two_sided",
    "DN_Outlet_pH": "two_sided",
    "EOA_pH": "two_sided",
    "EO_pH": "two_sided",
    "F_M_Total_Bio_kgCOD_kgMLVSS_d": "two_sided",
    "F_M_Total_Bio_7d": "two_sided",
    "SRT_TSS_7d_Recalc_d": "two_sided",
    "SRT_TSS_14d_Recalc_d": "two_sided",
    "N_COD_g_kg": "two_sided",
    "P_COD_g_kg": "two_sided",
    "MLVSS_MLSS_Ratio_Recalc": "two_sided",
    "RS_Ratio_Excel": "two_sided",
    "Internal_Recycle_Ratio_Recalc": "two_sided",
    "Aeration_Inlet_Temperature": "two_sided",
}

# Friendly names for the KPIs most often shown. Other fields are automatically
# humanized from their database names.
DISPLAY_OVERRIDES = {
    "PM_Total_Flow": "PM Total Flow",
    "PM_COD_Load_Recalc_kg_d": "PM COD Load",
    "COD_Load_Recalc_kg_d": "COD Load",
    "AEQ_TN_Load_kg_d": "AEQ TN Load",
    "AEQ_TP_Load_kg_d": "AEQ TP Load",
    "Total_N_to_Bio_kg_d": "Total N to Biology",
    "Total_P_to_Bio_kg_d": "Total P to Biology",
    "N_COD_g_kg": "N/COD",
    "P_COD_g_kg": "P/COD",
    "EOA_SS": "MLSS",
    "EOA_MLVSS": "MLVSS",
    "MLVSS_MLSS_Ratio_Recalc": "MLVSS/MLSS",
    "F_M_Total_Bio_kgCOD_kgMLVSS_d": "F/M",
    "F_M_Total_Bio_7d": "F/M 7d",
    "SRT_TSS_Daily_Recalc_d": "SRT Daily",
    "SRT_TSS_7d_Recalc_d": "SRT 7d",
    "SRT_TSS_14d_Recalc_d": "SRT 14d",
    "EOA_SVI": "SVI",
    "EOA_30min_Settling": "SV30",
    "SC_Sludge_Blanket": "Sludge Blanket",
    "RS_Ratio_Excel": "Return Sludge Ratio",
    "RS_Flow": "RAS Flow",
    "RS_SS": "RAS SS",
    "ES_Flow": "ES Flow",
    "ES_SS": "ES SS",
    "ES_TSS_Removed_kg_d": "ES TSS Removed",
    "ES_VSS_Removed_kg_d_Est": "ES VSS Removed (Estimated)",
    "ES_Removal_kgVSS_kgCOD_Est": "ES Removal (Estimated)",
    "Aeration_Inlet_Temperature": "Aeration Inlet Temperature",
    "Internal_Recycle_Ratio_Recalc": "Internal Recycle Ratio",
    "TN_Removal_Load_Basis_pct": "TN Removal (Load Basis)",
    "TP_Removal_Load_Basis_pct": "TP Removal (Load Basis)",
    "Overall_COD_Removal_pct": "Overall COD Removal",
}

# Main process-stage fields. Only fields that actually exist are rendered.
PROCESS_STAGES = [
    {
        "name": "PM1 / PM2 / PM3",
        "subtitle": "Source wastewater",
        "core": ["PM1_Flow", "PM2_Flow", "PM3_Flow"],
        "details": [
            "PM1_pH", "PM1_SS", "PM1_COD", "PM1_TN", "PM1_NH3N", "PM1_TP", "PM1_Flow",
            "PM2_pH", "PM2_SS", "PM2_COD", "PM2_TN", "PM2_NH3N", "PM2_TP", "PM2_Flow",
            "PM3_pH", "PM3_SS", "PM3_COD", "PM3_TN", "PM3_NH3N", "PM3_TP", "PM3_Flow",
            "PM1_COD_Load_kg_d", "PM2_COD_Load_kg_d", "PM3_COD_Load_kg_d",
            "PM1_COD_Load_Share_pct", "PM2_COD_Load_Share_pct", "PM3_COD_Load_Share_pct",
        ],
    },
    {
        "name": "Mixed PM",
        "subtitle": "Combined paper-machine wastewater",
        "core": ["PM_Total_Flow", "PM_COD", "PM_COD_Load_Recalc_kg_d", "PM_pH"],
        "details": ["PM_Total_Flow", "PM_pH", "PM_SS", "PM_COD", "PM_TN", "PM_SS_Load_Excel", "PM_COD_Load_Recalc_kg_d"],
    },
    {
        "name": "AEQ",
        "subtitle": "Equalization basin",
        "core": ["AEQ_COD", "COD_Load_Recalc_kg_d", "AEQ_TN", "AEQ_TP"],
        "details": [
            "AEQ_pH", "AEQ_SS", "AEQ_COD", "AEQ_TN", "AEQ_NH3N", "AEQ_TP",
            "AEQ_SS_Load_kg_d", "COD_Load_Recalc_kg_d", "AEQ_TN_Load_kg_d", "AEQ_TP_Load_kg_d", "AEQ_HRT_h",
        ],
    },
    {
        "name": "D/N",
        "subtitle": "D/N basin · includes Aeration I",
        "core": ["N_COD_g_kg", "P_COD_g_kg", "DN_Outlet_COD", "Internal_Recycle_Ratio_Recalc"],
        "details": [
            "DN_Inlet_pH", "DN_Inlet_SS", "DN_Inlet_TN", "DN_Inlet_NO3N", "DN_Inlet_30min_Settling",
            "DN_Outlet_pH", "DN_Outlet_SS", "DN_Outlet_NH3N", "DN_Outlet_NO3N", "DN_Outlet_COD",
            "DN_TN_Inlet_Recalc_mg_L", "N_COD_g_kg", "P_COD_g_kg", "Nutrient_N_kg_d", "Nutrient_P_kg_d",
            "Internal_Recycle_Ratio_Recalc", "Stage_I_HRT_h", "Stage_I_Volumetric_COD_Load_kg_m3_d",
            "Stage_I_Outlet_COD_Load_kg_d", "DN_Apparent_COD_Removal_pct",
        ],
    },
    {
        "name": "Aeration Tank",
        "subtitle": "Independent second aeration stage",
        "core": ["EOA_SS", "F_M_Total_Bio_7d", "SRT_TSS_7d_Recalc_d", "Aeration_Inlet_Temperature"],
        "details": [
            "EOA_pH", "EOA_SS", "EOA_MLVSS", "MLVSS_MLSS_Ratio_Recalc",
            "F_M_Total_Bio_kgCOD_kgMLVSS_d", "F_M_Total_Bio_7d",
            "SRT_TSS_Daily_Recalc_d", "SRT_TSS_7d_Recalc_d", "SRT_TSS_14d_Recalc_d",
            "Aeration_Inlet_Temperature", "Online_MLSS_1", "Online_MLSS_2",
            "Aeration_II_HRT_h", "Aeration_II_Volumetric_COD_Load_kg_m3_d",
        ],
    },
    {
        "name": "Secondary / Sludge",
        "subtitle": "Settling, return and excess sludge",
        "core": ["EOA_SVI", "SC_Sludge_Blanket", "RS_Ratio_Excel", "ES_TSS_Removed_kg_d"],
        "details": [
            "EOA_30min_Settling", "EOA_SVI", "SVI_Recalc_mL_g", "SC_Sludge_Blanket",
            "RS_Flow", "RS_SS", "RS_SS_Online", "RS_Ratio_Excel", "RS_Ratio_Recalc",
            "RAS_TSS_Circulation_kg_d", "ES_Flow", "ES_SS", "ES_SS_Online", "ES_TSS_Removed_kg_d",
            "ES_VSS_Removed_kg_d_Est", "ES_Removal_kgVSS_kgCOD_Est", "DryCake_DrySolids_t_d_Est",
        ],
    },
    {
        "name": "EO",
        "subtitle": "Final effluent",
        "core": ["EO_COD", "EO_TN", "EO_NH3N", "EO_TP"],
        "details": [
            "EO_Flow", "EO_pH", "EO_Conductivity", "EO_BOD5", "EO_SS", "EO_COD", "EO_TN", "EO_NH3N", "EO_NO3N", "EO_TP",
            "EO_SS_Load_kg_d", "EO_COD_Load_kg_d", "EO_TN_Load_kg_d", "EO_NH3N_Load_kg_d", "EO_TP_Load_kg_d",
            "Overall_COD_Removal_pct", "TN_Removal_Load_Basis_pct", "TP_Removal_Load_Basis_pct",
        ],
    },
]

CURRENT_KPIS = [
    "EO_Flow",
    "COD_Load_Recalc_kg_d",
    "Floc_Load_Excel",
    "EOA_SS",
    "F_M_Total_Bio_7d",
    "SRT_TSS_7d_Recalc_d",
    "EOA_SVI",
    "N_COD_g_kg",
    "P_COD_g_kg",
    "EO_COD",
    "EO_TN",
    "EO_NH3N",
    "EO_TP",
]


# Used only to identify the latest day that contains meaningful operating data.
# The source workbook may contain future calendar placeholder rows with no measurements.
VALIDITY_CORE_FIELDS = [
    "EO_Flow", "COD_Load_Recalc_kg_d", "EOA_SS", "EO_COD", "EO_TN", "EOA_SVI"
]

RELATIONSHIP_RECOMMENDATIONS = [
    ("COD_Load_Recalc_kg_d", "F_M_Total_Bio_kgCOD_kgMLVSS_d", "Loading → biomass pressure"),
    ("COD_Load_Recalc_kg_d", "EOA_SS", "Loading → MLSS response"),
    ("F_M_Total_Bio_kgCOD_kgMLVSS_d", "EOA_SVI", "F/M → settling"),
    ("SRT_TSS_7d_Recalc_d", "EO_NH3N", "SRT → nitrification outcome"),
    ("Aeration_Inlet_Temperature", "EO_NH3N", "Temperature → nitrification outcome"),
    ("EOA_SVI", "EO_SS", "Settling → effluent solids"),
    ("N_COD_g_kg", "EO_TN", "N/COD → effluent TN"),
    ("P_COD_g_kg", "EO_TP", "P/COD → effluent TP"),
    ("COD_Load_Recalc_kg_d", "EO_COD", "Loading → effluent COD"),
]


# ============================================================
# 1. STYLE
# ============================================================

st.markdown(
    """
    <style>
      .block-container {padding-top: 1.2rem; padding-bottom: 2.5rem; max-width: 1600px;}
      h1 {font-size: 2rem !important; letter-spacing: -0.02em;}
      h2 {margin-top: 0.7rem !important;}
      [data-testid="stMetric"] {
        background: #FFFFFF;
        border: 1px solid #E4E7EC;
        border-radius: 12px;
        padding: 0.75rem 0.9rem;
      }
      .status-card {
        border: 1px solid #E4E7EC;
        border-radius: 12px;
        padding: 12px 14px;
        background: #FFFFFF;
        min-height: 126px;
      }
      .status-name {font-size: 0.83rem; color: #475467; margin-bottom: 7px; font-weight: 600;}
      .status-value {font-size: 1.45rem; color: #101828; font-weight: 700; line-height: 1.15;}
      .status-unit {font-size: 0.76rem; color: #667085; font-weight: 500; margin-left: 4px;}
      .status-sub {font-size: 0.76rem; color: #667085; margin-top: 7px;}
      .status-pill {display: inline-flex; align-items: center; gap: 6px; border-radius: 999px; padding: 3px 8px; font-size: 0.72rem; font-weight: 650; margin-top: 8px;}
      .status-dot {width: 7px; height: 7px; border-radius: 50%; display: inline-block;}
      .process-card {
        border: 1px solid #D0D5DD;
        border-radius: 13px;
        background: #FFFFFF;
        padding: 12px 13px;
        min-height: 190px;
      }
      .process-title {font-size: 0.98rem; font-weight: 750; color: #101828;}
      .process-subtitle {font-size: 0.70rem; color: #667085; min-height: 32px; margin-top: 2px;}
      .process-metric {border-top: 1px solid #EAECF0; padding-top: 7px; margin-top: 7px;}
      .process-metric-name {font-size: 0.70rem; color: #667085;}
      .process-metric-value {font-size: 0.92rem; font-weight: 700; color: #101828;}
      .flow-arrow {font-size: 1.45rem; color: #98A2B3; text-align: center; padding-top: 77px;}
      .muted-note {color:#667085; font-size:0.82rem;}
      .section-note {border-left: 3px solid #98A2B3; padding-left: 10px; color:#475467; font-size:0.84rem; margin-bottom:0.8rem;}
      .legend-row {display:flex; flex-wrap:wrap; gap:12px; font-size:0.76rem; color:#667085; margin: 4px 0 14px 0;}
      .legend-item {display:inline-flex; align-items:center; gap:5px;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 2. DATA LOADING / METADATA
# ============================================================

@st.cache_data(show_spinner=False)
def load_database(path: str):
    xls_path = Path(path)
    if not xls_path.exists():
        raise FileNotFoundError(
            f"Database not found: {xls_path.resolve()}\n"
            "Change DATA_FILE near the top of wwtp_dashboard.py."
        )

    daily = pd.read_excel(xls_path, sheet_name="Daily_KPI")
    definition = pd.read_excel(xls_path, sheet_name="KPI_Definition")
    units = pd.read_excel(xls_path, sheet_name="KPI_Units")
    targets = pd.read_excel(xls_path, sheet_name="KPI_Target")

    daily["Date"] = pd.to_datetime(daily["Date"], errors="coerce")
    daily = daily.loc[daily["Date"].notna()].sort_values("Date").reset_index(drop=True)

    return daily, definition, units, targets


try:
    df, kpi_def, kpi_units, kpi_targets = load_database(DATA_FILE)
except Exception as exc:
    st.error(str(exc))
    st.stop()

UNIT_MAP: Dict[str, str] = dict(
    zip(kpi_units["KPI_Name"].astype(str), kpi_units["Unit"].fillna("").astype(str))
)
DISPLAY_MAP: Dict[str, str] = dict(
    zip(kpi_units["KPI_Name"].astype(str), kpi_units["Display_Name"].fillna("").astype(str))
)

TARGET_MAP = {}
for _, r in kpi_targets.iterrows():
    TARGET_MAP[str(r["KPI_Name"])] = {
        "lower": pd.to_numeric(r.get("Lower"), errors="coerce"),
        "upper": pd.to_numeric(r.get("Upper"), errors="coerce"),
        "type": str(r.get("Target_Type", "")),
        "source": str(r.get("Source", "")),
    }

NUMERIC_COLS = [
    c for c in df.columns
    if c != "Date" and pd.api.types.is_numeric_dtype(df[c])
]


# ============================================================
# 3. HELPERS
# ============================================================

def humanize(name: str) -> str:
    if name in DISPLAY_OVERRIDES:
        return DISPLAY_OVERRIDES[name]
    d = DISPLAY_MAP.get(name, "")
    if d and d != name:
        return d
    return name.replace("_", " ")


def unit_of(name: str) -> str:
    unit = UNIT_MAP.get(name, "")
    if unit.lower() in {"nan", "none"}:
        return ""
    return unit


def label_with_unit(name: str) -> str:
    unit = unit_of(name)
    return f"{humanize(name)} ({unit})" if unit else humanize(name)


def fmt_value(value, name: str, compact: bool = False) -> str:
    if value is None or pd.isna(value):
        return "—"
    try:
        v = float(value)
    except Exception:
        return str(value)

    if abs(v) >= 1000000:
        text = f"{v/1_000_000:.2f}M" if compact else f"{v:,.0f}"
    elif abs(v) >= 10000:
        text = f"{v/1000:.1f}k" if compact else f"{v:,.0f}"
    elif abs(v) >= 1000:
        text = f"{v:,.0f}"
    elif abs(v) >= 100:
        text = f"{v:,.1f}"
    elif abs(v) >= 10:
        text = f"{v:,.2f}"
    elif abs(v) >= 1:
        text = f"{v:,.2f}"
    else:
        text = f"{v:,.3f}"
    return text


def unit_html(name: str) -> str:
    u = unit_of(name)
    return f'<span class="status-unit">{html.escape(u)}</span>' if u else ""


def available(cols: Iterable[str]) -> List[str]:
    return [c for c in cols if c in df.columns]


def latest_non_null(series_df: pd.DataFrame, col: str):
    if col not in series_df.columns:
        return np.nan, pd.NaT
    s = series_df[["Date", col]].dropna()
    if s.empty:
        return np.nan, pd.NaT
    row = s.iloc[-1]
    return row[col], row["Date"]


def historical_reference(full_df: pd.DataFrame, current_date: pd.Timestamp, mode: str) -> pd.DataFrame:
    hist = full_df.loc[full_df["Date"] < current_date].copy()
    if mode == "Same calendar month":
        month_hist = hist.loc[hist["Date"].dt.month == current_date.month]
        # Fall back to all history if same-month sample is too small.
        if len(month_hist) >= 60:
            hist = month_hist
    return hist


def calc_status(
    col: str,
    value: float,
    ref_df: pd.DataFrame,
) -> Tuple[str, str]:
    """Return (status_key, reason)."""
    if pd.isna(value):
        return "no_status", "No valid latest value"

    # 1) Approved engineering/current target wins.
    if col in TARGET_MAP:
        lo = TARGET_MAP[col]["lower"]
        hi = TARGET_MAP[col]["upper"]
        if pd.notna(lo) and value < lo:
            return "abnormal", f"Below approved target lower bound ({lo:g})"
        if pd.notna(hi) and value > hi:
            return "abnormal", f"Above approved target upper bound ({hi:g})"
        return "normal", "Within approved engineering target"

    # 2) Historical statistical status.
    if col not in ref_df.columns:
        return "no_status", "No historical reference"
    hist = pd.to_numeric(ref_df[col], errors="coerce").dropna()
    if len(hist) < 30:
        return "no_status", "Insufficient historical observations"

    q05, q10, q90, q95 = hist.quantile([0.05, 0.10, 0.90, 0.95]).tolist()
    direction = DIRECTION_TYPE.get(col, "two_sided")

    if direction == "higher_worse":
        if value > q95:
            return "abnormal", f"> historical P95 ({q95:.3g})"
        if value > q90:
            return "attention", f"> historical P90 ({q90:.3g})"
        return "normal", "Within historical ≤P90 range"

    if direction == "lower_worse":
        if value < q05:
            return "abnormal", f"< historical P5 ({q05:.3g})"
        if value < q10:
            return "attention", f"< historical P10 ({q10:.3g})"
        return "normal", "Within historical ≥P10 range"

    # two-sided
    if value < q05 or value > q95:
        return "abnormal", f"Outside historical P5–P95 ({q05:.3g}–{q95:.3g})"
    if value < q10 or value > q90:
        return "attention", f"Outside historical P10–P90 ({q10:.3g}–{q90:.3g})"
    return "normal", "Within historical P10–P90 range"


def seven_day_average(full_df: pd.DataFrame, col: str, latest_date: pd.Timestamp):
    start = latest_date - pd.Timedelta(days=6)
    s = pd.to_numeric(
        full_df.loc[(full_df["Date"] >= start) & (full_df["Date"] <= latest_date), col],
        errors="coerce",
    ).dropna()
    return s.mean() if not s.empty else np.nan


def previous_seven_day_average(full_df: pd.DataFrame, col: str, latest_date: pd.Timestamp):
    end = latest_date - pd.Timedelta(days=7)
    start = end - pd.Timedelta(days=6)
    s = pd.to_numeric(
        full_df.loc[(full_df["Date"] >= start) & (full_df["Date"] <= end), col],
        errors="coerce",
    ).dropna()
    return s.mean() if not s.empty else np.nan


def status_card(col: str, current_date: pd.Timestamp, ref_df: pd.DataFrame) -> str:
    value, value_date = latest_non_null(df.loc[df["Date"] <= current_date], col)
    avg7 = seven_day_average(df, col, current_date)
    prev7 = previous_seven_day_average(df, col, current_date)
    status, reason = calc_status(col, value, ref_df)
    style = STATUS_STYLE[status]

    if pd.notna(avg7) and pd.notna(prev7) and abs(prev7) > 1e-12:
        delta = (avg7 - prev7) / abs(prev7) * 100
        delta_txt = f"7d avg {fmt_value(avg7, col)} · vs prev 7d {delta:+.1f}%"
    elif pd.notna(avg7):
        delta_txt = f"7d avg {fmt_value(avg7, col)}"
    else:
        delta_txt = "7d average unavailable"

    date_note = value_date.strftime("%Y-%m-%d") if pd.notna(value_date) else "no date"
    return f"""
    <div class="status-card" style="border-top:3px solid {style['color']};">
      <div class="status-name">{html.escape(humanize(col))}</div>
      <div class="status-value">{fmt_value(value, col)}{unit_html(col)}</div>
      <div class="status-sub">{html.escape(delta_txt)}<br>Latest valid: {date_note}</div>
      <div class="status-pill" title="{html.escape(reason)}" style="background:{style['bg']}; color:{style['color']};">
        <span class="status-dot" style="background:{style['color']};"></span>{style['label']}
      </div>
    </div>
    """


def process_card(stage: dict, current_date: pd.Timestamp, ref_df: pd.DataFrame) -> str:
    metrics_html = []
    for col in available(stage["core"])[:4]:
        val, _ = latest_non_null(df.loc[df["Date"] <= current_date], col)
        status, reason = calc_status(col, val, ref_df)
        style = STATUS_STYLE[status]
        metrics_html.append(
            f"""
            <div class="process-metric" title="{html.escape(reason)}">
              <div class="process-metric-name"><span class="status-dot" style="background:{style['color']}; margin-right:5px;"></span>{html.escape(humanize(col))}</div>
              <div class="process-metric-value">{fmt_value(val, col, compact=True)} {html.escape(unit_of(col))}</div>
            </div>
            """
        )
    return f"""
      <div class="process-card">
        <div class="process-title">{html.escape(stage['name'])}</div>
        <div class="process-subtitle">{html.escape(stage['subtitle'])}</div>
        {''.join(metrics_html)}
      </div>
    """


def line_chart(
    data: pd.DataFrame,
    col: str,
    title: Optional[str] = None,
    rolling_col: Optional[str] = None,
    height: int = 330,
):
    plot = data[["Date", col] + ([rolling_col] if rolling_col and rolling_col in data.columns else [])].copy()
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=plot["Date"], y=plot[col], mode="lines+markers",
        name=humanize(col),
        line=dict(width=1.5), marker=dict(size=4),
        hovertemplate=f"%{{x|%Y-%m-%d}}<br>{html.escape(humanize(col))}: %{{y:.3g}} {html.escape(unit_of(col))}<extra></extra>",
    ))
    if rolling_col and rolling_col in plot.columns:
        fig.add_trace(go.Scatter(
            x=plot["Date"], y=plot[rolling_col], mode="lines",
            name=humanize(rolling_col), line=dict(width=3),
            hovertemplate=f"%{{x|%Y-%m-%d}}<br>{html.escape(humanize(rolling_col))}: %{{y:.3g}} {html.escape(unit_of(rolling_col))}<extra></extra>",
        ))
    fig.update_layout(
        title=title or humanize(col),
        height=height,
        margin=dict(l=15, r=15, t=48, b=15),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.01, x=0),
        xaxis_title="Date",
        yaxis_title=label_with_unit(col),
    )
    return fig


def paired_lag_data(data: pd.DataFrame, x_col: str, y_col: str, lag_days: int) -> pd.DataFrame:
    """Pair X at day t with Y at day t + lag_days using calendar-date matching."""
    x = data[["Date", x_col]].dropna().copy()
    y = data[["Date", y_col]].dropna().copy()
    y["Date"] = y["Date"] - pd.to_timedelta(lag_days, unit="D")
    merged = x.merge(y, on="Date", how="inner")
    merged = merged.replace([np.inf, -np.inf], np.nan).dropna(subset=[x_col, y_col])
    return merged


def correlations(pairs: pd.DataFrame, x_col: str, y_col: str) -> Tuple[float, float, int]:
    if len(pairs) < 3:
        return np.nan, np.nan, len(pairs)
    pearson = pairs[[x_col, y_col]].corr(method="pearson").iloc[0, 1]
    spearman = pairs[[x_col, y_col]].corr(method="spearman").iloc[0, 1]
    return float(pearson), float(spearman), len(pairs)


def best_lag_table(data: pd.DataFrame, x_col: str, y_col: str, max_lag: int = 14) -> pd.DataFrame:
    rows = []
    for lag in range(0, max_lag + 1):
        pairs = paired_lag_data(data, x_col, y_col, lag)
        p, s, n = correlations(pairs, x_col, y_col)
        rows.append({"Lag_days": lag, "Pearson_r": p, "Spearman_rho": s, "N_pairs": n})
    result = pd.DataFrame(rows)
    return result


def relationship_scatter(data: pd.DataFrame, x_col: str, y_col: str, lag_days: int):
    pairs = paired_lag_data(data, x_col, y_col, lag_days)
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=pairs[x_col],
        y=pairs[y_col],
        mode="markers",
        customdata=pairs[["Date"]].astype(str),
        marker=dict(size=7, opacity=0.65),
        hovertemplate=(
            "X date: %{customdata[0]}<br>"
            + html.escape(humanize(x_col)) + ": %{x:.3g} " + html.escape(unit_of(x_col)) + "<br>"
            + html.escape(humanize(y_col)) + f" (+{lag_days}d): " + "%{y:.3g} " + html.escape(unit_of(y_col))
            + "<extra></extra>"
        ),
        name="Daily pairs",
    ))

    # Simple linear fit for visual context only.
    clean = pairs[[x_col, y_col]].dropna()
    if len(clean) >= 3 and clean[x_col].nunique() >= 2:
        try:
            slope, intercept = np.polyfit(clean[x_col].astype(float), clean[y_col].astype(float), 1)
            xs = np.linspace(clean[x_col].min(), clean[x_col].max(), 100)
            ys = slope * xs + intercept
            fig.add_trace(go.Scatter(x=xs, y=ys, mode="lines", name="Linear fit", line=dict(width=2)))
        except Exception:
            pass

    fig.update_layout(
        height=470,
        margin=dict(l=15, r=15, t=45, b=15),
        xaxis_title=label_with_unit(x_col),
        yaxis_title=f"{label_with_unit(y_col)} at t+{lag_days}d",
        title=f"{humanize(x_col)} → {humanize(y_col)} · lag {lag_days} d",
        legend=dict(orientation="h", yanchor="bottom", y=1.01, x=0),
    )
    return fig, pairs


# ============================================================
# 4. GLOBAL CONTROLS
# ============================================================

raw_max_date = df["Date"].max()
earliest_date = df["Date"].min()

validity_cols = [c for c in VALIDITY_CORE_FIELDS if c in df.columns]
if validity_cols:
    meaningful_mask = df[validity_cols].notna().sum(axis=1) >= min(2, len(validity_cols))
    meaningful_dates = df.loc[meaningful_mask, "Date"]
    latest_date = meaningful_dates.max() if not meaningful_dates.empty else raw_max_date
else:
    latest_date = raw_max_date

st.title("WWTP Performance & Diagnostic Platform")
st.caption(
    f"Meaningful operating-data coverage: {earliest_date:%Y-%m-%d} → {latest_date:%Y-%m-%d} · "
    f"{len(df):,} calendar rows · {len(df.columns)-1:,} data fields"
)
if raw_max_date > latest_date:
    st.caption(
        f"Note: the workbook contains calendar placeholder rows through {raw_max_date:%Y-%m-%d}; "
        f"the dashboard automatically uses {latest_date:%Y-%m-%d} as the latest valid operating date."
    )

with st.sidebar:
    st.header("View settings")

    analysis_date = st.date_input(
        "Status date",
        value=latest_date.date(),
        min_value=earliest_date.date(),
        max_value=latest_date.date(),
    )
    analysis_date = pd.Timestamp(analysis_date)

    recent_window = st.radio(
        "Recent trend window",
        options=[7, 30, 90, 365],
        index=1,
        horizontal=True,
        format_func=lambda x: f"{x}D" if x < 365 else "1Y",
    )

    reference_mode = st.selectbox(
        "Historical status reference",
        ["Same calendar month", "All history"],
        index=0,
        help=(
            "Used only where no approved engineering target exists. "
            "Same-calendar-month comparison reduces seasonality. If fewer than 60 historical observations exist, all history is used."
        ),
    )

    st.divider()
    st.markdown("**Status logic**")
    st.caption("Approved engineering/current target first; otherwise historical percentiles from the selected reference population.")
    st.markdown(
        "🟢 Normal  \n🟠 Attention  \n🔴 Abnormal  \n⚪ No status / insufficient data"
    )
    with st.expander("How are the status colors calculated?", expanded=False):
        st.markdown(
            """
**Priority 1 — Approved target**  
If a KPI exists in `KPI_Target`, that target overrides the percentile rule.

**Priority 2 — Historical percentile status**  
Used only when no approved target exists. The historical population is controlled by **Historical status reference** above.

- **Higher-is-worse KPI**: 🟢 ≤ P90 · 🟠 P90–P95 · 🔴 > P95
- **Lower-is-worse KPI**: 🟢 ≥ P10 · 🟠 P5–P10 · 🔴 < P5
- **Two-sided KPI**: 🟢 P10–P90 · 🟠 P5–P10 or P90–P95 · 🔴 outside P5–P95
- ⚪ means no approved target, no usable historical reference, or fewer than 30 historical observations.

Historical colors indicate **statistical deviation**, not a compliance limit or proven process abnormality.
            """
        )

status_reference = historical_reference(df, analysis_date, reference_mode)
recent_start = analysis_date - pd.Timedelta(days=recent_window - 1)
recent_df = df.loc[(df["Date"] >= recent_start) & (df["Date"] <= analysis_date)].copy()

if recent_df.empty:
    st.warning("No data are available in the selected recent period.")


# ============================================================
# 5. TABS
# ============================================================

tab_current, tab_process, tab_bio, tab_effluent, tab_history = st.tabs([
    "Current Status",
    "Process & Loading",
    "Biology & Sludge",
    "Effluent Performance",
    "Historical Analysis",
])


# ------------------------------------------------------------
# TAB 1 — CURRENT STATUS
# ------------------------------------------------------------
with tab_current:
    st.subheader("Current operating status")
    st.markdown(
        '<div class="section-note">Latest available values are paired with trailing 7-day context. '
        'Colors indicate approved-target status where available; otherwise they indicate statistical deviation from the selected historical reference.</div>',
        unsafe_allow_html=True,
    )

    current_cols = available(CURRENT_KPIS)
    for i in range(0, len(current_cols), 5):
        row_cols = st.columns(min(5, len(current_cols) - i))
        for c_ui, kpi in zip(row_cols, current_cols[i:i+5]):
            with c_ui:
                st.markdown(status_card(kpi, analysis_date, status_reference), unsafe_allow_html=True)

    st.markdown("### Process performance flow")
    st.markdown(
        '<div class="legend-row">'
        '<span class="legend-item"><span class="status-dot" style="background:#2E7D32"></span>Normal</span>'
        '<span class="legend-item"><span class="status-dot" style="background:#C77700"></span>Attention</span>'
        '<span class="legend-item"><span class="status-dot" style="background:#C62828"></span>Abnormal</span>'
        '<span class="legend-item"><span class="status-dot" style="background:#667085"></span>No approved/statistical status</span>'
        '</div>',
        unsafe_allow_html=True,
    )

    # Seven process cards in one scroll-free, responsive Streamlit row.
    cols = st.columns([1.2, 0.15, 1.0, 0.15, 1.0, 0.15, 1.0])
    # First row: source -> mixed -> AEQ -> D/N
    for idx, stage_idx in enumerate([0, 1, 2, 3]):
        colpos = idx * 2
        with cols[colpos]:
            st.markdown(process_card(PROCESS_STAGES[stage_idx], analysis_date, status_reference), unsafe_allow_html=True)
        if idx < 3:
            with cols[colpos + 1]:
                st.markdown('<div class="flow-arrow">→</div>', unsafe_allow_html=True)

    cols2 = st.columns([1.0, 0.15, 1.0, 0.15, 1.0])
    for idx, stage_idx in enumerate([4, 5, 6]):
        colpos = idx * 2
        with cols2[colpos]:
            st.markdown(process_card(PROCESS_STAGES[stage_idx], analysis_date, status_reference), unsafe_allow_html=True)
        if idx < 2:
            with cols2[colpos + 1]:
                st.markdown('<div class="flow-arrow">→</div>', unsafe_allow_html=True)

    with st.expander("View all latest values by process stage", expanded=False):
        stage_choice = st.selectbox("Stage", [s["name"] for s in PROCESS_STAGES], key="stage_detail")
        stage = next(s for s in PROCESS_STAGES if s["name"] == stage_choice)
        rows = []
        for kpi in available(stage["details"]):
            val, dt = latest_non_null(df.loc[df["Date"] <= analysis_date], kpi)
            status, reason = calc_status(kpi, val, status_reference)
            rows.append({
                "KPI": humanize(kpi),
                "Value": val,
                "Unit": unit_of(kpi),
                "Latest valid date": dt,
                "Status": STATUS_STYLE[status]["label"],
                "Status basis": reason,
                "Database field": kpi,
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.markdown(f"### Recent {recent_window}-day trends")
    trend_choices = available([
        "COD_Load_Recalc_kg_d", "EOA_SS", "F_M_Total_Bio_kgCOD_kgMLVSS_d",
        "SRT_TSS_7d_Recalc_d", "EOA_SVI", "EO_COD", "EO_NH3N"
    ])
    chosen = st.selectbox(
        "Trend KPI",
        trend_choices,
        format_func=label_with_unit,
        key="current_trend_kpi",
    )
    rolling_candidate = f"{chosen}_MA7"
    st.plotly_chart(
        line_chart(recent_df, chosen, rolling_col=rolling_candidate if rolling_candidate in df.columns else None),
        use_container_width=True,
        key="current_status_recent_trend",
    )


# ------------------------------------------------------------
# TAB 2 — PROCESS & LOADING
# ------------------------------------------------------------
with tab_process:
    st.subheader("Process & loading")
    st.caption("Follow the wastewater from the paper machines through equalization and biological treatment.")

    st.markdown("### Paper machines → Mixed PM")
    source_flow_cols = available(["PM1_Flow", "PM2_Flow", "PM3_Flow", "PM_Total_Flow"])
    if source_flow_cols:
        fig = go.Figure()
        for c in source_flow_cols:
            fig.add_trace(go.Scatter(
                x=recent_df["Date"], y=recent_df[c], mode="lines",
                name=humanize(c),
                hovertemplate=f"%{{x|%Y-%m-%d}}<br>{humanize(c)}: %{{y:.3g}} {unit_of(c)}<extra></extra>",
            ))
        fig.update_layout(
            height=360, margin=dict(l=15, r=15, t=20, b=15), hovermode="x unified",
            xaxis_title="Date", yaxis_title=f"Flow ({unit_of(source_flow_cols[0])})",
            legend=dict(orientation="h", yanchor="bottom", y=1.01, x=0),
        )
        st.plotly_chart(fig, use_container_width=True, key="process_pm_flow")

    c1, c2 = st.columns(2)
    with c1:
        if "AEQ_COD" in df.columns:
            st.plotly_chart(line_chart(recent_df, "AEQ_COD", "AEQ COD concentration", "AEQ_COD_MA7"), use_container_width=True, key="process_aeq_cod")
    with c2:
        if "COD_Load_Recalc_kg_d" in df.columns:
            st.plotly_chart(line_chart(recent_df, "COD_Load_Recalc_kg_d", "Biological COD load", "COD_Load_Recalc_kg_d_MA7"), use_container_width=True, key="process_cod_load")

    st.markdown("### AEQ → nutrient balance → D/N")
    nutrient_kpis = available([
        "AEQ_TN_Load_kg_d", "AEQ_TP_Load_kg_d", "Nutrient_N_kg_d", "Nutrient_P_kg_d",
        "Total_N_to_Bio_kg_d", "Total_P_to_Bio_kg_d", "N_COD_g_kg", "P_COD_g_kg"
    ])
    selected_nutrient = st.selectbox("Nutrient / loading KPI", nutrient_kpis, format_func=label_with_unit)
    st.plotly_chart(line_chart(recent_df, selected_nutrient), use_container_width=True, key="process_nutrient_trend")

    st.markdown("### Stage measurements")
    stage_metric = st.selectbox(
        "Choose a process measurement",
        available([
            "AEQ_pH", "AEQ_SS", "AEQ_COD", "AEQ_TN", "AEQ_TP", "AEQ_NH3N",
            "DN_Inlet_pH", "DN_Inlet_SS", "DN_Inlet_TN", "DN_Inlet_NO3N",
            "DN_Outlet_pH", "DN_Outlet_SS", "DN_Outlet_NH3N", "DN_Outlet_NO3N", "DN_Outlet_COD",
            "Internal_Recycle_Ratio_Recalc", "Stage_I_HRT_h", "Aeration_II_HRT_h",
        ]),
        format_func=label_with_unit,
    )
    st.plotly_chart(line_chart(recent_df, stage_metric), use_container_width=True, key="process_stage_metric")


# ------------------------------------------------------------
# TAB 3 — BIOLOGY & SLUDGE
# ------------------------------------------------------------
with tab_bio:
    st.subheader("Biology & sludge")
    st.caption("Biomass condition, solids age, settling, return sludge and excess sludge.")

    c1, c2 = st.columns(2)
    with c1:
        if "EOA_SS" in df.columns:
            st.plotly_chart(line_chart(recent_df, "EOA_SS", "MLSS", "EOA_SS_MA7"), use_container_width=True, key="biology_mlss")
    with c2:
        if "EOA_MLVSS" in df.columns:
            st.plotly_chart(line_chart(recent_df, "EOA_MLVSS", "MLVSS", "EOA_MLVSS_MA7"), use_container_width=True, key="biology_mlvss")

    c1, c2 = st.columns(2)
    with c1:
        if "F_M_Total_Bio_kgCOD_kgMLVSS_d" in df.columns:
            st.plotly_chart(
                line_chart(recent_df, "F_M_Total_Bio_kgCOD_kgMLVSS_d", "F/M", "F_M_Total_Bio_kgCOD_kgMLVSS_d_MA7"),
                use_container_width=True,
                key="biology_fm",
            )
    with c2:
        if "SRT_TSS_7d_Recalc_d" in df.columns:
            fig = go.Figure()
            for col in available(["SRT_TSS_Daily_Recalc_d", "SRT_TSS_7d_Recalc_d", "SRT_TSS_14d_Recalc_d"]):
                fig.add_trace(go.Scatter(
                    x=recent_df["Date"], y=recent_df[col], mode="lines",
                    name=humanize(col),
                    hovertemplate=f"%{{x|%Y-%m-%d}}<br>{humanize(col)}: %{{y:.3g}} {unit_of(col)}<extra></extra>",
                ))
            fig.update_layout(
                title="SRT diagnostics", height=330, margin=dict(l=15,r=15,t=48,b=15), hovermode="x unified",
                xaxis_title="Date", yaxis_title=f"SRT ({unit_of('SRT_TSS_7d_Recalc_d')})",
                legend=dict(orientation="h", yanchor="bottom", y=1.01, x=0),
            )
            st.plotly_chart(fig, use_container_width=True, key="biology_srt")

    st.markdown("### Settling & sludge handling")
    sludge_metric = st.selectbox(
        "Sludge KPI",
        available([
            "EOA_SVI", "EOA_30min_Settling", "SC_Sludge_Blanket",
            "RS_Ratio_Excel", "RS_Flow", "RS_SS", "RAS_TSS_Circulation_kg_d",
            "ES_Flow", "ES_SS", "ES_TSS_Removed_kg_d",
            "ES_VSS_Removed_kg_d_Est", "ES_Removal_kgVSS_kgCOD_Est",
            "DryCake_DrySolids_t_d_Est",
        ]),
        format_func=label_with_unit,
    )
    sludge_roll = f"{sludge_metric}_MA7"
    st.plotly_chart(
        line_chart(recent_df, sludge_metric, rolling_col=sludge_roll if sludge_roll in df.columns else None),
        use_container_width=True,
        key="biology_sludge_metric",
    )

    st.info(
        "ES TSS Removed is based on measured ES flow and SS. ES VSS / ES Removal fields marked 'Estimated' use the mixed-liquor MLVSS/MLSS fraction as a proxy and should be interpreted as diagnostics, not direct measurements."
    )


# ------------------------------------------------------------
# TAB 4 — EFFLUENT PERFORMANCE
# ------------------------------------------------------------
with tab_effluent:
    st.subheader("Effluent performance")

    effluent_cards = available(["EO_COD", "EO_TN", "EO_NH3N", "EO_NO3N", "EO_TP", "EO_SS", "EO_pH", "EO_Conductivity"])
    for i in range(0, len(effluent_cards), 4):
        ui_cols = st.columns(min(4, len(effluent_cards)-i))
        for ui, col in zip(ui_cols, effluent_cards[i:i+4]):
            with ui:
                st.markdown(status_card(col, analysis_date, status_reference), unsafe_allow_html=True)

    effluent_metric = st.selectbox(
        "Effluent trend KPI",
        effluent_cards,
        format_func=label_with_unit,
    )
    eff_roll = f"{effluent_metric}_MA7"
    st.plotly_chart(
        line_chart(recent_df, effluent_metric, rolling_col=eff_roll if eff_roll in df.columns else None),
        use_container_width=True,
        key="effluent_metric_trend",
    )

    st.markdown("### Removal performance")
    removal_cols = available(["Overall_COD_Removal_pct", "TN_Removal_Load_Basis_pct", "TP_Removal_Load_Basis_pct"])
    if removal_cols:
        fig = go.Figure()
        for c in removal_cols:
            fig.add_trace(go.Scatter(
                x=recent_df["Date"], y=recent_df[c], mode="lines",
                name=humanize(c),
                hovertemplate=f"%{{x|%Y-%m-%d}}<br>{humanize(c)}: %{{y:.2f}} {unit_of(c)}<extra></extra>",
            ))
        fig.update_layout(
            height=370, margin=dict(l=15,r=15,t=20,b=15), hovermode="x unified",
            xaxis_title="Date", yaxis_title="Removal (%)",
            legend=dict(orientation="h", yanchor="bottom", y=1.01, x=0),
        )
        st.plotly_chart(fig, use_container_width=True, key="effluent_removal")


# ------------------------------------------------------------
# TAB 5 — HISTORICAL ANALYSIS
# ------------------------------------------------------------
with tab_history:
    st.subheader("Historical analysis")
    st.caption("Explore long-term distributions, relationships and time lags. Correlation indicates association, not causation.")

    hist_c1, hist_c2 = st.columns(2)
    with hist_c1:
        hist_start = st.date_input("Historical start", value=earliest_date.date(), min_value=earliest_date.date(), max_value=latest_date.date(), key="hist_start")
    with hist_c2:
        hist_end = st.date_input("Historical end", value=analysis_date.date(), min_value=earliest_date.date(), max_value=latest_date.date(), key="hist_end")

    hist_start_ts = pd.Timestamp(hist_start)
    hist_end_ts = pd.Timestamp(hist_end)
    history_df = df.loc[(df["Date"] >= hist_start_ts) & (df["Date"] <= hist_end_ts)].copy()

    st.markdown("### Historical distribution / operating envelope")
    dist_kpi = st.selectbox(
        "KPI",
        NUMERIC_COLS,
        index=NUMERIC_COLS.index("COD_Load_Recalc_kg_d") if "COD_Load_Recalc_kg_d" in NUMERIC_COLS else 0,
        format_func=label_with_unit,
        key="dist_kpi",
    )
    vals = pd.to_numeric(history_df[dist_kpi], errors="coerce").dropna()
    if len(vals) >= 5:
        qs = vals.quantile([0.05,0.10,0.25,0.50,0.75,0.90,0.95])
        qcols = st.columns(7)
        for ui, (q, v) in zip(qcols, qs.items()):
            with ui:
                st.metric(f"P{int(q*100)}", f"{fmt_value(v, dist_kpi)} {unit_of(dist_kpi)}")

        fig = px.histogram(history_df, x=dist_kpi, nbins=50, marginal="box")
        fig.update_layout(
            height=390, margin=dict(l=15,r=15,t=20,b=15),
            xaxis_title=label_with_unit(dist_kpi), yaxis_title="Daily observations",
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True, key="history_distribution")

    st.markdown("### Relationship Explorer")
    st.markdown(
        '<div class="section-note">Choose any two numeric KPIs. X is interpreted as the upstream/explanatory variable at day t; '
        'Y is evaluated at day t + lag. Axes auto-scale to the selected data.</div>',
        unsafe_allow_html=True,
    )

    rec_labels = [f"{humanize(x)} → {humanize(y)} · {note}" for x,y,note in RELATIONSHIP_RECOMMENDATIONS if x in df.columns and y in df.columns]
    rec_options = ["Custom"] + rec_labels
    rec_choice = st.selectbox("Recommended relationship", rec_options)

    default_x = "COD_Load_Recalc_kg_d" if "COD_Load_Recalc_kg_d" in NUMERIC_COLS else NUMERIC_COLS[0]
    default_y = "EO_COD" if "EO_COD" in NUMERIC_COLS else NUMERIC_COLS[min(1, len(NUMERIC_COLS)-1)]

    if rec_choice != "Custom":
        rec_idx = rec_labels.index(rec_choice)
        valid_recs = [(x,y,n) for x,y,n in RELATIONSHIP_RECOMMENDATIONS if x in df.columns and y in df.columns]
        default_x, default_y, _ = valid_recs[rec_idx]

    c1, c2, c3 = st.columns([1,1,0.7])
    with c1:
        x_col = st.selectbox(
            "X — upstream / explanatory KPI",
            NUMERIC_COLS,
            index=NUMERIC_COLS.index(default_x),
            format_func=label_with_unit,
            key=f"x_{default_x}_{rec_choice}",
        )
    with c2:
        y_col = st.selectbox(
            "Y — downstream / response KPI",
            NUMERIC_COLS,
            index=NUMERIC_COLS.index(default_y),
            format_func=label_with_unit,
            key=f"y_{default_y}_{rec_choice}",
        )
    with c3:
        manual_lag = st.number_input("Lag (days)", min_value=0, max_value=60, value=0, step=1)

    pairs = paired_lag_data(history_df, x_col, y_col, int(manual_lag))
    pearson, spearman, n_pairs = correlations(pairs, x_col, y_col)

    m1, m2, m3 = st.columns(3)
    m1.metric("Pearson r", "—" if pd.isna(pearson) else f"{pearson:.3f}")
    m2.metric("Spearman ρ", "—" if pd.isna(spearman) else f"{spearman:.3f}")
    m3.metric("Paired days", f"{n_pairs:,}")

    scatter, _ = relationship_scatter(history_df, x_col, y_col, int(manual_lag))
    st.plotly_chart(scatter, use_container_width=True, key="history_relationship_scatter")

    st.markdown("### Find best lag")
    lag_max = st.slider("Search lag range", min_value=3, max_value=30, value=14, step=1, format="0–%d days")
    lag_result = best_lag_table(history_df, x_col, y_col, lag_max)
    valid_lag = lag_result.dropna(subset=["Spearman_rho"]).copy()
    valid_lag = valid_lag.loc[valid_lag["N_pairs"] >= 10]

    if valid_lag.empty:
        st.warning("Not enough paired observations to evaluate lag correlation.")
    else:
        best_idx = valid_lag["Spearman_rho"].abs().idxmax()
        best = valid_lag.loc[best_idx]
        best_lag = int(best["Lag_days"])

        b1, b2, b3 = st.columns(3)
        b1.metric("Strongest lag", f"{best_lag} days")
        b2.metric("Spearman ρ", f"{best['Spearman_rho']:.3f}")
        b3.metric("Pearson r", f"{best['Pearson_r']:.3f}")

        lag_fig = go.Figure()
        lag_fig.add_trace(go.Scatter(
            x=lag_result["Lag_days"], y=lag_result["Spearman_rho"],
            mode="lines+markers", name="Spearman ρ",
            hovertemplate="Lag %{x} d<br>Spearman ρ %{y:.3f}<extra></extra>",
        ))
        lag_fig.add_trace(go.Scatter(
            x=lag_result["Lag_days"], y=lag_result["Pearson_r"],
            mode="lines+markers", name="Pearson r",
            hovertemplate="Lag %{x} d<br>Pearson r %{y:.3f}<extra></extra>",
        ))
        lag_fig.add_hline(y=0, line_width=1, line_dash="dot")
        lag_fig.update_layout(
            height=360, margin=dict(l=15,r=15,t=20,b=15),
            xaxis_title="Lag from X to Y (days)", yaxis_title="Correlation coefficient",
            legend=dict(orientation="h", yanchor="bottom", y=1.01, x=0),
        )
        st.plotly_chart(lag_fig, use_container_width=True, key="history_lag_curve")

        st.caption(
            f"Strongest observed association in the selected range: {humanize(x_col)} at day t vs "
            f"{humanize(y_col)} at day t+{best_lag}, Spearman ρ={best['Spearman_rho']:.3f}. "
            "This is an observational association and should be interpreted together with process knowledge."
        )

        best_scatter, best_pairs = relationship_scatter(history_df, x_col, y_col, best_lag)
        st.plotly_chart(best_scatter, use_container_width=True, key="history_best_lag_scatter")

    with st.expander("KPI definition / source / calculation", expanded=False):
        selected_meta = kpi_def.loc[kpi_def["KPI_Name"].isin([x_col, y_col])].copy()
        cols_to_show = [
            c for c in [
                "KPI_Name", "Display_Name", "Module", "Unit", "Source_Column", "Calculation",
                "Calculation_Type", "Definition", "Data_Status", "Definition_Source",
                "Target_Source", "Assumptions", "Notes"
            ] if c in selected_meta.columns
        ]
        st.dataframe(selected_meta[cols_to_show], use_container_width=True, hide_index=True)


# ============================================================
# FOOTER
# ============================================================

st.divider()
st.caption(
    "Status colors are operational aids. Historical percentile status represents statistical deviation, not a compliance or engineering limit. "
    "Approved targets in KPI_Target always take priority. All units are read from KPI_Units / KPI_Definition."
)
