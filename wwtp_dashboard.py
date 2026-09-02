# -*- coding: utf-8 -*-
"""
WWTP Performance & Diagnostic Platform — V1.4
============================================

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
    "SRT": "two_sided",
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
    "SRT": "SRT",
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
        "core": ["EOA_SS", "SRT", "Floc_Load_Excel", "EOA_SVI"],
        "details": [
            "EOA_pH", "EOA_SS", "EOA_MLVSS", "MLVSS_MLSS_Ratio_Recalc",
            "Floc_Load_Excel", "F_M_Total_Bio_kgCOD_kgMLVSS_d", "F_M_Total_Bio_7d",
            "SRT",
            "EOA_30min_Settling", "EOA_SVI", "SVI_Recalc_mL_g",
            "Aeration_Inlet_Temperature", "Online_MLSS_1", "Online_MLSS_2",
            "Aeration_II_HRT_h", "Aeration_II_Volumetric_COD_Load_kg_m3_d",
        ],
    },
    {
        "name": "Secondary / Sludge",
        "subtitle": "Settling, return and excess sludge",
        "core": ["SC_Sludge_Blanket", "RS_Ratio_Excel", "ES_Flow", "ES_TSS_Removed_kg_d"],
        "details": [
            "SC_Sludge_Blanket",
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


# Process schematic nodes are deliberately separated from PROCESS_STAGES.
# PM1/PM2/PM3 are drawn as three real inlet branches, while the detail table
# can still use the grouped source-wastewater stage above.
PROCESS_DIAGRAM_NODES = [
    {
        "id": "PM1", "name": "PM1", "subtitle": "Paper machine wastewater",
        "x": 0.0, "y": 0.80, "w": 0.70, "h": 0.20,
        "core": ["PM1_Flow", "PM1_COD"],
        "hover": ["PM1_pH", "PM1_SS", "PM1_TN", "PM1_TP"],
        "details": ["PM1_Flow", "PM1_pH", "PM1_SS", "PM1_COD", "PM1_TN", "PM1_NH3N", "PM1_TP", "PM1_COD_Load_kg_d"],
    },
    {
        "id": "PM2", "name": "PM2", "subtitle": "Paper machine wastewater",
        "x": 0.0, "y": 0.50, "w": 0.70, "h": 0.20,
        "core": ["PM2_Flow", "PM2_COD"],
        "hover": ["PM2_pH", "PM2_SS", "PM2_TN", "PM2_TP"],
        "details": ["PM2_Flow", "PM2_pH", "PM2_SS", "PM2_COD", "PM2_TN", "PM2_NH3N", "PM2_TP", "PM2_COD_Load_kg_d"],
    },
    {
        "id": "PM3", "name": "PM3", "subtitle": "Paper machine wastewater",
        "x": 0.0, "y": 0.20, "w": 0.70, "h": 0.20,
        "core": ["PM3_Flow", "PM3_COD"],
        "hover": ["PM3_pH", "PM3_SS", "PM3_TN", "PM3_TP"],
        "details": ["PM3_Flow", "PM3_pH", "PM3_SS", "PM3_COD", "PM3_TN", "PM3_NH3N", "PM3_TP", "PM3_COD_Load_kg_d"],
    },
    {
        "id": "PM", "name": "Mixed PM", "subtitle": "Combined PM wastewater",
        "x": 1.15, "y": 0.50, "w": 0.88, "h": 0.32,
        "core": ["PM_Total_Flow", "PM_COD", "PM_COD_Load_Recalc_kg_d"],
        "hover": ["PM_pH", "PM_SS", "PM_TN"],
        "details": ["PM_Total_Flow", "PM_pH", "PM_SS", "PM_COD", "PM_TN", "PM_SS_Load_Excel", "PM_COD_Load_Recalc_kg_d"],
    },
    {
        "id": "AEQ", "name": "AEQ", "subtitle": "Equalization basin",
        "x": 2.35, "y": 0.50, "w": 0.88, "h": 0.32,
        "core": ["AEQ_COD", "AEQ_TN", "AEQ_TP"],
        "hover": ["AEQ_pH", "AEQ_SS", "AEQ_HRT_h", "AEQ_NH3N"],
        "details": ["AEQ_pH", "AEQ_SS", "AEQ_COD", "AEQ_TN", "AEQ_NH3N", "AEQ_TP", "AEQ_SS_Load_kg_d", "COD_Load_Recalc_kg_d", "AEQ_TN_Load_kg_d", "AEQ_TP_Load_kg_d", "AEQ_HRT_h"],
    },
    {
        "id": "FEED", "name": "Biological Feed", "subtitle": "Load & nutrient balance",
        "x": 3.55, "y": 0.50, "w": 0.98, "h": 0.32,
        "core": ["COD_Load_Recalc_kg_d", "N_COD_g_kg", "P_COD_g_kg"],
        "hover": ["Nutrient_N_kg_d", "Nutrient_P_kg_d", "Total_N_to_Bio_kg_d", "Total_P_to_Bio_kg_d"],
        "details": ["COD_Load_Recalc_kg_d", "N_COD_g_kg", "P_COD_g_kg", "Nutrient_N_kg_d", "Nutrient_P_kg_d", "Total_N_to_Bio_kg_d", "Total_P_to_Bio_kg_d"],
    },
    {
        "id": "DN", "name": "D/N", "subtitle": "Anoxic + Aeration I",
        "x": 4.80, "y": 0.50, "w": 0.92, "h": 0.32,
        "core": ["DN_Outlet_COD", "DN_Outlet_NH3N", "DN_Outlet_NO3N"],
        "hover": ["DN_Inlet_TN", "Internal_Recycle_Ratio_Recalc", "Stage_I_HRT_h", "DN_Apparent_COD_Removal_pct"],
        "details": ["DN_Inlet_pH", "DN_Inlet_SS", "DN_Inlet_TN", "DN_Inlet_NO3N", "DN_Inlet_30min_Settling", "DN_Outlet_pH", "DN_Outlet_SS", "DN_Outlet_NH3N", "DN_Outlet_NO3N", "DN_Outlet_COD", "DN_TN_Inlet_Recalc_mg_L", "Internal_Recycle_Ratio_Recalc", "Stage_I_HRT_h", "Stage_I_Volumetric_COD_Load_kg_m3_d", "Stage_I_Outlet_COD_Load_kg_d", "DN_Apparent_COD_Removal_pct"],
    },
    {
        "id": "AER", "name": "Aeration / Biomass", "subtitle": "Aeration II + biomass condition",
        "x": 6.10, "y": 0.50, "w": 1.08, "h": 0.34,
        "core": ["EOA_SS", "SRT", "Floc_Load_Excel"],
        "hover": ["EOA_MLVSS", "EOA_SVI", "F_M_Total_Bio_7d", "Aeration_Inlet_Temperature"],
        "details": ["EOA_pH", "EOA_SS", "EOA_MLVSS", "MLVSS_MLSS_Ratio_Recalc", "Floc_Load_Excel", "F_M_Total_Bio_kgCOD_kgMLVSS_d", "F_M_Total_Bio_7d", "SRT", "EOA_30min_Settling", "EOA_SVI", "SVI_Recalc_mL_g", "Aeration_Inlet_Temperature", "Online_MLSS_1", "Online_MLSS_2", "Aeration_II_HRT_h", "Aeration_II_Volumetric_COD_Load_kg_m3_d"],
    },
    {
        "id": "SEC", "name": "Secondary", "subtitle": "Clarifier & sludge handling",
        "x": 7.55, "y": 0.50, "w": 1.02, "h": 0.32,
        "core": ["SC_Sludge_Blanket", "RS_Ratio_Excel", "ES_TSS_Removed_kg_d"],
        "hover": ["RS_Flow", "RS_SS", "ES_Flow", "ES_SS"],
        "details": ["SC_Sludge_Blanket", "RS_Flow", "RS_SS", "RS_SS_Online", "RS_Ratio_Excel", "RS_Ratio_Recalc", "RAS_TSS_Circulation_kg_d", "ES_Flow", "ES_SS", "ES_SS_Online", "ES_TSS_Removed_kg_d", "ES_VSS_Removed_kg_d_Est", "ES_Removal_kgVSS_kgCOD_Est", "DryCake_DrySolids_t_d_Est"],
    },
    {
        "id": "EO", "name": "EO", "subtitle": "Final effluent",
        "x": 8.90, "y": 0.50, "w": 0.92, "h": 0.32,
        "core": ["EO_COD", "EO_TN", "EO_NH3N"],
        "hover": ["EO_TP", "EO_SS", "EO_pH", "EO_Conductivity"],
        "details": ["EO_Flow", "EO_pH", "EO_Conductivity", "EO_BOD5", "EO_SS", "EO_COD", "EO_TN", "EO_NH3N", "EO_NO3N", "EO_TP", "EO_SS_Load_kg_d", "EO_COD_Load_kg_d", "EO_TN_Load_kg_d", "EO_NH3N_Load_kg_d", "EO_TP_Load_kg_d", "Overall_COD_Removal_pct", "TN_Removal_Load_Basis_pct", "TP_Removal_Load_Basis_pct"],
    },
]

CURRENT_KPIS = [
    "EO_Flow",
    "COD_Load_Recalc_kg_d",
    "Floc_Load_Excel",
    "EOA_SS",
    "F_M_Total_Bio_7d",
    "SRT",
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
    ("SRT", "EO_NH3N", "SRT → nitrification outcome"),
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
      .flow-arrow {font-size: 1.45rem; color: #98A2B3; text-align:center; flex:0 0 28px; align-self:center; padding:0 2px;}
      .process-flow-scroll {width:100%; overflow-x:auto; padding:6px 0 12px 0;}
      .process-flow-track {display:flex; align-items:stretch; gap:6px; min-width:max-content;}
      .process-flow-track .process-card {width:190px; min-width:190px; height:auto;}
      .muted-note {color:#667085; font-size:0.82rem;}
      .section-note {border-left: 3px solid #98A2B3; padding-left: 10px; color:#475467; font-size:0.84rem; margin-bottom:0.8rem;}
      .legend-row {display:flex; flex-wrap:wrap; gap:12px; font-size:0.76rem; color:#667085; margin: 4px 0 14px 0;}
      .legend-item {display:inline-flex; align-items:center; gap:5px;}
      .status-overview {display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:8px; margin:6px 0 16px 0;}
      .status-row {display:flex; align-items:center; justify-content:space-between; gap:10px; border:1px solid #EAECF0; border-radius:10px; padding:8px 10px; background:#FFFFFF;}
      .status-row-left {display:flex; align-items:center; gap:7px; min-width:0;}
      .status-row-name {font-size:0.78rem; color:#344054; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
      .status-row-value {font-size:0.78rem; color:#101828; font-weight:700; white-space:nowrap;}
      .quantile-wrap {border:1px solid #E4E7EC; border-radius:12px; background:#FFFFFF; padding:10px 12px; margin:6px 0 12px 0;}
      .quantile-title {font-size:0.78rem; color:#667085; margin-bottom:8px;}
      .quantile-grid {display:grid; grid-template-columns:repeat(7,1fr); gap:4px;}
      .quantile-cell {text-align:center; padding:5px 2px;}
      .quantile-label {font-size:0.68rem; color:#98A2B3;}
      .quantile-value {font-size:0.82rem; color:#101828; font-weight:700; margin-top:2px;}
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

EXCLUDED_DIAGNOSTIC_KPIS = {
    "SRT_7d_Diagnostic_d",
    "SRT_TSS_Daily_Recalc_d",
    "SRT_TSS_7d_Recalc_d",
    "SRT_TSS_14d_Recalc_d",
}

NUMERIC_COLS = [
    c for c in df.columns
    if c != "Date"
    and c not in EXCLUDED_DIAGNOSTIC_KPIS
    and pd.api.types.is_numeric_dtype(df[c])
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
    """Return compact HTML for one process node.

    Important: keep the HTML unindented. Indented raw HTML inside st.markdown
    can be interpreted by Markdown as a code block, which caused the literal
    <div ...> text seen in the first dashboard version.
    """
    metrics_html = []
    for col in available(stage["core"])[:4]:
        val, _ = latest_non_null(df.loc[df["Date"] <= current_date], col)
        status, reason = calc_status(col, val, ref_df)
        style = STATUS_STYLE[status]
        unit = unit_of(col)
        unit_text = "" if unit in {"", "-"} else f" {html.escape(unit)}"
        metrics_html.append(
            f'<div class="process-metric" title="{html.escape(reason)}">'
            f'<div class="process-metric-name">'
            f'<span class="status-dot" style="background:{style["color"]}; margin-right:5px;"></span>'
            f'{html.escape(humanize(col))}</div>'
            f'<div class="process-metric-value">{fmt_value(val, col, compact=True)}{unit_text}</div>'
            f'</div>'
        )
    return (
        f'<div class="process-card">'
        f'<div class="process-title">{html.escape(stage["name"])}</div>'
        f'<div class="process-subtitle">{html.escape(stage["subtitle"])}</div>'
        f'{"".join(metrics_html)}'
        f'</div>'
    )


def process_flow_html(current_date: pd.Timestamp, ref_df: pd.DataFrame) -> str:
    """Build the complete WWTP process map as one HTML block."""
    parts = ['<div class="process-flow-scroll"><div class="process-flow-track">']
    for i, stage in enumerate(PROCESS_STAGES):
        parts.append(process_card(stage, current_date, ref_df))
        if i < len(PROCESS_STAGES) - 1:
            parts.append('<div class="flow-arrow" aria-hidden="true">→</div>')
    parts.append('</div></div>')
    return ''.join(parts)



def _worst_status_for_metrics(metrics: Iterable[str], current_date: pd.Timestamp, ref_df: pd.DataFrame) -> str:
    """Return the most severe usable status among a node's core metrics."""
    rank = {"no_status": 0, "normal": 1, "attention": 2, "abnormal": 3}
    statuses = []
    history = df.loc[df["Date"] <= current_date]
    for col in available(metrics):
        val, _ = latest_non_null(history, col)
        status, _ = calc_status(col, val, ref_df)
        statuses.append(status)
    if not statuses:
        return "no_status"
    usable = [s for s in statuses if s != "no_status"]
    if not usable:
        return "no_status"
    return max(usable, key=lambda s: rank[s])


def _node_hover_text(node: dict, current_date: pd.Timestamp, ref_df: pd.DataFrame) -> str:
    """Compact hover: only a few supplementary stage measurements.

    The process diagram itself shows the primary KPIs. Hover is deliberately
    limited to at most four extra measurements so it remains readable.
    """
    history = df.loc[df["Date"] <= current_date]
    lines = [f"<b>{html.escape(node['name'])}</b>", html.escape(node['subtitle'])]
    shown = 0
    for col in available(node.get("hover", [])):
        val, _ = latest_non_null(history, col)
        if pd.isna(val):
            continue
        status, _ = calc_status(col, val, ref_df)
        style = STATUS_STYLE[status]
        unit = unit_of(col)
        unit_txt = f" {html.escape(unit)}" if unit not in {"", "-"} else ""
        lines.append(
            f"<span style='color:{style['color']}'>●</span> "
            f"<b>{html.escape(humanize(col))}</b>: {fmt_value(val, col)}{unit_txt}"
        )
        shown += 1
        if shown >= 4:
            break
    if shown == 0:
        lines.append("No additional valid measurements.")
    lines.append("<span style='color:#98A2B3'>Select the stage below for full detail.</span>")
    return "<br>".join(lines)


def _node_visible_text(node: dict, current_date: pd.Timestamp, ref_df: pd.DataFrame) -> str:
    """Primary node text with KPI-level status dots.

    Green remains visible but quiet. Attention / abnormal values are bold so the
    user can identify the responsible KPI without relying on the vessel border.
    """
    history = df.loc[df["Date"] <= current_date]
    lines = [f"<b>{html.escape(node['name'])}</b>"]
    max_metrics = 2 if node["id"] in {"PM1", "PM2", "PM3"} else 3
    for col in available(node.get("core", []))[:max_metrics]:
        val, _ = latest_non_null(history, col)
        if pd.isna(val):
            continue
        status, _ = calc_status(col, val, ref_df)
        style = STATUS_STYLE[status]
        unit = unit_of(col)
        unit_txt = f" {html.escape(unit)}" if unit not in {"", "-"} else ""
        value_txt = f"{fmt_value(val, col, compact=True)}{unit_txt}"
        if status in {"attention", "abnormal"}:
            value_txt = f"<b>{value_txt}</b>"
        lines.append(
            f"<span style='color:{style['color']}'>●</span> "
            f"{html.escape(humanize(col))}: {value_txt}"
        )
    return "<br>".join(lines)


def process_schematic_figure(current_date: pd.Timestamp, ref_df: pd.DataFrame) -> go.Figure:
    """Plotly process schematic with three PM inlet branches and rich hover details."""
    fig = go.Figure()
    nodes = {n["id"]: n for n in PROCESS_DIAGRAM_NODES}

    # Water path. Arrows are intentionally schematic, not a P&ID.
    connections = [
        ("PM1", "PM"), ("PM2", "PM"), ("PM3", "PM"),
        ("PM", "AEQ"), ("AEQ", "FEED"), ("FEED", "DN"),
        ("DN", "AER"), ("AER", "SEC"), ("SEC", "EO"),
    ]
    for source_id, target_id in connections:
        s = nodes[source_id]
        t = nodes[target_id]
        sx = s["x"] + s["w"] / 2
        sy = s["y"]
        tx = t["x"] - t["w"] / 2
        ty = t["y"]
        fig.add_annotation(
            x=tx, y=ty, ax=sx, ay=sy,
            xref="x", yref="y", axref="x", ayref="y",
            showarrow=True, arrowhead=2, arrowsize=1.05, arrowwidth=2.0,
            arrowcolor="#7A9BB8", text="", opacity=0.9,
        )

    # Draw nodes as neutral process equipment. KPI-level dots carry the main
    # status signal; the vessel itself gets only a weak side accent when needed.
    for node in PROCESS_DIAGRAM_NODES:
        status = _worst_status_for_metrics(node["core"], current_date, ref_df)
        style = STATUS_STYLE[status]
        x0, x1 = node["x"] - node["w"] / 2, node["x"] + node["w"] / 2
        y0, y1 = node["y"] - node["h"] / 2, node["y"] + node["h"] / 2

        fig.add_shape(
            type="rect", x0=x0, x1=x1, y0=y0, y1=y1,
            line=dict(color="#D0D5DD", width=1.5),
            fillcolor="#FFFFFF",
            layer="below",
        )
        # Weak stage-level cue: only attention / abnormal gets a colored side bar.
        if status in {"attention", "abnormal"}:
            fig.add_shape(
                type="line", x0=x0, x1=x0, y0=y0, y1=y1,
                line=dict(color=style["color"], width=5),
                layer="above",
            )
        fig.add_annotation(
            x=node["x"], y=node["y"],
            text=_node_visible_text(node, current_date, ref_df),
            showarrow=False,
            align="center",
            font=dict(size=10 if node["id"] in {"PM1", "PM2", "PM3"} else 11, color="#101828"),
        )
        # Transparent marker supplies a proper Plotly hover target over each vessel/node.
        fig.add_trace(go.Scatter(
            x=[node["x"]], y=[node["y"]], mode="markers",
            marker=dict(size=78 if node["id"] in {"PM1", "PM2", "PM3"} else 105, color="rgba(0,0,0,0.001)"),
            hovertemplate=_node_hover_text(node, current_date, ref_df) + "<extra></extra>",
            showlegend=False,
            name=node["name"],
        ))

    # Labels that make the diagram read like a process schematic.
    fig.add_annotation(x=0.0, y=1.02, text="Paper machine sources", showarrow=False, font=dict(size=11, color="#667085"))
    fig.add_annotation(x=5.35, y=0.95, text="Biological treatment", showarrow=False, font=dict(size=11, color="#667085"))
    fig.add_annotation(x=8.20, y=0.95, text="Separation / effluent", showarrow=False, font=dict(size=11, color="#667085"))

    fig.update_xaxes(visible=False, range=[-0.55, 9.50], fixedrange=True)
    fig.update_yaxes(visible=False, range=[0.04, 1.08], fixedrange=True)
    fig.update_layout(
        height=470,
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="white",
        paper_bgcolor="white",
        hoverlabel=dict(bgcolor="white", font_size=12, font_family="Arial", align="left"),
        showlegend=False,
    )
    return fig


def compact_status_overview(cols: Iterable[str], current_date: pd.Timestamp, ref_df: pd.DataFrame) -> str:
    """Render every selected KPI, including green/normal ones, with light visual weight."""
    history = df.loc[df["Date"] <= current_date]
    items = []
    for col in available(cols):
        val, _ = latest_non_null(history, col)
        status, reason = calc_status(col, val, ref_df)
        style = STATUS_STYLE[status]
        unit = unit_of(col)
        unit_txt = f" {html.escape(unit)}" if unit not in {"", "-"} else ""
        items.append(
            f'<div class="status-row" title="{html.escape(reason)}">'
            f'<div class="status-row-left"><span class="status-dot" style="background:{style["color"]}"></span>'
            f'<span class="status-row-name">{html.escape(humanize(col))}</span></div>'
            f'<span class="status-row-value">{fmt_value(val, col, compact=True)}{unit_txt}</span>'
            f'</div>'
        )
    return '<div class="status-overview">' + ''.join(items) + '</div>'


def quantile_strip_html(kpi: str, qs: pd.Series) -> str:
    unit = unit_of(kpi)
    unit_txt = "" if unit in {"", "-"} else f" · {html.escape(unit)}"
    cells = []
    for q, v in qs.items():
        label = "Median" if abs(float(q)-0.50) < 1e-9 else f"P{int(round(float(q)*100))}"
        cells.append(
            f'<div class="quantile-cell"><div class="quantile-label">{label}</div>'
            f'<div class="quantile-value">{fmt_value(v, kpi)}</div></div>'
        )
    return (
        f'<div class="quantile-wrap"><div class="quantile-title">{html.escape(humanize(kpi))}{unit_txt}</div>'
        f'<div class="quantile-grid">{"".join(cells)}</div></div>'
    )


def process_stage_rows(stage: dict, current_date: pd.Timestamp, ref_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    history = df.loc[df["Date"] <= current_date]
    for kpi in available(stage.get("details", [])):
        val, dt = latest_non_null(history, kpi)
        avg7 = seven_day_average(df, kpi, current_date)
        status, reason = calc_status(kpi, val, ref_df)
        rows.append({
            "KPI": humanize(kpi),
            "Latest": val,
            "Unit": unit_of(kpi),
            "7d average": avg7,
            "Latest valid date": dt,
            "Status": STATUS_STYLE[status]["label"],
            "Status basis": reason,
            "Database field": kpi,
        })
    return pd.DataFrame(rows)


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
    st.subheader("Overview")
    st.markdown(
        '<div class="section-note">One-screen plant view. All green, orange and red states remain visible; '
        'the compact overview keeps normal values visible without giving every KPI the same visual weight.</div>',
        unsafe_allow_html=True,
    )

    st.markdown("### Key operating snapshot")
    key_snapshot = available([
        "EO_Flow", "COD_Load_Recalc_kg_d", "Floc_Load_Excel", "EOA_SS",
        "SRT", "EOA_SVI", "EO_COD", "EO_TN"
    ])
    for i in range(0, len(key_snapshot), 4):
        row_cols = st.columns(min(4, len(key_snapshot) - i))
        for c_ui, kpi in zip(row_cols, key_snapshot[i:i+4]):
            with c_ui:
                st.markdown(status_card(kpi, analysis_date, status_reference), unsafe_allow_html=True)

    st.markdown("### Status overview")
    snapshot_set = set(key_snapshot)
    overview_kpis = [k for k in CURRENT_KPIS if k not in snapshot_set]
    if overview_kpis:
        st.markdown(
            compact_status_overview(overview_kpis, analysis_date, status_reference),
            unsafe_allow_html=True,
        )
    st.caption("Snapshot KPIs are not repeated here; this row complements the snapshot with additional process-status indicators.")

    st.markdown("### Process performance flow")
    st.caption(
        "Schematic process view. Primary KPIs stay on the diagram; hover is intentionally limited to a few supplementary measurements. "
        "SVI is shown with Aeration / Biomass because it characterizes activated-sludge settleability; sludge blanket and RAS/ES remain under Secondary."
    )
    st.markdown(
        '<div class="legend-row">'
        '<span class="legend-item"><span class="status-dot" style="background:#2E7D32"></span>Normal / typical</span>'
        '<span class="legend-item"><span class="status-dot" style="background:#C77700"></span>Attention / statistically unusual</span>'
        '<span class="legend-item"><span class="status-dot" style="background:#C62828"></span>Abnormal / highly unusual</span>'
        '<span class="legend-item"><span class="status-dot" style="background:#667085"></span>No approved/statistical status</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    st.plotly_chart(
        process_schematic_figure(analysis_date, status_reference),
        use_container_width=True,
        config={"displayModeBar": False},
        key="current_process_schematic",
    )

    st.markdown("### Stage detail")
    stage_choice = st.selectbox(
        "Process stage",
        [n["name"] for n in PROCESS_DIAGRAM_NODES],
        key="stage_detail_v14",
    )
    stage = next(n for n in PROCESS_DIAGRAM_NODES if n["name"] == stage_choice)
    detail_rows = process_stage_rows(stage, analysis_date, status_reference)
    if not detail_rows.empty:
        show_cols = ["KPI", "Latest", "Unit", "7d average", "Status", "Status basis"]
        st.dataframe(detail_rows[show_cols], use_container_width=True, hide_index=True)

        detail_metric_options = available(stage.get("details", []))
        if detail_metric_options:
            selected_stage_metric = st.selectbox(
                "Stage trend KPI",
                detail_metric_options,
                format_func=label_with_unit,
                key="current_stage_trend_kpi",
            )
            rolling_candidate = f"{selected_stage_metric}_MA7"
            st.plotly_chart(
                line_chart(
                    recent_df, selected_stage_metric,
                    rolling_col=rolling_candidate if rolling_candidate in df.columns else None,
                    height=300,
                ),
                use_container_width=True,
                key="current_stage_detail_trend",
            )


# ------------------------------------------------------------
# TAB 2 — PROCESS & LOADING
# ------------------------------------------------------------
with tab_process:
    st.subheader("Influent & loading")
    st.caption("Hydraulic load → organic load → nutrient balance → D/N feed conditions.")

    st.markdown("### 1 · Hydraulic load")
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
            height=330, margin=dict(l=15, r=15, t=20, b=15), hovermode="x unified",
            xaxis_title="Date", yaxis_title=f"Flow ({unit_of(source_flow_cols[0])})",
            legend=dict(orientation="h", yanchor="bottom", y=1.01, x=0),
        )
        st.plotly_chart(fig, use_container_width=True, key="process_pm_flow")

    hydraulic_kpis = available(["PM_Total_Flow", "EO_Flow", "AEQ_HRT_h", "Stage_I_HRT_h", "Aeration_II_HRT_h", "Total_Bio_HRT_h"])
    if hydraulic_kpis:
        st.markdown(compact_status_overview(hydraulic_kpis, analysis_date, status_reference), unsafe_allow_html=True)

    st.markdown("### 2 · Organic load")
    c1, c2 = st.columns(2)
    with c1:
        if "AEQ_COD" in df.columns:
            st.plotly_chart(line_chart(recent_df, "AEQ_COD", "AEQ COD concentration", "AEQ_COD_MA7"), use_container_width=True, key="process_aeq_cod")
    with c2:
        if "COD_Load_Recalc_kg_d" in df.columns:
            st.plotly_chart(line_chart(recent_df, "COD_Load_Recalc_kg_d", "Biological COD load", "COD_Load_Recalc_kg_d_MA7"), use_container_width=True, key="process_cod_load")

    st.markdown("### 3 · Nutrient balance / biological feed")
    nutrient_kpis = available([
        "N_COD_g_kg", "P_COD_g_kg", "Nutrient_N_kg_d", "Nutrient_P_kg_d",
        "Total_N_to_Bio_kg_d", "Total_P_to_Bio_kg_d"
    ])
    if nutrient_kpis:
        st.markdown(compact_status_overview(nutrient_kpis, analysis_date, status_reference), unsafe_allow_html=True)
        selected_nutrient = st.selectbox("Nutrient / loading trend", nutrient_kpis, format_func=label_with_unit, key="process_nutrient_select")
        st.plotly_chart(line_chart(recent_df, selected_nutrient), use_container_width=True, key="process_nutrient_trend")

    st.markdown("### 4 · D/N stage response")
    dn_kpis = available(["DN_Outlet_COD", "DN_Outlet_NH3N", "DN_Outlet_NO3N", "Internal_Recycle_Ratio_Recalc", "DN_Apparent_COD_Removal_pct"])
    if dn_kpis:
        st.markdown(compact_status_overview(dn_kpis, analysis_date, status_reference), unsafe_allow_html=True)
        stage_metric = st.selectbox("D/N trend KPI", dn_kpis, format_func=label_with_unit, key="process_stage_metric_select")
        st.plotly_chart(line_chart(recent_df, stage_metric), use_container_width=True, key="process_stage_metric")


# ------------------------------------------------------------
# TAB 3 — BIOLOGY & SLUDGE
# ------------------------------------------------------------
with tab_bio:
    st.subheader("Biology & sludge")
    st.caption("Biomass quantity → loading → sludge age → settleability → clarifier / solids handling.")

    st.markdown("### 1 · Biomass quantity")
    biomass = available(["EOA_SS", "EOA_MLVSS", "MLVSS_MLSS_Ratio_Recalc"])
    if biomass:
        st.markdown(compact_status_overview(biomass, analysis_date, status_reference), unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if "EOA_SS" in df.columns:
            st.plotly_chart(line_chart(recent_df, "EOA_SS", "MLSS", "EOA_SS_MA7"), use_container_width=True, key="biology_mlss")
    with c2:
        if "EOA_MLVSS" in df.columns:
            st.plotly_chart(line_chart(recent_df, "EOA_MLVSS", "MLVSS", "EOA_MLVSS_MA7"), use_container_width=True, key="biology_mlvss")

    st.markdown("### 2 · Loading on biomass")
    loading_kpis = available(["Floc_Load_Excel", "F_M_Total_Bio_kgCOD_kgMLVSS_d", "F_M_Total_Bio_7d"])
    if loading_kpis:
        st.markdown(compact_status_overview(loading_kpis, analysis_date, status_reference), unsafe_allow_html=True)
        loading_metric = st.selectbox("Loading KPI trend", loading_kpis, format_func=label_with_unit, key="biology_loading_select")
        loading_roll = f"{loading_metric}_MA7"
        st.plotly_chart(line_chart(recent_df, loading_metric, rolling_col=loading_roll if loading_roll in df.columns else None), use_container_width=True, key="biology_fm")
    st.caption("Floc Load is the existing plant KPI. Calculated F/M is a diagnostic reconstruction and may use a different biomass / effective-volume basis.")

    st.markdown("### 3 · Sludge age")
    if "SRT" in df.columns:
        srt_hist = df.loc[df["Date"] <= analysis_date, ["Date", "SRT"]].copy()
        srt_hist["SRT_MA7"] = pd.to_numeric(srt_hist["SRT"], errors="coerce").rolling(7, min_periods=3).mean()
        srt_hist["SRT_MA14"] = pd.to_numeric(srt_hist["SRT"], errors="coerce").rolling(14, min_periods=5).mean()
        srt_plot = srt_hist.loc[srt_hist["Date"] >= recent_start].copy()

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=srt_plot["Date"], y=srt_plot["SRT"], mode="lines+markers",
            name="Recorded SRT", line=dict(width=1.5), marker=dict(size=4),
            hovertemplate=f"%{{x|%Y-%m-%d}}<br>Recorded SRT: %{{y:.2f}} {unit_of('SRT')}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=srt_plot["Date"], y=srt_plot["SRT_MA7"], mode="lines",
            name="SRT MA7", line=dict(width=2.5),
            hovertemplate=f"%{{x|%Y-%m-%d}}<br>SRT MA7: %{{y:.2f}} {unit_of('SRT')}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=srt_plot["Date"], y=srt_plot["SRT_MA14"], mode="lines",
            name="SRT MA14", line=dict(width=2),
            hovertemplate=f"%{{x|%Y-%m-%d}}<br>SRT MA14: %{{y:.2f}} {unit_of('SRT')}<extra></extra>",
        ))
        fig.update_layout(
            title="Recorded SRT and moving averages", height=350,
            margin=dict(l=15, r=15, t=75, b=15), hovermode="x unified",
            xaxis_title="Date", yaxis_title=f"SRT ({unit_of('SRT')})",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        )
        st.plotly_chart(fig, use_container_width=True, key="biology_srt")
        st.caption("SRT uses the cleaned recorded source field. MA7 and MA14 are simple moving averages of that recorded SRT; reconstructed inventory/loss SRT values are not used in the dashboard.")

    st.markdown("### 4 · Settleability")
    settle_kpis = available(["EOA_SVI", "EOA_30min_Settling", "Aeration_Inlet_Temperature"])
    if settle_kpis:
        st.markdown(compact_status_overview(settle_kpis, analysis_date, status_reference), unsafe_allow_html=True)
        settle_metric = st.selectbox("Settleability / condition trend", settle_kpis, format_func=label_with_unit, key="biology_settle_select")
        settle_roll = f"{settle_metric}_MA7"
        st.plotly_chart(line_chart(recent_df, settle_metric, rolling_col=settle_roll if settle_roll in df.columns else None), use_container_width=True, key="biology_settle_metric")

    st.markdown("### 5 · Secondary clarifier & solids handling")
    sludge_kpis = available([
        "SC_Sludge_Blanket", "RS_Ratio_Excel", "RS_Flow", "RS_SS",
        "ES_Flow", "ES_SS", "ES_TSS_Removed_kg_d", "DryCake_DrySolids_t_d_Est"
    ])
    if sludge_kpis:
        st.markdown(compact_status_overview(sludge_kpis, analysis_date, status_reference), unsafe_allow_html=True)
        sludge_metric = st.selectbox("Clarifier / sludge KPI trend", sludge_kpis, format_func=label_with_unit, key="biology_sludge_select")
        sludge_roll = f"{sludge_metric}_MA7"
        st.plotly_chart(line_chart(recent_df, sludge_metric, rolling_col=sludge_roll if sludge_roll in df.columns else None), use_container_width=True, key="biology_sludge_metric")

    st.info(
        "SVI is grouped with biomass settleability because it characterizes activated-sludge settling behavior. "
        "Sludge blanket, RAS and ES are grouped under the secondary clarifier / solids-handling response. "
        "Estimated ES-VSS indicators remain diagnostic rather than direct measurements."
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
    st.caption(
        "Nitrogen performance is shown first because nitrogen removal is the primary biological objective. "
        "Stage-wise TN removal is not fabricated where a reliable D/N-outlet TN series is unavailable."
    )

    # ---- Nitrogen performance ----
    st.markdown("#### Nitrogen performance")
    n_summary = available(["TN_Removal_Load_Basis_pct", "EO_TN", "EO_NH3N", "EO_NO3N"])
    if n_summary:
        n_cols = st.columns(len(n_summary))
        for ui, col in zip(n_cols, n_summary):
            with ui:
                st.markdown(status_card(col, analysis_date, status_reference), unsafe_allow_html=True)

    n_left, n_right = st.columns([1, 1.25])
    with n_left:
        if "TN_Removal_Load_Basis_pct" in df.columns:
            st.plotly_chart(
                line_chart(recent_df, "TN_Removal_Load_Basis_pct", title="Overall TN removal"),
                use_container_width=True,
                key="effluent_tn_removal_trend",
            )
    with n_right:
        n_species = available(["EO_TN", "EO_NH3N", "EO_NO3N"])
        if n_species:
            fig_n = go.Figure()
            for c in n_species:
                fig_n.add_trace(go.Scatter(
                    x=recent_df["Date"], y=recent_df[c], mode="lines", name=humanize(c),
                    hovertemplate=f"%{{x|%Y-%m-%d}}<br>{html.escape(humanize(c))}: %{{y:.3g}} {html.escape(unit_of(c))}<extra></extra>",
                ))
            fig_n.update_layout(
                title=dict(text="Effluent nitrogen species", x=0.0, xanchor="left", y=0.98, yanchor="top"),
                height=370,
                margin=dict(l=55, r=25, t=100, b=55),
                hovermode="x unified",
                xaxis=dict(title="Date", automargin=True),
                yaxis=dict(title="Concentration (mg/L)", automargin=True),
                legend=dict(
                    orientation="h",
                    yanchor="bottom", y=1.02,
                    xanchor="left", x=0,
                    title_text="",
                    itemwidth=55,
                ),
            )
            st.plotly_chart(fig_n, use_container_width=True, key="effluent_n_species_trend")

    # ---- COD removal by stage ----
    st.markdown("#### COD removal by stage")
    cod_stage_cols = available([
        "DN_Apparent_COD_Removal_pct",
        "Post_DN_to_EO_COD_Removal_pct",
        "Overall_COD_Removal_pct",
    ])
    if cod_stage_cols:
        cod_cards = st.columns(len(cod_stage_cols))
        for ui, col in zip(cod_cards, cod_stage_cols):
            history = df.loc[df["Date"] <= analysis_date]
            val, _ = latest_non_null(history, col)
            with ui:
                st.metric(label_with_unit(col), fmt_value(val, col) if pd.notna(val) else "—")

        fig_cod = go.Figure()
        for c in cod_stage_cols:
            fig_cod.add_trace(go.Scatter(
                x=recent_df["Date"], y=recent_df[c], mode="lines", name=humanize(c),
                hovertemplate=f"%{{x|%Y-%m-%d}}<br>{html.escape(humanize(c))}: %{{y:.2f}} {html.escape(unit_of(c))}<extra></extra>",
            ))
        fig_cod.update_layout(
            height=360, margin=dict(l=15,r=15,t=20,b=15), hovermode="x unified",
            xaxis_title="Date", yaxis_title="COD removal (%)",
            legend=dict(orientation="h", yanchor="bottom", y=1.01, x=0),
        )
        st.plotly_chart(fig_cod, use_container_width=True, key="effluent_cod_stage_removal")
        st.caption(
            "D/N COD removal is labelled apparent because internal recycle, external carbon and load-basis differences can affect the observed stage balance."
        )


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

    st.markdown("### KPI history & operating envelope")
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
        st.markdown(quantile_strip_html(dist_kpi, qs), unsafe_allow_html=True)

        fig = px.histogram(history_df, x=dist_kpi, nbins=50, marginal="box")
        fig.update_layout(
            height=390, margin=dict(l=15,r=15,t=20,b=15),
            xaxis_title=label_with_unit(dist_kpi), yaxis_title="Daily observations",
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True, key="history_distribution")

    st.markdown("### Relationship & lag explorer")
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
