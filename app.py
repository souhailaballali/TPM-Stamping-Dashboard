"""
╔══════════════════════════════════════════════════════════════════════════╗
║  TE Connectivity — Stamping Department                                   ║
║  TPM Maintenance KPI Dashboard — Full Version                       ║
║  Bruderer Presses S-001 → S-006 + Peripherals                         ║
║                                                                          ║
║  INSTALLATION:                                                           ║
║    pip install streamlit plotly pandas openpyxl numpy kaleido           ║
║                                                                          ║
║  RUN:                                                                    ║
║    streamlit run app.py                                                  ║
╚══════════════════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, date, timedelta
import io
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TE Connectivity — TPM Dashboard",
    page_icon="🔩",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────────────────────────────────────
#  COLONNES SOURCE (noms exacts du fichier Hydra)
# ─────────────────────────────────────────────────────────────────────────────
COL_MACHINE   = "machine_id"
COL_STATUS    = "machine_status_name"
COL_DATE      = "plant_shift_date"
COL_MTTR      = "Sum of mttr_workcenter_numerator_seconds_quantity"
COL_MTBF      = "Sum of mtbf_numerator_seconds_quantity"
COL_PROD      = "hydra_bmk_production_status_name"   # optionnel
REQUIRED_COLS = [COL_MACHINE, COL_STATUS, COL_MTTR, COL_MTBF]

# ─────────────────────────────────────────────────────────────────────────────
#  COULEURS TE CONNECTIVITY
# ─────────────────────────────────────────────────────────────────────────────
TE_ORANGE  = "#E8650A"
TE_ORANGE2 = "#F0934A"
TE_ORANGE3 = "#F5B87A"
TE_ORANGE4 = "#FAD9B5"
TE_DARK    = "#C04D05"
TE_BLACK   = "#1C1C1C"
TE_NAVY    = "#1B2A4A"
TE_BROWN   = "#A07858"
TE_BG      = "#F7F4F0"
TE_WHITE   = "#FFFFFF"
TE_GREEN   = "#27AE60"
TE_RED     = "#C0392B"
TE_AMBER   = "#E67E22"

PALETTE = [TE_ORANGE, TE_NAVY, TE_RED, "#8E44AD", TE_GREEN, "#16A085", "#D4AC0D", TE_ORANGE2]

# ─────────────────────────────────────────────────────────────────────────────
#  CSS GLOBAL
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@300;400;500;600;700;800&family=Barlow+Condensed:wght@400;600;700;800&family=JetBrains+Mono:wght@300;400;500&display=swap');
@import url('https://fonts.googleapis.com/icon?family=Material+Icons');

/* ── GLOBAL ── */
html, body, .stApp {{
    background-color: {TE_BG} !important;
    font-family: 'Barlow', sans-serif;
}}
#MainMenu, footer, header {{ visibility: hidden; }}
.block-container {{ padding-top: 0 !important; max-width: 100% !important; }}

/* ── MASQUER keyboard_double_arrow — remplacer par signe CSS ── */
[data-testid="collapsedControl"] span {{
    font-size: 0 !important;
    color: transparent !important;
    visibility: hidden !important;
    width: 0 !important;
    height: 0 !important;
    display: none !important;
}}

/* ── SIDEBAR — toujours visible + thème sombre ── */
[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, {TE_BLACK} 0%, #2A1A0A 100%) !important;
    border-right: 3px solid {TE_ORANGE} !important;
    transform: none !important;
    left: 0 !important;
    display: block !important;
    visibility: visible !important;
    min-width: 260px !important;
    max-width: 320px !important;
}}
[data-testid="stSidebar"] * {{
    color: #F0E8DF !important;
    font-family: 'Barlow', sans-serif !important;
}}
[data-testid="stSidebar"] > div:first-child {{ padding: 18px 16px !important; }}
[data-testid="stSidebar"] hr {{ border-color: #3D2A18 !important; }}
[data-testid="stSidebar"] h3 {{
    color: {TE_ORANGE} !important;
    font-family: 'Barlow Condensed', sans-serif !important;
    font-size: 11px !important; font-weight: 700 !important;
    letter-spacing: 3px !important; text-transform: uppercase !important;
}}
[data-testid="stSidebar"] [data-testid="stMultiSelect"] > div > div {{
    background-color: #2C1F14 !important;
    border: 1px solid {TE_ORANGE} !important;
    border-radius: 6px !important;
}}
/* Tags machines — orange */
[data-testid="stSidebar"] [data-testid="stMultiSelect"] span[data-baseweb="tag"] {{
    background-color: {TE_ORANGE} !important;
    border-radius: 4px !important;
}}
[data-testid="stSidebar"] [data-testid="stMultiSelect"] span[data-baseweb="tag"] span {{
    color: white !important;
    font-weight: 700 !important;
    font-size: 11px !important;
}}
[data-testid="stSidebar"] [data-testid="stMultiSelect"] span[data-baseweb="tag"] [data-testid="stMultiSelectRemoveButton"],
[data-testid="stSidebar"] [data-testid="stMultiSelect"] span[data-baseweb="tag"] svg {{
    color: white !important;
    fill: white !important;
    stroke: white !important;
}}
/* File uploader — thème marron foncé */
section[data-testid="stSidebar"] div[data-testid="stFileUploadDropzone"] {{
    background: #FFF8F2 !important;
    border: 1.5px dashed {TE_ORANGE} !important;
    border-radius: 10px !important;
}}
section[data-testid="stSidebar"] div[data-testid="stFileUploadDropzone"] * {{
    color: #2e1808 !important;
    opacity: 1 !important;
}}
section[data-testid="stSidebar"] div[data-testid="stFileUploadDropzone"] p,
section[data-testid="stSidebar"] div[data-testid="stFileUploadDropzone"] span,
section[data-testid="stSidebar"] div[data-testid="stFileUploadDropzone"] small {{
    color: #2e1808 !important;
    opacity: 1 !important;
    font-weight: 700 !important;
}}
section[data-testid="stSidebar"] div[data-testid="stFileUploader"] div[class*="uploadedFile"] {{
    background-color: #2C1A0E !important;
    border: 1px solid {TE_ORANGE} !important;
    border-radius: 6px !important;
}}

/* ── HEADER BANNER ── */
.te-header {{
    background: linear-gradient(135deg, {TE_BLACK} 0%, #2A1A0A 60%, #3D2508 100%);
    border-radius: 14px; padding: 26px 36px; margin-bottom: 20px;
    border-left: 6px solid {TE_ORANGE};
    box-shadow: 0 8px 32px rgba(232,101,10,0.18);
    position: relative; overflow: hidden;
}}
.te-header::after {{
    content: ''; position: absolute; top: -40px; right: -40px;
    width: 180px; height: 180px;
    background: radial-gradient(circle, rgba(232,101,10,0.15) 0%, transparent 70%);
    border-radius: 50%;
}}
.te-header-tag {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 9px; font-weight: 700; letter-spacing: 3px;
    text-transform: uppercase; color: {TE_ORANGE}; margin-bottom: 6px;
    display: flex; align-items: center; gap: 8px;
}}
.te-header-tag::before {{
    content: ''; width: 20px; height: 2px;
    background: {TE_ORANGE}; border-radius: 1px;
}}
.te-header-title {{
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 38px; font-weight: 800; color: {TE_WHITE};
    line-height: 1.0; margin-bottom: 4px; text-transform: uppercase;
    letter-spacing: 0.5px;
}}
.te-header-title span {{ color: {TE_ORANGE}; }}
.te-header-sub {{
    font-size: 13px; color: {TE_BROWN}; font-weight: 400; margin-bottom: 10px;
}}
.te-header-badge {{
    display: inline-flex; align-items: center; gap: 6px;
    background: rgba(232,101,10,0.15); border: 1px solid rgba(232,101,10,0.4);
    color: {TE_ORANGE2}; font-size: 11px; font-weight: 600;
    letter-spacing: 1px; text-transform: uppercase;
    border-radius: 20px; padding: 4px 14px;
}}
.te-header-right {{
    display: flex; flex-direction: column; align-items: flex-end; gap: 6px;
}}
.te-live {{
    display: flex; align-items: center; gap: 6px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 10px; letter-spacing: 1px; color: {TE_GREEN};
    background: rgba(39,174,96,0.12); border: 1px solid rgba(39,174,96,0.3);
    border-radius: 20px; padding: 5px 14px;
}}
.te-live-dot {{
    width: 7px; height: 7px; border-radius: 50%; background: {TE_GREEN};
    animation: blink 2s infinite;
}}
@keyframes blink {{
    0%,100% {{ opacity:1; transform:scale(1); }}
    50%      {{ opacity:0.3; transform:scale(1.5); }}
}}

/* ── STATUS BAR ── */
.te-statusbar {{
    display: flex; align-items: center; gap: 18px;
    background: {TE_WHITE}; border: 1px solid #EDE0D4; border-radius: 10px;
    padding: 10px 20px; margin-bottom: 20px;
    font-size: 13px; color: #7A6050; flex-wrap: wrap;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}}
.te-statusbar strong {{ color: {TE_BLACK}; font-weight: 600; }}
.te-sep {{ width: 1px; height: 16px; background: #E0D0C0; flex-shrink: 0; }}
.te-statusbar-item {{ display: flex; align-items: center; gap: 5px; }}
.te-dot-green {{
    width: 8px; height: 8px; background: {TE_GREEN}; border-radius: 50%;
    box-shadow: 0 0 6px rgba(39,174,96,0.5); flex-shrink: 0;
}}

/* ── KPI CARDS ── */
.kpi-card {{
    background: {TE_BLACK}; border-radius: 12px; padding: 22px;
    box-shadow: 0 4px 24px rgba(0,0,0,0.30); border: 1px solid #2E2E2E;
    position: relative; overflow: hidden;
    transition: transform 0.2s, box-shadow 0.2s;
}}
.kpi-card:hover {{
    transform: translateY(-3px);
    box-shadow: 0 8px 32px rgba(232,101,10,0.25);
}}
.kpi-card::before {{
    content: ''; position: absolute; top: 0; left: 0; right: 0; height: 4px;
    background: linear-gradient(90deg, {TE_ORANGE}, {TE_ORANGE3});
    border-radius: 12px 12px 0 0;
}}
.kpi-icon {{
    width: 40px; height: 40px;
    background: linear-gradient(135deg, #2C1A0A, #3D2510);
    border-radius: 10px; display: flex; align-items: center;
    justify-content: center; font-size: 20px; margin-bottom: 14px;
    border: 1px solid #4A2A10;
}}
.kpi-label {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 9px; font-weight: 700; letter-spacing: 2.5px;
    text-transform: uppercase; color: {TE_ORANGE}; margin-bottom: 4px;
}}
.kpi-value {{
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 38px; font-weight: 700; color: {TE_WHITE};
    line-height: 1; margin-bottom: 4px; letter-spacing: -0.5px;
}}
.kpi-divider {{
    width: 28px; height: 3px;
    background: linear-gradient(90deg, {TE_ORANGE}, {TE_ORANGE3});
    border-radius: 2px; margin: 6px 0;
}}
.kpi-unit {{ font-size: 11px; color: {TE_BROWN}; font-weight: 500; }}

/* ── SECTION LABEL ── */
.te-section {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 9px; font-weight: 700; letter-spacing: 3px;
    text-transform: uppercase; color: {TE_ORANGE};
    margin: 24px 0 12px 0;
    display: flex; align-items: center; gap: 10px;
}}
.te-section::before {{
    content: ''; width: 20px; height: 2px;
    background: {TE_ORANGE}; border-radius: 1px;
}}
.te-section::after {{
    content: ''; flex: 1; height: 1px;
    background: linear-gradient(90deg, #F0D0B0, transparent);
}}

/* ── CHART CARDS ── */
.chart-card {{
    background: {TE_WHITE}; border-radius: 12px;
    padding: 18px 18px 8px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.06);
    border: 1px solid #EDE0D4; margin-bottom: 18px;
}}
.chart-header {{
    display: flex; align-items: center; gap: 10px;
    margin-bottom: 12px; padding-bottom: 10px;
    border-bottom: 1px solid #F0E4D8;
}}
.chart-dot {{
    width: 10px; height: 10px; background: {TE_ORANGE};
    border-radius: 50%; flex-shrink: 0;
}}
.chart-title {{
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 15px; font-weight: 700; color: {TE_BLACK};
    letter-spacing: 0.5px; text-transform: uppercase;
}}

/* ── QUADRANT CARDS ── */
.quad-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-top: 12px; }}
.quad {{ padding: 11px 14px; border-radius: 8px; border: 1px solid; }}
.quad h5 {{ font-family:'Barlow Condensed',sans-serif; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:3px; }}
.quad p  {{ font-size:10px; line-height:1.5; margin:0; }}
.q-good   {{ background:#eafaf1; border-color:#a9dfbf; }} .q-good h5   {{ color:#1e8449; }} .q-good p   {{ color:#145a32; }}
.q-watch  {{ background:#eaf2ff; border-color:#aed6f1; }} .q-watch h5  {{ color:#1a5276; }} .q-watch p  {{ color:#1a3d6d; }}
.q-warn   {{ background:#fef9e7; border-color:#f9e79f; }} .q-warn h5   {{ color:#d68910; }} .q-warn p   {{ color:#7d6608; }}
.q-crit   {{ background:#fdf2f2; border-color:#e8a0a0; }} .q-crit h5   {{ color:#c0392b; }} .q-crit p   {{ color:#7b241c; }}

/* ── INSIGHT BOX ── */
.te-insight {{
    background: #FEF0E1; border: 1px solid #FAC98A;
    border-left: 4px solid {TE_ORANGE}; border-radius: 8px;
    padding: 12px 16px; font-size: 12px; color: #4A3020; line-height: 1.65;
    margin-top: 12px;
}}
.te-insight strong {{ color: {TE_BLACK}; }}
.te-insight-crit {{
    background: #fdf2f2; border: 1px solid #e8a0a0;
    border-left: 4px solid {TE_RED}; border-radius: 8px;
    padding: 12px 16px; font-size: 12px; color: #4A3020; line-height: 1.65;
    margin-top: 12px;
}}
.te-insight-ok {{
    background: #eafaf1; border: 1px solid #a9dfbf;
    border-left: 4px solid {TE_GREEN}; border-radius: 8px;
    padding: 12px 16px; font-size: 12px; color: #4A3020; line-height: 1.65;
    margin-top: 12px;
}}

/* ── DEMO BAR ── */
.te-demo {{
    background: linear-gradient(90deg, {TE_ORANGE} 0%, {TE_DARK} 100%);
    border-radius: 9px; padding: 11px 18px; margin-bottom: 18px;
    color: white; font-size: 13px; font-weight: 500;
    box-shadow: 0 3px 14px rgba(232,101,10,0.28);
    display: flex; align-items: center; gap: 10px;
}}

/* ── WELCOME SCREEN ── */
.welcome-card {{
    background: {TE_WHITE}; border: 2px dashed #F0C8A0;
    border-radius: 16px; padding: 52px 40px; text-align: center;
    margin: 40px auto; max-width: 540px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.06);
}}
.welcome-icon {{ font-size: 48px; margin-bottom: 16px; }}
.welcome-title {{
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 26px; font-weight: 800; color: {TE_BLACK};
    text-transform: uppercase; margin-bottom: 8px;
}}
.welcome-sub {{ font-size: 13px; color: #9A7A60; margin-bottom: 20px; line-height: 1.65; }}
.welcome-chip {{
    display: inline-block; background: #FFF0E6; border: 1px solid #F5C8A0;
    color: #B36030; font-size: 11px; font-weight: 600;
    border-radius: 20px; padding: 4px 12px; margin: 3px;
}}
.te-logo-mini {{
    display: inline-flex; align-items: center; gap: 8px;
    background: {TE_ORANGE}; border-radius: 8px; padding: 6px 14px; margin-bottom: 20px;
    font-family: 'Barlow Condensed', sans-serif; font-size: 15px;
    font-weight: 800; letter-spacing: 1px; color: white;
}}

/* ── BUTTONS ── */
.stDownloadButton > button {{
    background: linear-gradient(135deg, {TE_ORANGE}, {TE_DARK}) !important;
    color: white !important; border: none !important; border-radius: 8px !important;
    font-weight: 600 !important; padding: 8px 16px !important;
    box-shadow: 0 4px 14px rgba(232,101,10,0.3) !important;
    width: 100% !important; font-family: 'Barlow Condensed', sans-serif !important;
    font-size: 13px !important; letter-spacing: 0.5px !important; text-transform: uppercase !important;
}}
.stDownloadButton > button:hover {{
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(232,101,10,0.45) !important;
}}
.stButton > button {{
    background: {TE_NAVY} !important; color: white !important;
    border: none !important; border-radius: 8px !important;
    font-family: 'Barlow Condensed', sans-serif !important;
    font-size: 13px !important; font-weight: 700 !important;
    text-transform: uppercase !important; letter-spacing: 0.5px !important;
    padding: 9px 20px !important;
}}

/* ── EXPANDER ── */
[data-testid="stExpander"] {{
    background: {TE_WHITE} !important;
    border: 1px solid #EDE0D4 !important; border-radius: 12px !important;
}}
[data-testid="stExpander"] summary {{
    font-family: 'Barlow Condensed', sans-serif !important;
    font-weight: 700 !important; color: {TE_BLACK} !important;
    font-size: 14px !important; text-transform: uppercase !important;
    letter-spacing: 0.5px !important;
    background: #FFFAF6 !important; padding: 14px 20px !important;
}}

/* ── TABS ── */
[data-testid="stTabs"] [data-baseweb="tab"] {{
    font-family: 'Barlow Condensed', sans-serif !important;
    font-size: 14px !important; font-weight: 700 !important;
    letter-spacing: 1px !important; text-transform: uppercase !important;
}}
[data-testid="stTabs"] [aria-selected="true"] {{
    color: {TE_ORANGE} !important;
    border-bottom: 3px solid {TE_ORANGE} !important;
}}

/* ── DATAFRAME ── */
[data-testid="stDataFrame"] {{
    border: 1px solid #EDE0D4 !important;
    border-radius: 10px !important; overflow: hidden !important;
}}


/* ── MASQUER keyboard_double_arrow — remplacer par signe CSS ── */
[data-testid="collapsedControl"] {{
    display: flex !important;
    visibility: visible !important;
    opacity: 1 !important;
    background: {TE_ORANGE} !important;
    border-radius: 0 10px 10px 0 !important;
    width: 22px !important;
    min-height: 60px !important;
    align-items: center !important;
    justify-content: center !important;
    z-index: 9999 !important;
    cursor: pointer !important;
    overflow: hidden !important;
}}
[data-testid="collapsedControl"] span,
[data-testid="collapsedControl"] svg {{
    display: none !important;
    width: 0 !important;
    height: 0 !important;
    font-size: 0 !important;
    overflow: hidden !important;
}}
[data-testid="collapsedControl"]::before {{
    content: "❯" !important;
    color: white !important;
    font-size: 13px !important;
    font-weight: 900 !important;
    font-family: Arial, sans-serif !important;
    line-height: 1 !important;
    display: block !important;
}}

/* Browse files button couleur #2e1808 */
section[data-testid="stSidebar"] div[data-testid="stFileUploadDropzone"] button,
section[data-testid="stSidebar"] [data-testid="stFileUploader"] button {{
    background-color: #2e1808 !important;
    color: {TE_ORANGE} !important;
    border: 1px solid {TE_ORANGE} !important;
    border-radius: 6px !important;
    font-weight: 700 !important;
}}
section[data-testid="stSidebar"] div[data-testid="stFileUploadDropzone"] button:hover {{
    background-color: {TE_ORANGE} !important;
    color: white !important;
}}

/* Date input Period — texte #2e1808 */
section[data-testid="stSidebar"] [data-testid="stDateInput"] input,
section[data-testid="stSidebar"] input[type="text"] {{
    color: #2e1808 !important;
    background: white !important;
    font-weight: 700 !important;
    border: 1.5px solid {TE_ORANGE} !important;
    border-radius: 6px !important;
}}
section[data-testid="stSidebar"] [data-testid="stDateInput"] input::placeholder {{
    color: #2e1808 !important;
    opacity: 0.8 !important;
}}

::-webkit-scrollbar {{ width: 5px; height: 5px; }}
::-webkit-scrollbar-track {{ background: #F0EAE3; }}
::-webkit-scrollbar-thumb {{ background: {TE_ORANGE3}; border-radius: 3px; }}
::-webkit-scrollbar-thumb:hover {{ background: {TE_ORANGE}; }}
</style>
""", unsafe_allow_html=True)

# ── JS : remplacer "keyboard_double_arrow_left/right" par "‹" / "›" ──
st.markdown("""
<script>
function fixSidebarBtn() {
    const btn = document.querySelector('[data-testid="collapsedControl"]');
    if (!btn) return;
    const span = btn.querySelector('span');
    if (span) {
        const txt = span.textContent || '';
        if (txt.includes('keyboard')) {
            span.textContent = txt.includes('left') ? '‹' : '›';
            span.style.cssText = 'font-size:18px!important;color:white!important;font-weight:900!important;visibility:visible!important;display:block!important;width:auto!important;height:auto!important;font-family:Arial,sans-serif!important;line-height:1!important;';
        }
    }
}
// Lancer au chargement + observer les changements DOM
document.addEventListener('DOMContentLoaded', fixSidebarBtn);
setTimeout(fixSidebarBtn, 500);
setTimeout(fixSidebarBtn, 1500);
const obs = new MutationObserver(fixSidebarBtn);
obs.observe(document.body, {childList:true, subtree:true});
</script>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
#  LISTES DÉROULANTES — Editable Table
# ─────────────────────────────────────────────────────────────────────────────
SHIFTS = ["", "A (6-14h)", "B (14-22h)", "C (22-6h)"]

KEY_FAILURES = [
    "", "Laser monitoring", "Laser temp too high", "Dirt detected on the part",
    "part pos variation in front of the cam", "Hydra count issue", "Force monitoring",
    "Feeder issue", "Label print issue", "Camera loop variation", "Strip welding issue",
    "Reeling errors", "Dereeling errors", "air pressure error", "oil pressure error",
    "Raziol lubrication issue", "Bond welding detection issue", "camera strip driving",
    "Electrical power issue", "Bruderer system issue", "Extraction issue",
    "Cooling system issue", "Compressed air issue", "Camera setting issue",
    "Camera hardware issue", "Laser communication issue", "Laser water level",
    "Tooling detection", "Safety detection", "Blowing system error", "Hydraulic issue",
    "Equipement changeover", "Laser HW issue", "Laser internal cooling", "Machine guarding",
]


def load_data(f) -> pd.DataFrame:
    """Read CSV (auto-detect separator) or Excel. Clean column names."""
    name = f.name.lower()
    try:
        if name.endswith(".csv"):
            raw = f.read()
            sample = raw[:2048].decode("utf-8", errors="replace")
            sep = ";" if sample.count(";") >= sample.count(",") else ","
            df = pd.read_csv(io.BytesIO(raw), sep=sep, encoding="utf-8", on_bad_lines="skip")
        else:
            df = pd.read_excel(f)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"❌ File read error : `{e}`")
        return pd.DataFrame()


def check_missing(df: pd.DataFrame) -> list[str]:
    return [c for c in REQUIRED_COLS if c not in df.columns]


def fmt(val, decimals=2):
    """Formate un nombre pour affichage carte KPI."""
    if pd.isna(val): return "—"
    if val >= 1_000_000: return f"{val/1_000_000:.{decimals}f}M"
    if val >= 1000:      return f"{val/1000:.{decimals}f}k"
    return f"{val:,.{decimals}f}"


def sec_to_h(s):
    """Seconds → hours (4 decimal places)."""
    return round(float(s) / 3600.0, 4) if not pd.isna(s) else 0.0


def dl_png(fig, filename, label="⬇ Download PNG"):
    """PNG export button for a Plotly figure (requires kaleido)."""
    try:
        img = fig.to_image(format="png", width=1400, height=680, scale=2)
        st.download_button(label=label, data=img,
                           file_name=filename, mime="image/png",
                           use_container_width=True)
    except Exception:
        st.caption("_`pip install kaleido` pour activer l'export PNG_")


def export_excel(df: pd.DataFrame, kpi: dict) -> bytes:
    """
    Generate a multi-sheet Excel with KPIs + filtered rows.
    Uses openpyxl (included with pandas) — no need for xlsxwriter.
    """
    from openpyxl import Workbook
    from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side)
    from openpyxl.utils import get_column_letter

    buf = io.BytesIO()
    try:
        wb = Workbook()

        # ── Styles ──
        def hdr_style(ws, row, col, value, bg="E8650A", fg="FFFFFF"):
            cell = ws.cell(row=row, column=col, value=value)
            cell.fill      = PatternFill("solid", fgColor=bg)
            cell.font      = Font(bold=True, color=fg, name="Arial", size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin"))
            return cell

        def title_style(ws, row, col, value):
            cell = ws.cell(row=row, column=col, value=value)
            cell.font = Font(bold=True, color="1B2A4A", name="Arial", size=15)
            return cell

        def auto_width(ws):
            for col in ws.columns:
                max_w = max((len(str(c.value or "")) for c in col), default=10)
                ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_w + 4, 50)

        # ── Feuille 1 : KPI Summary ──
        ws1 = wb.active
        ws1.title = "KPI Summary"
        ws1.row_dimensions[1].height = 22
        title_style(ws1, 1, 1, "TE Connectivity — Stamping KPI Report")
        ws1.cell(2, 1, f"Generated on: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        hdr_style(ws1, 4, 1, "Indicateur")
        hdr_style(ws1, 4, 2, "Valeur")
        kpi_rows = [
            ("Availability globale (%)",    f"{kpi['dispo']:.2f}%"),
            ("Avg MTTR / stop (h)",       f"{kpi['mttr_mean_h']:.4f} h"),
            ("MTBF Moyen (h)",               f"{kpi['mtbf_mean_h']:.4f} h"),
            ("Total stops",        str(kpi['nb_arrets'])),
            ("Cumulative MTTR (h)",         f"{kpi['mttr_total_h']:.2f} h"),
            ("Cumulative MTBF (h)",         f"{kpi['mtbf_total_h']:.2f} h"),
            ("Total events analyzed", str(kpi['nb_rows'])),
        ]
        for i, (k, v) in enumerate(kpi_rows, start=5):
            ws1.cell(i, 1, k)
            ws1.cell(i, 2, v)
        auto_width(ws1)

        # ── Feuille 2 : Par Machine ──
        if "by_machine" in kpi and not kpi["by_machine"].empty:
            ws2 = wb.create_sheet("Par Machine")
            bm  = kpi["by_machine"].copy()
            bm.columns = ["Machine","Total MTTR (h)","Total MTBF (h)",
                          "Events","Availability (%)"]
            for ci, col_name in enumerate(bm.columns, start=1):
                hdr_style(ws2, 1, ci, col_name, bg="1B2A4A")
            for ri, row_vals in enumerate(bm.itertuples(index=False), start=2):
                for ci, v in enumerate(row_vals, start=1):
                    ws2.cell(ri, ci, round(v, 4) if isinstance(v, float) else v)
            auto_width(ws2)

        # ── Feuille 3 : Pareto ──
        if "pareto" in kpi and not kpi["pareto"].empty:
            ws3 = wb.create_sheet("Pareto Downtime")
            par = kpi["pareto"]
            for ci, col_name in enumerate(par.columns, start=1):
                hdr_style(ws3, 1, ci, col_name)
            for ri, row_vals in enumerate(par.itertuples(index=False), start=2):
                for ci, v in enumerate(row_vals, start=1):
                    ws3.cell(ri, ci, round(v, 4) if isinstance(v, float) else v)
            auto_width(ws3)

        # ── Feuille 4 : Données filtrées ──
        ws4      = wb.create_sheet("Filtered Data")
        exp_cols = [c for c in [COL_MACHINE, COL_DATE, COL_STATUS,
                                "mttr_h", "mtbf_h"] if c in df.columns]
        for ci, col_name in enumerate(exp_cols, start=1):
            hdr_style(ws4, 1, ci, col_name)
        for ri, row_vals in enumerate(df[exp_cols].itertuples(index=False), start=2):
            for ci, v in enumerate(row_vals, start=1):
                ws4.cell(ri, ci, str(v) if not isinstance(v, (int, float)) else v)
        auto_width(ws4)

        wb.save(buf)

    except Exception as e:
        st.error(f"Erreur export Excel : {e}")
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
#  PLOTLY BASE LAYOUT
# ─────────────────────────────────────────────────────────────────────────────
PL = dict(
    plot_bgcolor=TE_WHITE, paper_bgcolor=TE_WHITE,
    font=dict(family="Barlow, sans-serif", color="#4A3020", size=11),
    margin=dict(l=20, r=20, t=40, b=20),
    xaxis=dict(gridcolor="#F0E8E0", showgrid=True, zeroline=False,
               linecolor="#EDE0D4", tickfont=dict(size=10, color="#9A7A60")),
    yaxis=dict(gridcolor="#F0E8E0", showgrid=True, zeroline=False,
               linecolor="#EDE0D4", tickfont=dict(size=10, color="#9A7A60")),
    legend=dict(bgcolor=TE_WHITE, bordercolor="#EDE0D4", borderwidth=1,
                font=dict(size=11)),
    hoverlabel=dict(bgcolor=TE_BLACK, bordercolor=TE_BLACK,
                    font=dict(color="white", family="JetBrains Mono", size=11)),
)
PCONF = dict(displayModeBar=False, responsive=True)

def apply(fig, **kw):
    fig.update_layout(**{**PL, **kw})
    return fig

# ─────────────────────────────────────────────────────────────────────────────
#  SIDEBAR — toujours visible : Import + Filtres
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:

    # Logo
    st.markdown(f"""
    <div style="background:rgba(232,101,10,0.12);border:1px solid rgba(232,101,10,0.35);
                border-radius:10px;padding:12px 14px;margin-bottom:16px">
        <div style="font-family:'Barlow Condensed',sans-serif;font-size:20px;
                    font-weight:800;letter-spacing:1.5px;color:{TE_ORANGE}">
            ≡ TE CONNECTIVITY
        </div>
        <div style="font-family:'JetBrains Mono',monospace;font-size:7px;
                    letter-spacing:2px;color:rgba(255,255,255,0.18);margin-top:4px">
            TPM KPI DASHBOARD
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Import Data ──
    st.markdown('<p style="font-size:9px;font-weight:700;letter-spacing:3px;'
                f'text-transform:uppercase;color:{TE_ORANGE};margin-bottom:6px">'
                '📂 IMPORT DATA</p>', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "",
        type=["csv", "xlsx", "xls"],
        key="sidebar_uploader",
        label_visibility="collapsed"
    )

    st.markdown("---")

    # ── Filtres (seulement si fichier chargé) ──
    if uploaded is not None:
        st.markdown('<p style="font-size:9px;font-weight:700;letter-spacing:3px;'
                    f'text-transform:uppercase;color:{TE_ORANGE};margin-bottom:6px">'
                    '🔧 FILTERS</p>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
#  ÉCRAN D'ACCUEIL — si aucun fichier
# ─────────────────────────────────────────────────────────────────────────────
if uploaded is None:
    st.markdown(f"""
    <div style="display:flex;justify-content:center;margin-top:60px">
    <div style="background:white;border:2px dashed #F0C8A0;border-radius:18px;
                padding:48px 44px;text-align:center;max-width:500px;
                box-shadow:0 4px 24px rgba(0,0,0,0.07)">
        <div style="display:inline-flex;align-items:center;gap:8px;
                    background:{TE_ORANGE};border-radius:8px;padding:8px 18px;
                    font-family:'Barlow Condensed',sans-serif;font-size:16px;
                    font-weight:800;letter-spacing:2px;color:white;margin-bottom:20px">
            ≡ TE CONNECTIVITY
        </div>
        <div style="font-family:'Barlow Condensed',sans-serif;font-size:24px;
                    font-weight:800;color:#1C1C1C;text-transform:uppercase;
                    letter-spacing:1px;margin-bottom:12px">TPM KPI Dashboard</div>
        <div style="font-size:13px;color:#9A7A60;margin-bottom:20px;line-height:1.7">
            Import your Hydra MES file<br>
            to visualize the maintenance KPIs of Bruderer presses.
        </div>
        <div>
            <span style="background:#FFF0E6;border:1px solid #F5C8A0;color:#B36030;
                         font-size:11px;font-weight:600;border-radius:20px;
                         padding:4px 12px;margin:3px;display:inline-block">.csv comma</span>
            <span style="background:#FFF0E6;border:1px solid #F5C8A0;color:#B36030;
                         font-size:11px;font-weight:600;border-radius:20px;
                         padding:4px 12px;margin:3px;display:inline-block">.csv semicolon</span>
            <span style="background:#FFF0E6;border:1px solid #F5C8A0;color:#B36030;
                         font-size:11px;font-weight:600;border-radius:20px;
                         padding:4px 12px;margin:3px;display:inline-block">.xlsx</span>
        </div>
    </div></div>
    """, unsafe_allow_html=True)
    st.stop()


# Reset editable table if new file uploaded
if "last_file" not in st.session_state or st.session_state.last_file != uploaded.name:
    st.session_state.last_file  = uploaded.name
    st.session_state.edited_df  = None

df_raw = load_data(uploaded)
if df_raw.empty:
    st.stop()

# Supprimer les colonnes en double si le fichier source en contient
df_raw = df_raw.loc[:, ~df_raw.columns.duplicated()]

missing = check_missing(df_raw)
if missing:
    st.error(
        f"**Missing columns :** `{'`, `'.join(missing)}`\n\n"
        f"**Columns found in file :**\n```\n{chr(10).join(df_raw.columns.tolist())}\n```"
    )
    st.stop()

# ── Ajouter colonnes qualification si absentes ──
for col, default in [("Shift", ""), ("Key Failure", ""), ("User ID", "")]:
    if col not in df_raw.columns:
        df_raw[col] = default

# ── Initialiser session_state.edited_df ──
if st.session_state.edited_df is None:
    st.session_state.edited_df = df_raw.copy()

# ── Utiliser la version éditée (préserve les qualifications utilisateur) ──
df_raw = st.session_state.edited_df.copy()

df_raw[COL_MTTR] = pd.to_numeric(df_raw[COL_MTTR], errors="coerce").fillna(0.0)

# MTBF : optionnel — crée une colonne à 0 si absente du fichier
has_mtbf = COL_MTBF in df_raw.columns
if has_mtbf:
    df_raw[COL_MTBF] = pd.to_numeric(df_raw[COL_MTBF], errors="coerce").fillna(0.0)
    mtbf_all_zero = (df_raw[COL_MTBF] == 0).all()
    if mtbf_all_zero:
        has_mtbf = False   # présente mais vide → on traite comme absente
else:
    df_raw[COL_MTBF] = 0.0

# Colonnes en heures
df_raw["mttr_h"] = df_raw[COL_MTTR] / 3600.0
df_raw["mtbf_h"] = df_raw[COL_MTBF] / 3600.0 if has_mtbf else 0.0

# ── Override MTTR avec Manual Duration (min) si renseigné par l'utilisateur ──
if "Manual Duration (min)" in df_raw.columns:
    dur_mask = df_raw["Manual Duration (min)"].notna() & (df_raw["Manual Duration (min)"] > 0)
    df_raw.loc[dur_mask, "mttr_h"] = df_raw.loc[dur_mask, "Manual Duration (min)"] / 60.0

# Date — multi-format américain (MM/DD/YYYY, M/D/YYYY, MM-DD-YYYY, M-D-YYYY)
if COL_DATE in df_raw.columns:
    raw_dates = df_raw[COL_DATE].astype(str)
    parsed = pd.Series([pd.NaT] * len(df_raw), dtype="datetime64[ns]")

    formats_to_try = [
        "%m/%d/%Y %H:%M",   # 03/13/2025 14:30
        "%m/%d/%Y",          # 03/13/2025
        "%m-%d-%Y %H:%M",   # 03-13-2025 14:30
        "%m-%d-%Y",          # 03-13-2025
        "%-m/%-d/%Y",        # 3/13/2025  (Linux)
        "%#m/%#d/%Y",        # 3/13/2025  (Windows)
    ]

    for fmt in formats_to_try:
        mask = parsed.isna()
        if not mask.any():
            break
        try:
            parsed[mask] = pd.to_datetime(
                raw_dates[mask], format=fmt, errors="coerce", dayfirst=False)
        except Exception:
            pass

    # Fallback final : pandas auto avec dayfirst=False
    mask = parsed.isna()
    if mask.any():
        parsed[mask] = pd.to_datetime(
            raw_dates[mask], errors="coerce", dayfirst=False)

    df_raw[COL_DATE] = parsed
    df_raw["date_only"] = df_raw[COL_DATE].dt.normalize()
    n_ok = df_raw["date_only"].notna().sum()
    if n_ok == 0:
        st.warning(f"⚠ Column `{COL_DATE}`: no valid date parsed. "
                   f"Exemple de valeur brute : `{df_raw.iloc[0][COL_DATE] if len(df_raw) else 'N/A'}`")


# ─────────────────────────────────────────────────────────────────────────────
#  FILTRES SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
#  FILTRES SIDEBAR — injectés après chargement des données
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    all_machines = sorted(df_raw[COL_MACHINE].dropna().unique().tolist())
    sel_machines = st.multiselect(
        "Machines", options=all_machines, default=all_machines,
        placeholder="Select…"
    )

    if "date_only" in df_raw.columns:
        valid_d = df_raw["date_only"].dropna()
        if len(valid_d):
            from datetime import date as dt_date
            dmin = valid_d.min().date()
            dmax = valid_d.max().date()
            dmax_cal = max(dmax, dt_date.today())  # permettre jusqu'à aujourd'hui
            dr = st.date_input("Period", value=(dmin, dmax),
                               min_value=dmin, max_value=dmax_cal,
                               format="DD/MM/YYYY")
        else:
            dr = None
    else:
        dr = None

    st.markdown("---")
    st.markdown(f"""
    <div style="font-size:10px;color:rgba(255,255,255,0.3);
                font-family:'JetBrains Mono',monospace;letter-spacing:1px">
        📁 {uploaded.name}<br>
        📋 {len(df_raw):,} rows<br><br>
        TE CONNECTIVITY © {datetime.now().year}
    </div>
    """, unsafe_allow_html=True)

# Appliquer filtres
if not sel_machines:
    st.warning("⚠ Please select at least one machine.")
    st.stop()

df = df_raw[df_raw[COL_MACHINE].isin(sel_machines)].copy()
if "date_only" in df.columns and dr and isinstance(dr, (list, tuple)) and len(dr) == 2:
    df = df[(df["date_only"].dt.date >= dr[0]) & (df["date_only"].dt.date <= dr[1])]

if df.empty:
    st.warning("No data for this selection.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
#  CALCUL KPIs
# ─────────────────────────────────────────────────────────────────────────────
mt_total = df["mttr_h"].sum()
mb_total = df["mtbf_h"].sum()

# ── Availability ──
# Si MTBF disponible → formule standard MTBF/(MTBF+MTTR)
# Sinon → ratio rows PRODUCTION / total (basé out of machine_status_name)
if has_mtbf and mb_total > 0:
    total = mt_total + mb_total
    dispo = round(mb_total / total * 100, 2)
    dispo_mode = "MTBF"
else:
    # Fallback : comptage des rows par statut
    prod_mask = df[COL_STATUS].str.upper().str.contains("PRODUCTION", na=False)
    n_prod    = prod_mask.sum()
    n_total   = len(df)
    dispo     = round(n_prod / n_total * 100, 2) if n_total > 0 else 100.0
    dispo_mode = "STATUS"

stop_rows   = df[df["mttr_h"] > 0]
mttr_mean_h = round(stop_rows["mttr_h"].mean(), 4) if len(stop_rows) > 0 else 0.0
mtbf_mean_h = round(df["mtbf_h"].mean(), 4) if has_mtbf else 0.0
nb_arrets   = len(stop_rows)

# Par machine — "ne" est une méthode pandas, on utilise "nb_evt"
ma = df.groupby(COL_MACHINE, as_index=False).agg(
    tm=("mttr_h","sum"), tb=("mtbf_h","sum"), nb_evt=("mttr_h","count"))

if has_mtbf and mb_total > 0:
    # Formule MTBF standard
    ma["dispo"] = ma.apply(
        lambda r: round(r["tb"]/(r["tb"]+r["tm"])*100, 1)
        if (r["tb"]+r["tm"]) > 0 else 100.0, axis=1)
else:
    # Fallback : % rows PRODUCTION par machine
    prod_per_mac = (
        df[df[COL_STATUS].str.upper().str.contains("PRODUCTION", na=False)]
        .groupby(COL_MACHINE).size()
        .reset_index(name="n_prod")
    )
    total_per_mac = df.groupby(COL_MACHINE).size().reset_index(name="n_total")
    ratio = prod_per_mac.merge(total_per_mac, on=COL_MACHINE, how="right").fillna(0)
    ratio["dispo"] = (ratio["n_prod"] / ratio["n_total"] * 100).round(1)
    ma = ma.merge(ratio[[COL_MACHINE, "dispo"]], on=COL_MACHINE, how="left")
    ma["dispo"] = ma["dispo"].fillna(0.0)

# Pareto
pareto = df[df["mttr_h"]>0].groupby(COL_MACHINE)["mttr_h"].sum().reset_index()
pareto.columns = ["Machine","MTTR_total_h"]
pareto = pareto.sort_values("MTTR_total_h", ascending=False).reset_index(drop=True)
pareto["Pct"]   = (pareto["MTTR_total_h"]/pareto["MTTR_total_h"].sum()*100).round(1)
pareto["Cumul"] = pareto["Pct"].cumsum().round(1)

kpi = dict(
    dispo=dispo, mttr_mean_h=mttr_mean_h, mtbf_mean_h=mtbf_mean_h,
    nb_arrets=nb_arrets, mttr_total_h=round(mt_total,2),
    mtbf_total_h=round(mb_total,2), nb_rows=len(df),
    by_machine=ma, pareto=pareto,
)

# Production status (si colonne présente)
has_prod = COL_PROD in df.columns
if has_prod:
    prod_ct    = df[df[COL_PROD].str.lower().str.contains("prod", na=False)].shape[0]
    nonprod_ct = len(df) - prod_ct
else:
    prod_ct = nonprod_ct = 0

# ─────────────────────────────────────────────────────────────────────────────
#  HEADER — UNIQUE
# ─────────────────────────────────────────────────────────────────────────────
dispo_col = TE_GREEN if dispo >= 95 else TE_AMBER if dispo >= 90 else TE_RED
dispo_lbl = "On Target ✅" if dispo >= 95 else "Watch ⚠" if dispo >= 90 else "Critical 🔴"

st.markdown(f"""
<div class="te-header">
  <div style="display:flex;justify-content:space-between;align-items:flex-start">
    <div>
      <div class="te-header-tag">Stamping Department · Bruderer Presses</div>
      <div class="te-header-title">Dashboard <span>KPI</span> Maintenance</div>
      <div class="te-header-sub">MTTR · MTBF · Availability · Criticality · Pareto Analysis</div>
      <div class="te-header-badge">
        ⚙ {df[COL_MACHINE].nunique()} machine{"s" if df[COL_MACHINE].nunique()>1 else ""}
        &nbsp;·&nbsp; {len(df):,} events
        &nbsp;·&nbsp; <span style="color:{dispo_col};font-weight:700">{dispo}% — {dispo_lbl}</span>
      </div>
    </div>
    <div class="te-header-right">
      <div class="te-live"><div class="te-live-dot"></div>SYSTEM ACTIVE</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:10px;
                  color:rgba(255,255,255,0.35);text-align:right;margin-top:4px">
        TANGIER · PLANT 1310
      </div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# MTBF info banner
if not has_mtbf:
    st.markdown(f"""
    <div style="background:#fff8e1;border:1px solid #ffe082;border-left:4px solid {TE_AMBER};
                border-radius:8px;padding:10px 18px;font-size:12px;color:#5d4037;
                margin-bottom:12px;display:flex;align-items:center;gap:10px">
      ℹ️ <span><strong>MTBF column absent or empty.</strong>
      Availability computed from machine status (PRODUCTION rows / total).
      Export <code>Sum of mtbf_numerator_seconds_quantity</code> from Hydra for exact formula.</span>
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
#  NAVIGATION — TABS (exclusif)
# ─────────────────────────────────────────────────────────────────────────────
tab_kpi, tab_qual = st.tabs([
    "📊  Analyse des Performances (KPIs)",
    "📝  Qualification des Arrêts",
])

# ═══════════════════════════════════════════════════════════════════════════════
#  TAB 1 — ANALYSE DES PERFORMANCES
# ═══════════════════════════════════════════════════════════════════════════════
with tab_kpi:

    # Status bar
    st.markdown(f"""
    <div class="te-statusbar">
      <div class="te-dot-green"></div>
      <div class="te-statusbar-item"><strong>{len(df):,}</strong>&nbsp;events</div>
      <div class="te-sep"></div>
      <div class="te-statusbar-item">Stops: <strong>{nb_arrets}</strong></div>
      <div class="te-sep"></div>
      <div class="te-statusbar-item">Machines: <strong>{df[COL_MACHINE].nunique()}</strong></div>
      {"<div class='te-sep'></div><div class='te-statusbar-item'>Production: <strong>"+str(prod_ct)+"</strong></div>" if has_prod else ""}
      <div class="te-sep"></div>
      <div class="te-statusbar-item">
        Availability: <strong style="color:{dispo_col}">{dispo}% — {dispo_lbl}</strong>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── KPI Cards ──────────────────────────────────────────────────────────────
    st.markdown('<div class="te-section">Main KPIs</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    for col, icon, label, value, unit in [
        (c1, "⚡", "AVAILABILITY",
         f"{dispo}%", f"{'MTBF-based' if has_mtbf else 'Status-based'}"),
        (c2, "🔧", "AVG MTTR / STOP",
         f"{mttr_mean_h:.3f} h", f"{round(mttr_mean_h*60,1)} min per stop"),
        (c3, "✅", "AVG MTBF",
         f"{mtbf_mean_h:.2f} h" if has_mtbf else "N/A",
         "Time between failures" if has_mtbf else "Column absent"),
        (c4, "⚠",  "TOTAL STOPS",
         f"{nb_arrets}", f"Out of {len(df):,} events"),
    ]:
        with col:
            st.markdown(f"""
            <div class="kpi-card">
              <div class="kpi-icon">{icon}</div>
              <div class="kpi-label">{label}</div>
              <div class="kpi-value">{value}</div>
              <div class="kpi-divider"></div>
              <div class="kpi-unit">{unit}</div>
            </div>""", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════════
    #  SECTION — ÉVOLUTION DES PERFORMANCES (hebdo + mensuel)
    # ═══════════════════════════════════════════════════════════════════════════
    st.markdown('<div class="te-section">📈 Évolution des Performances</div>',
                unsafe_allow_html=True)

    # ── Extraction des time features (robuste, sans conflit colonne Hydra) ────
    if "date_only" in df.columns:
        _MONTH_FR = {1:"Jan", 2:"Fév", 3:"Mar", 4:"Avr", 5:"Mai", 6:"Juin",
                     7:"Juil", 8:"Aoû", 9:"Sep", 10:"Oct", 11:"Nov", 12:"Déc"}
        _df_t = df.copy()
        _dt   = pd.to_datetime(_df_t["date_only"], errors="coerce")
        _df_t["_month_num"]  = _dt.dt.month
        _df_t["_month_year"] = _dt.dt.to_period("M").astype(str)   # "2025-03"
        _df_t["_month_lbl"]  = _dt.dt.month.map(_MONTH_FR).fillna("—") + \
                                " " + _dt.dt.year.astype(str).str[-2:]  # "Mar 25"
        _df_t["_week_num"]   = _dt.dt.isocalendar().week.astype("Int64")
        _df_t["_week_year"]  = _dt.dt.isocalendar().year.astype("Int64")
        _df_t["_week_lbl"]   = "S" + _df_t["_week_num"].astype(str)  # "S12"

        # ── Fonction d'agrégat générique ──────────────────────────────────────
        def _te_agg(lbl_col, sort_col):
            """Agrège MTTR/MTBF/dispo par période. Retourne df trié."""
            agg = _df_t.groupby(lbl_col, as_index=False).agg(
                _sort    =(sort_col,  "first"),
                mttr_sum =("mttr_h",  "sum"),
                mtbf_sum =("mtbf_h",  "sum"),
                nb_stops =("mttr_h",  lambda x: (x > 0).sum()),
                nb_events=("mttr_h",  "count"),
            ).sort_values("_sort")

            # Disponibilité = MTBF / (MTBF + MTTR) ou fallback statut
            if has_mtbf:
                agg["dispo"] = agg.apply(
                    lambda r: round(r.mtbf_sum / (r.mtbf_sum + r.mttr_sum) * 100, 2)
                    if (r.mtbf_sum + r.mttr_sum) > 0 else 100.0, axis=1)
            else:
                prod_grp = (
                    _df_t[_df_t[COL_STATUS].str.upper().str.contains("PRODUCTION", na=False)]
                    .groupby(lbl_col).size().reset_index(name="n_prod"))
                tot_grp = _df_t.groupby(lbl_col).size().reset_index(name="n_tot")
                ratio   = prod_grp.merge(tot_grp, on=lbl_col, how="right").fillna(0)
                ratio["dispo"] = (ratio["n_prod"] / ratio["n_tot"] * 100).round(2)
                agg = agg.merge(ratio[[lbl_col, "dispo"]], on=lbl_col, how="left")
                agg["dispo"] = agg["dispo"].fillna(0.0)

            agg = agg.rename(columns={lbl_col: "label",
                                       "mttr_sum": "mttr_h",
                                       "mtbf_sum": "mtbf_h"})
            agg["mttr_h"] = agg["mttr_h"].round(4)
            agg["mtbf_h"] = agg["mtbf_h"].round(3)
            return agg[["label","mttr_h","mtbf_h","dispo","nb_stops","nb_events"]].reset_index(drop=True)

        _df_week  = _te_agg("_week_lbl",  "_week_year")
        _df_month = _te_agg("_month_lbl", "_month_year")

        # ── Line chart helper ─────────────────────────────────────────────────
        def _te_line(x_vals, y_vals, title, y_title, color,
                     target=None, y_fmt=None, height=240):
            """Retourne une figure Plotly line+markers avec fill subtil."""
            fig = go.Figure()
            # Ligne target
            if target is not None:
                fig.add_trace(go.Scatter(
                    x=x_vals, y=[target]*len(x_vals),
                    mode="lines", name=f"Cible {target}%",
                    line=dict(color=TE_RED, dash="dot", width=1.5),
                    hoverinfo="skip"))
            # Fill zone sous la courbe
            fill_color = (f"rgba(232,101,10,0.07)"  if color == TE_ORANGE else
                          f"rgba(27,42,74,0.07)"     if color == TE_NAVY   else
                          f"rgba(39,174,96,0.07)")
            fig.add_trace(go.Scatter(
                x=x_vals, y=y_vals,
                mode="lines+markers",
                name=y_title,
                line=dict(color=color, width=2.5),
                marker=dict(size=8, color=color,
                            line=dict(color="white", width=2),
                            symbol="circle"),
                fill="tozeroy", fillcolor=fill_color,
                hovertemplate=f"<b>%{{x}}</b><br>{y_title}: <b>%{{y}}</b><extra></extra>",
            ))
            y_axis = dict(gridcolor="#F0E8E0", zeroline=False,
                          tickfont=dict(size=9, color="#9A7A60"))
            if y_fmt:
                y_axis["tickformat"] = y_fmt
            if target is not None:
                _safe_min = min((v for v in y_vals if pd.notna(v)), default=0)
                y_axis["range"] = [max(0, _safe_min - 5), 105]
            apply(fig, height=height, showlegend=False,
                  title=dict(text=title,
                             font=dict(size=11, color=TE_BLACK,
                                       family="Barlow Condensed"),
                             x=0.01, y=0.97),
                  xaxis=dict(tickfont=dict(size=9, color="#9A7A60"),
                             gridcolor="#F0E8E0", zeroline=False,
                             tickangle=-35 if len(x_vals) > 8 else 0),
                  yaxis=y_axis,
                  margin=dict(l=10, r=10, t=34, b=30))
            return fig

        # ── Label mini-section ─────────────────────────────────────────────────
        def _te_mini_label(txt):
            st.markdown(
                f'<div style="font-family:\'JetBrains Mono\',monospace;font-size:8px;'
                f'font-weight:700;letter-spacing:2px;text-transform:uppercase;'
                f'color:{TE_ORANGE};margin:12px 0 6px 0;display:flex;align-items:center;gap:8px">'
                f'<span style="width:12px;height:2px;background:{TE_ORANGE};display:inline-block"></span>'
                f'{txt}'
                f'<span style="flex:1;height:1px;background:linear-gradient(90deg,#F0D0B0,transparent);'
                f'display:inline-block"></span></div>',
                unsafe_allow_html=True)

        # ── Tableau récap helper ───────────────────────────────────────────────
        def _te_recap_table(df_agg, periode_col):
            _tbl = df_agg.rename(columns={
                "label":     periode_col,
                "mttr_h":    "MTTR (h)",
                "mtbf_h":    "MTBF (h)",
                "dispo":     "Disponibilité (%)",
                "nb_stops":  "Arrêts",
                "nb_events": "Events",
            })
            def _sd(val):
                try:
                    v = float(val)
                    if v >= 95: return "background:#d5f5e3;color:#1e8449;font-weight:700"
                    if v >= 90: return "background:#fef9e7;color:#d68910;font-weight:700"
                    return              "background:#fdf2f2;color:#c0392b;font-weight:700"
                except: return ""
            st.dataframe(
                _tbl.style
                    .applymap(_sd, subset=["Disponibilité (%)"])
                    .format({"Disponibilité (%)": "{:.2f}%",
                             "MTTR (h)": "{:.4f}",
                             "MTBF (h)": "{:.3f}"}),
                use_container_width=True, hide_index=True,
                height=min(380, len(_tbl) * 36 + 42))

        # ── Sub-tabs ───────────────────────────────────────────────────────────
        _stab_w, _stab_m = st.tabs(["📆  Vue Hebdomadaire", "📅  Vue Mensuelle"])

        # ─── VUE HEBDOMADAIRE ─────────────────────────────────────────────────
        with _stab_w:
            if len(_df_week) < 2:
                st.info("Pas assez de données hebdomadaires (minimum 2 semaines requises).")
            else:
                _wc1, _wc2, _wc3 = st.columns(3, gap="medium")
                with _wc1:
                    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
                    st.plotly_chart(
                        _te_line(_df_week["label"], _df_week["dispo"],
                                 "Disponibilité (%)", "Dispo (%)", TE_GREEN,
                                 target=95, y_fmt=".1f"),
                        use_container_width=True, config=PCONF)
                    st.markdown("</div>", unsafe_allow_html=True)
                with _wc2:
                    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
                    st.plotly_chart(
                        _te_line(_df_week["label"],
                                 _df_week["mtbf_h"] if has_mtbf else [0]*len(_df_week),
                                 "MTBF (h) — Fiabilité", "MTBF (h)", TE_NAVY),
                        use_container_width=True, config=PCONF)
                    if not has_mtbf:
                        st.caption("⚠ Colonne MTBF absente du fichier Hydra.")
                    st.markdown("</div>", unsafe_allow_html=True)
                with _wc3:
                    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
                    st.plotly_chart(
                        _te_line(_df_week["label"], _df_week["mttr_h"],
                                 "MTTR (h) — Réparabilité", "MTTR (h)", TE_ORANGE),
                        use_container_width=True, config=PCONF)
                    st.markdown("</div>", unsafe_allow_html=True)

                _te_mini_label("Récapitulatif hebdomadaire")
                _te_recap_table(_df_week, "Semaine")

        # ─── VUE MENSUELLE ────────────────────────────────────────────────────
        with _stab_m:
            if len(_df_month) < 2:
                st.info("Pas assez de données mensuelles (minimum 2 mois requis).")
            else:
                _mc1, _mc2, _mc3 = st.columns(3, gap="medium")
                with _mc1:
                    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
                    st.plotly_chart(
                        _te_line(_df_month["label"], _df_month["dispo"],
                                 "Disponibilité (%)", "Dispo (%)", TE_GREEN,
                                 target=95, y_fmt=".1f"),
                        use_container_width=True, config=PCONF)
                    st.markdown("</div>", unsafe_allow_html=True)
                with _mc2:
                    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
                    st.plotly_chart(
                        _te_line(_df_month["label"],
                                 _df_month["mtbf_h"] if has_mtbf else [0]*len(_df_month),
                                 "MTBF (h) — Fiabilité", "MTBF (h)", TE_NAVY),
                        use_container_width=True, config=PCONF)
                    if not has_mtbf:
                        st.caption("⚠ Colonne MTBF absente du fichier Hydra.")
                    st.markdown("</div>", unsafe_allow_html=True)
                with _mc3:
                    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
                    st.plotly_chart(
                        _te_line(_df_month["label"], _df_month["mttr_h"],
                                 "MTTR (h) — Réparabilité", "MTTR (h)", TE_ORANGE),
                        use_container_width=True, config=PCONF)
                    st.markdown("</div>", unsafe_allow_html=True)

                _te_mini_label("Récapitulatif mensuel")
                _te_recap_table(_df_month, "Mois")

    else:
        st.info("⚠ Colonne `plant_shift_date` absente — évolution temporelle indisponible.")

    # ── Pareto + Pie ───────────────────────────────────────────────────────────
    st.markdown('<div class="te-section">Pareto & Cause Analysis</div>', unsafe_allow_html=True)
    col_l, col_r = st.columns(2, gap="medium")

    with col_l:
        st.markdown("""<div class="chart-card">
          <div class="chart-header"><div class="chart-dot"></div>
          <div class="chart-title">📊 Downtime Pareto</div></div>""",
          unsafe_allow_html=True)
        if not pareto.empty:
            bc = [TE_ORANGE if i<2 else TE_NAVY if i<4 else "#A8A8A8"
                  for i in range(len(pareto))]
            fig_par = make_subplots(specs=[[{"secondary_y": True}]])
            fig_par.add_trace(go.Bar(
                x=pareto["Machine"], y=pareto["MTTR_total_h"], name="Downtime (h)",
                marker=dict(color=bc, line=dict(width=0)),
                text=[f"{v:.2f}h" for v in pareto["MTTR_total_h"]],
                textposition="outside", textfont=dict(size=10, color="#4A3020"),
                hovertemplate="<b>%{x}</b><br>Downtime: %{y:.3f} h<extra></extra>"
            ), secondary_y=False)
            fig_par.add_trace(go.Scatter(
                x=pareto["Machine"], y=pareto["Cumul"], name="Cumul (%)",
                mode="lines+markers",
                line=dict(color=TE_RED, width=2.5),
                marker=dict(size=8, color=TE_RED, line=dict(color="white", width=2)),
                hovertemplate="<b>%{x}</b><br>Cumul: <b>%{y:.1f}%</b><extra></extra>"
            ), secondary_y=True)
            fig_par.add_hline(y=80, line_dash="dot", line_color=TE_RED, line_width=1.5,
                              secondary_y=True, annotation_text="80%",
                              annotation_font=dict(color=TE_RED, size=10))
            fig_par.update_layout(**{**PL, "height": 320, "bargap": 0.3,
                "yaxis":  dict(title="Downtime (h)", gridcolor="#F0E8E0",
                               tickfont=dict(size=9, color="#9A7A60"), zeroline=False),
                "yaxis2": dict(title="Cumul (%)", range=[0, 115], ticksuffix="%",
                               gridcolor="#F0E8E0", tickfont=dict(size=9, color="#9A7A60"),
                               zeroline=False),
                "xaxis":  dict(tickfont=dict(size=11, color="#4A3020"), zeroline=False),
            })
            st.plotly_chart(fig_par, use_container_width=True, config=PCONF)
            top1 = pareto.iloc[0]
            top2_pct = pareto.head(2)["Pct"].sum()
            st.markdown(
                f'<div class="te-insight-crit">🔴 <strong>{top1["Machine"]}</strong> '
                f'= <strong>{top1["Pct"]}%</strong> of downtime. '
                f'Top 2 = <strong>{top2_pct:.1f}%</strong>.</div>',
                unsafe_allow_html=True)
        else:
            st.info("No stops detected.")
        st.markdown("</div>", unsafe_allow_html=True)

    with col_r:
        st.markdown("""<div class="chart-card">
          <div class="chart-header"><div class="chart-dot"></div>
          <div class="chart-title">🍩 Cause Analysis</div></div>""",
          unsafe_allow_html=True)
        sc = df[COL_STATUS].value_counts().reset_index()
        sc.columns = ["Statut", "Nombre"]
        status_colors = {"PRODUCTION": TE_GREEN, "micro stops": TE_ORANGE,
                         "Réparation Peripheri": TE_RED, "Réparation": TE_RED}
        pie_colors = [status_colors.get(s, "#A8A8A8") for s in sc["Statut"]]
        fig_pie = go.Figure(go.Pie(
            labels=sc["Statut"], values=sc["Nombre"], hole=0.58,
            marker=dict(colors=pie_colors, line=dict(color="white", width=3)),
            textinfo="percent", textfont=dict(size=11, family="Barlow Condensed"),
            hovertemplate="<b>%{label}</b><br>%{value} events<br>%{percent}<extra></extra>"
        ))
        apply(fig_pie, height=320,
            annotations=[dict(text=f"<b>{len(df):,}</b><br>events", x=0.5, y=0.5,
                showarrow=False, font=dict(size=14, color=TE_BLACK, family="Barlow Condensed"))],
            legend=dict(orientation="v", x=1, y=0.5, font=dict(size=10)))
        st.plotly_chart(fig_pie, use_container_width=True, config=PCONF)
        micro_row = sc[sc["Statut"].str.lower().str.contains("micro", na=False)]
        rep_row   = sc[sc["Statut"].str.lower().str.contains("réparat|reparat", na=False)]
        if not micro_row.empty and not rep_row.empty:
            mc = int(micro_row["Nombre"].sum())
            rc = int(rep_row["Nombre"].sum())
            dominant = "Micro-Stops" if mc > rc else "Repairs"
            st.markdown(
                f'<div class="te-insight">ℹ️ <strong>{dominant}</strong> dominate '
                f'({max(mc,rc)} occurrences). '
                f'{"Focus on 5S and standardization." if mc>rc else "Strengthen preventive maintenance."}'
                f'</div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # ── Criticality Matrix + Daily Trend ──────────────────────────────────────
    st.markdown('<div class="te-section">Criticality & Time Trend</div>', unsafe_allow_html=True)
    col_a, col_b = st.columns([2, 3], gap="medium")

    with col_a:
        st.markdown("""<div class="chart-card">
          <div class="chart-header"><div class="chart-dot"></div>
          <div class="chart-title">🎯 Criticality Matrix</div></div>""",
          unsafe_allow_html=True)
        mx_v = ma["tm"].max() * 1.55 or 10
        my_v = ma["tb"].max() * 1.55 or 100
        midx, midy = mx_v / 2, my_v / 2
        fig_mat = go.Figure()
        for x0, y0, x1, y1, c in [
            (0, midy, midx, my_v, "rgba(39,174,96,0.07)"),
            (midx, midy, mx_v, my_v, "rgba(27,42,74,0.07)"),
            (0, 0, midx, midy, "rgba(230,126,34,0.07)"),
            (midx, 0, mx_v, midy, "rgba(192,57,43,0.10)"),
        ]:
            fig_mat.add_shape(type="rect", x0=x0, y0=y0, x1=x1, y1=y1,
                              fillcolor=c, line_width=0, layer="below")
        for txt, x, y, tc in [
            ("✅ RELIABLE",  midx*0.04, my_v*0.97, TE_GREEN),
            ("👁 WATCH",    mx_v*0.54, my_v*0.97, TE_NAVY),
            ("⚠ IMPROVE",  midx*0.04, my_v*0.47, TE_AMBER),
            ("🔴 CRITICAL", mx_v*0.54, my_v*0.47, TE_RED),
        ]:
            fig_mat.add_annotation(x=x, y=y, text=txt, showarrow=False,
                font=dict(size=9, color=tc, family="Barlow Condensed"),
                xanchor="left", yanchor="top")
        fig_mat.add_hline(y=midy, line_dash="dot", line_color="#D4CFC9", line_width=1.5)
        fig_mat.add_vline(x=midx, line_dash="dot", line_color="#D4CFC9", line_width=1.5)
        for i, row in ma.iterrows():
            c_dot = PALETTE[i % len(PALETTE)]
            fig_mat.add_trace(go.Scatter(
                x=[row["tm"]], y=[row["tb"]],
                mode="markers+text", name=row[COL_MACHINE],
                text=[row[COL_MACHINE]], textposition="top center",
                textfont=dict(size=11, color=c_dot, family="Barlow Condensed"),
                marker=dict(size=min(60, max(22, int(row["nb_evt"]) * 3)),
                            color=c_dot, opacity=0.88,
                            line=dict(color="white", width=3)),
                hovertemplate=(f"<b>{row[COL_MACHINE]}</b><br>"
                               "MTTR: %{x:.2f} h<br>MTBF: %{y:.2f} h<br>"
                               f"Avail.: {row['dispo']}%<extra></extra>")
            ))
        apply(fig_mat, height=360, showlegend=False,
            xaxis=dict(title="Total MTTR (h)", range=[0, mx_v], gridcolor="#F0E8E0",
                       tickfont=dict(size=9, color="#9A7A60"), zeroline=False),
            yaxis=dict(title="Total MTBF (h)", range=[0, my_v], gridcolor="#F0E8E0",
                       tickfont=dict(size=9, color="#9A7A60"), zeroline=False))
        st.plotly_chart(fig_mat, use_container_width=True, config=PCONF)
        st.markdown("""
        <div class="quad-grid">
          <div class="quad q-good"><h5>↖ Reliable + Fast</h5><p>Maintain standard PM</p></div>
          <div class="quad q-watch"><h5>↗ Monitor</h5><p>Improve repair procedure</p></div>
          <div class="quad q-warn"><h5>↙ Improve</h5><p>Reinforced preventive PM</p></div>
          <div class="quad q-crit"><h5>↘ Bad Actor 🔴</h5><p>Absolute TPM priority</p></div>
        </div>""", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with col_b:
        st.markdown("""<div class="chart-card">
          <div class="chart-header"><div class="chart-dot"></div>
          <div class="chart-title">📈 Daily Availability Trend</div></div>""",
          unsafe_allow_html=True)
        if "date_only" in df.columns:
            dm = df.groupby(["date_only", COL_MACHINE]).agg(
                mt=("mttr_h","sum"), mb=("mtbf_h","sum")).reset_index()
            dm["dp"] = dm.apply(
                lambda r: round(r.mb/(r.mb+r.mt)*100,1) if (r.mb+r.mt)>0 else 100.0, axis=1)
            da = df.groupby("date_only").agg(
                mt=("mttr_h","sum"), mb=("mtbf_h","sum")).reset_index()
            da["dp"] = da.apply(
                lambda r: round(r.mb/(r.mb+r.mt)*100,1) if (r.mb+r.mt)>0 else 100.0, axis=1)
            fig_evo = go.Figure()
            fig_evo.add_trace(go.Scatter(
                x=da["date_only"], y=[95]*len(da), mode="lines", name="Target 95%",
                line=dict(color=TE_RED, dash="dot", width=1.5)))
            for i, mac in enumerate(sorted(dm[COL_MACHINE].unique())):
                d2 = dm[dm[COL_MACHINE]==mac].sort_values("date_only")
                c2 = PALETTE[i % len(PALETTE)]
                fig_evo.add_trace(go.Scatter(
                    x=d2["date_only"], y=d2["dp"], mode="lines+markers", name=mac,
                    line=dict(color=c2, width=2),
                    marker=dict(size=6, color=c2, line=dict(color="white", width=2)),
                    hovertemplate=f"<b>{mac}</b><br>%{{x|%m/%d/%Y}}<br>Avail.: <b>%{{y}}%</b><extra></extra>"
                ))
            fig_evo.add_trace(go.Scatter(
                x=da["date_only"], y=da["dp"], mode="lines", name="▶ Global",
                line=dict(color=TE_NAVY, width=3, dash="dot"),
                hovertemplate="Global<br>%{x|%m/%d/%Y}<br>Avail.: <b>%{y}%</b><extra></extra>"))
            apply(fig_evo, height=360,
                yaxis=dict(ticksuffix="%", range=[60,105], gridcolor="#F0E8E0",
                           tickfont=dict(size=9, color="#9A7A60"), zeroline=False),
                xaxis=dict(tickformat="%d/%m", gridcolor="#F0E8E0",
                           tickfont=dict(size=9, color="#9A7A60"), zeroline=False))
            st.plotly_chart(fig_evo, use_container_width=True, config=PCONF)
        else:
            st.info("Column `plant_shift_date` absent — trend unavailable.")
        st.markdown("</div>", unsafe_allow_html=True)

    # ── MTBF + MTTR by Machine ─────────────────────────────────────────────────
    st.markdown('<div class="te-section">MTTR & MTBF Detail by Machine</div>', unsafe_allow_html=True)
    col_c, col_d = st.columns(2, gap="medium")

    with col_c:
        st.markdown("""<div class="chart-card">
          <div class="chart-header"><div class="chart-dot"></div>
          <div class="chart-title">Avg MTBF per Machine (h)</div></div>""",
          unsafe_allow_html=True)
        mb_m = df.groupby(COL_MACHINE)["mtbf_h"].mean().reset_index()
        mb_m.columns = ["Machine","MTBF"]
        mb_m = mb_m.sort_values("MTBF", ascending=True)
        fig_mb = go.Figure(go.Bar(
            x=mb_m["MTBF"], y=mb_m["Machine"], orientation="h",
            marker=dict(color=mb_m["MTBF"],
                        colorscale=[[0,"#FAD9B5"],[0.5,TE_ORANGE2],[1,TE_DARK]],
                        showscale=False, line=dict(width=0)),
            text=mb_m["MTBF"].apply(lambda v: f"{v:.3f}h"),
            textposition="outside", textfont=dict(size=10, color="#6A4030"),
            hovertemplate="<b>%{y}</b><br>MTBF: %{x:.4f} h<extra></extra>"
        ))
        apply(fig_mb, height=max(240, len(mb_m)*55), bargap=0.35, showlegend=False,
            xaxis=dict(gridcolor="#F0E8E0", tickfont=dict(size=9,color="#9A7A60"), zeroline=False),
            yaxis=dict(showgrid=False, tickfont=dict(size=11,color="#4A3020")))
        st.plotly_chart(fig_mb, use_container_width=True, config=PCONF)
        st.markdown("</div>", unsafe_allow_html=True)

    with col_d:
        st.markdown("""<div class="chart-card">
          <div class="chart-header"><div class="chart-dot"></div>
          <div class="chart-title">Avg MTTR per Machine (h)</div></div>""",
          unsafe_allow_html=True)
        mt_m = df[df["mttr_h"]>0].groupby(COL_MACHINE)["mttr_h"].mean().reset_index()
        mt_m.columns = ["Machine","MTTR"]
        mt_m = mt_m.sort_values("MTTR", ascending=False)
        fig_mt = go.Figure(go.Bar(
            x=mt_m["Machine"], y=mt_m["MTTR"],
            marker=dict(color=mt_m["MTTR"],
                        colorscale=[[0,"#FAD9B5"],[0.5,TE_ORANGE],[1,TE_RED]],
                        showscale=False, line=dict(width=0)),
            hovertemplate="<b>%{x}</b><br>MTTR: %{y:.4f} h<extra></extra>"
        ))
        apply(fig_mt, height=max(240, len(mt_m)*55), bargap=0.35, showlegend=False,
            xaxis=dict(showgrid=False, tickfont=dict(size=11,color="#4A3020")),
            yaxis=dict(gridcolor="#F0E8E0", tickfont=dict(size=9,color="#9A7A60"), zeroline=False))
        st.plotly_chart(fig_mt, use_container_width=True, config=PCONF)
        st.markdown("</div>", unsafe_allow_html=True)

    # ── Summary Table ──────────────────────────────────────────────────────────
    st.markdown('<div class="te-section">Summary Table by Machine</div>', unsafe_allow_html=True)

    ma_disp = ma.copy()
    ma_disp.columns = ["Machine","Total MTTR (h)","Total MTBF (h)","Events","Availability (%)"]
    ma_disp["Total MTTR (h)"]   = ma_disp["Total MTTR (h)"].round(4)
    ma_disp["Total MTBF (h)"]   = ma_disp["Total MTBF (h)"].round(2)
    ma_disp["Availability (%)"] = ma_disp["Availability (%)"].astype(float)

    worst = ma_disp.loc[ma_disp["Availability (%)"].idxmin()]
    best  = ma_disp.loc[ma_disp["Availability (%)"].idxmax()]
    ci1, ci2 = st.columns(2)
    with ci1:
        if worst["Availability (%)"] < 90:
            st.markdown(
                f'<div class="te-insight-crit">🔴 <strong>Bad Actor: {worst["Machine"]}</strong>'
                f' — Avail. {worst["Availability (%)"]:.1f}% '
                f'(MTTR = {worst["Total MTTR (h)"]:.3f} h). Priority TPM action.</div>',
                unsafe_allow_html=True)
        else:
            st.markdown(
                f'<div class="te-insight-ok">✅ All machines ≥ 90%. '
                f'Best: <strong>{best["Machine"]}</strong> ({best["Availability (%)"]:.1f}%).</div>',
                unsafe_allow_html=True)
    with ci2:
        st.markdown(
            f'<div class="te-insight">📊 <strong>Global avg MTTR:</strong> '
            f'{mttr_mean_h:.3f} h ({round(mttr_mean_h*60,1)} min) · '
            f'<strong>Total Stops:</strong> {nb_arrets} / {len(df):,} events.</div>',
            unsafe_allow_html=True)

    def style_dispo(val):
        try:
            v = float(str(val).replace("%",""))
            if v >= 95: return "background-color:#d5f5e3;color:#1e8449;font-weight:700"
            if v >= 90: return "background-color:#fef9e7;color:#d68910;font-weight:700"
            return              "background-color:#fdf2f2;color:#c0392b;font-weight:700"
        except Exception: return ""

    def style_mttr(val):
        try:
            v  = float(val)
            mx2 = float(ma_disp["Total MTTR (h)"].max()) or 1.0
            ratio = min(v / mx2, 1.0)
            g   = int(255 - ratio * 160)
            b2  = int(255 - ratio * 210)
            txt = "#7a2005" if ratio > 0.6 else "#4a3020"
            return f"background-color:rgb(255,{g},{b2});color:{txt};font-weight:{'700' if ratio>0.6 else '400'}"
        except Exception: return ""

    st.dataframe(
        ma_disp.style
            .applymap(style_dispo, subset=["Availability (%)"])
            .applymap(style_mttr,  subset=["Total MTTR (h)"])
            .format({"Total MTTR (h)":"{:.4f}","Total MTBF (h)":"{:.2f}",
                     "Availability (%)":"{:.1f}%"}),
        use_container_width=True, hide_index=True
    )

    st.markdown('<div class="te-section">Data Export</div>', unsafe_allow_html=True)
    today_str = datetime.now().strftime("%Y%m%d_%H%M")

    # PDF builder
    def build_pdf(df_in: pd.DataFrame, kpi_d: dict,
                  ma_table: pd.DataFrame, pareto_df: pd.DataFrame) -> bytes:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors as rl_colors
        from reportlab.lib.units import cm
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                         Table, TableStyle, HRFlowable, PageBreak)
        from reportlab.lib.colors import HexColor

        buf_pdf = io.BytesIO()
        W, _ = A4
        OR  = HexColor("#E8650A"); NV = HexColor("#1B2A4A")
        WH  = rl_colors.white;    BR = HexColor("#A07858")
        BH  = HexColor("#2A1A0A"); BC = HexColor("#FFF8F2")
        GR  = HexColor("#d5f5e3"); AM = HexColor("#fef9e7"); RD = HexColor("#fdf2f2")

        sty  = getSampleStyleSheet()
        s_h1 = ParagraphStyle("h1", parent=sty["Heading1"], fontSize=22,
                               textColor=WH, fontName="Helvetica-Bold", spaceAfter=4)
        s_sb = ParagraphStyle("sb", parent=sty["Normal"], fontSize=9, textColor=BR, leading=13)
        s_sc = ParagraphStyle("sc", parent=sty["Normal"], fontSize=10, textColor=OR,
                               fontName="Helvetica-Bold", spaceAfter=4, spaceBefore=12)

        story = []

        # ── Cover ──
        story.append(Paragraph("TE CONNECTIVITY", ParagraphStyle(
            "br", fontSize=10, textColor=OR, fontName="Helvetica-Bold",
            leading=13, spaceAfter=2, letterSpacing=3)))
        story.append(Paragraph("TPM KPI MAINTENANCE REPORT", s_h1))
        story.append(Paragraph("Stamping Dept — Bruderer Presses · Tangier Plant 1310", s_sb))
        story.append(Spacer(1, 0.3*cm))
        story.append(HRFlowable(width="100%", thickness=2, color=OR, spaceAfter=12))

        cov = [
            ["Indicator", "Value"],
            ["Global Availability",   f"{kpi_d['dispo']:.2f}%"],
            ["Avg MTTR / Stop",        f"{kpi_d['mttr_mean_h']:.4f} h  ({round(kpi_d['mttr_mean_h']*60,1)} min)"],
            ["Avg MTBF",               f"{kpi_d['mtbf_mean_h']:.2f} h"],
            ["Total Stops",            str(kpi_d['nb_arrets'])],
            ["Total Events",           f"{kpi_d['nb_rows']:,}"],
            ["Report Generated",       datetime.now().strftime("%m/%d/%Y at %H:%M")],
        ]
        cw2 = [(W-4*cm)/2]*2
        ct = Table(cov, colWidths=cw2, repeatRows=1)
        ct.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),NV), ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("TEXTCOLOR",(0,0),(-1,0),WH),  ("FONTSIZE",(0,0),(-1,-1),9),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),("TOPPADDING",(0,0),(-1,-1),7),
            ("BOTTOMPADDING",(0,0),(-1,-1),7),
            ("BOX",(0,0),(-1,-1),0.5,HexColor("#EDE0D4")),
            ("INNERGRID",(0,0),(-1,-1),0.3,HexColor("#F0E0D0")),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[BC,WH]*10),
        ]))
        story.append(ct)
        story.append(PageBreak())

        # ── Summary by Machine ──
        story.append(Paragraph("── SUMMARY TABLE BY MACHINE", s_sc))
        story.append(HRFlowable(width="100%", thickness=1, color=OR, spaceAfter=8))
        hd = [["Machine","Total MTTR (h)","Total MTBF (h)","Events","Availability (%)"]]
        for _, row in ma_table.iterrows():
            d = float(row["Availability (%)"])
            hd.append([str(row["Machine"]),
                        f"{float(row['Total MTTR (h)']):.4f}",
                        f"{float(row['Total MTBF (h)']):.2f}",
                        str(int(row["Events"])), f"{d:.1f}%"])
        cw5 = [(W-4*cm)/5]*5
        mt2 = Table(hd, colWidths=cw5, repeatRows=1)
        rs2 = [
            ("BACKGROUND",(0,0),(-1,0),BH), ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("TEXTCOLOR",(0,0),(-1,0),WH),  ("FONTSIZE",(0,0),(-1,-1),9),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),("TOPPADDING",(0,0),(-1,-1),6),
            ("BOTTOMPADDING",(0,0),(-1,-1),6),
            ("BOX",(0,0),(-1,-1),0.5,HexColor("#EDE0D4")),
            ("INNERGRID",(0,0),(-1,-1),0.3,HexColor("#F0E0D0")),
        ]
        for ri, row in enumerate(hd[1:], start=1):
            try:
                d = float(row[4].replace("%",""))
                bg = GR if d>=95 else AM if d>=90 else RD
                rs2.append(("BACKGROUND",(4,ri),(4,ri),bg))
            except Exception: pass
        mt2.setStyle(TableStyle(rs2))
        story.append(mt2)
        story.append(Spacer(1, 14))

        # ── Pareto ──
        if not pareto_df.empty:
            story.append(Paragraph("── DOWNTIME PARETO", s_sc))
            story.append(HRFlowable(width="100%", thickness=1, color=OR, spaceAfter=8))
            pd2 = [["Machine","Total MTTR (h)","Part (%)","Cumul (%)"]]
            for _, row in pareto_df.iterrows():
                pd2.append([str(row["Machine"]),
                             f"{float(row['MTTR_total_h']):.3f}",
                             f"{float(row['Pct']):.1f}%",
                             f"{float(row['Cumul']):.1f}%"])
            pt2 = Table(pd2, colWidths=[(W-4*cm)/4]*4, repeatRows=1)
            pt2.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),NV), ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("TEXTCOLOR",(0,0),(-1,0),WH),  ("FONTSIZE",(0,0),(-1,-1),9),
                ("ALIGN",(0,0),(-1,-1),"CENTER"),("TOPPADDING",(0,0),(-1,-1),6),
                ("BOTTOMPADDING",(0,0),(-1,-1),6),
                ("BOX",(0,0),(-1,-1),0.5,HexColor("#EDE0D4")),
                ("INNERGRID",(0,0),(-1,-1),0.3,HexColor("#F0E0D0")),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[BC,WH]*20),
            ]))
            story.append(pt2)
            story.append(Spacer(1, 14))

        # ── Qualified Stops (Shift + Key Failure + User ID) ──
        if all(c in df_in.columns for c in ["Shift","Key Failure"]):
            q_df = df_in[
                (df_in["mttr_h"] > 0) &
                df_in[["Shift","Key Failure"]].apply(
                    lambda r: (str(r["Shift"]).strip() not in ("","None","nan") or
                               str(r["Key Failure"]).strip() not in ("","None","nan")), axis=1)
            ]
            if not q_df.empty:
                story.append(Paragraph("── QUALIFIED STOPS — USER INPUT", s_sc))
                story.append(HRFlowable(width="100%", thickness=1, color=OR, spaceAfter=8))
                q_cols = [c for c in [COL_MACHINE, COL_DATE, COL_STATUS,
                                       "Shift","Key Failure","User ID","mttr_h"] if c in q_df.columns]
                q_hdrs = {COL_MACHINE:"Machine", COL_DATE:"Date",
                           COL_STATUS:"Status", "Shift":"Shift",
                           "Key Failure":"Key Failure", "User ID":"User ID", "mttr_h":"MTTR (h)"}
                qd = [[q_hdrs.get(c,c) for c in q_cols]]
                for _, row in q_df[q_cols].iterrows():
                    rv = []
                    for c in q_cols:
                        v = row[c]
                        if c == "mttr_h":
                            rv.append(f"{float(v):.3f}" if pd.notna(v) else "—")
                        elif c == COL_DATE:
                            try:    rv.append(pd.to_datetime(v).strftime("%m/%d/%Y"))
                            except: rv.append(str(v)[:10])
                        else:
                            rv.append(str(v) if pd.notna(v) else "—")
                    qd.append(rv)
                qcw = [(W-4*cm)/len(q_cols)]*len(q_cols)
                qtbl = Table(qd, colWidths=qcw, repeatRows=1)
                qtbl.setStyle(TableStyle([
                    ("BACKGROUND",(0,0),(-1,0),BH), ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                    ("TEXTCOLOR",(0,0),(-1,0),WH),  ("FONTSIZE",(0,0),(-1,-1),7),
                    ("ALIGN",(0,0),(-1,-1),"CENTER"),("TOPPADDING",(0,0),(-1,-1),4),
                    ("BOTTOMPADDING",(0,0),(-1,-1),4),
                    ("BOX",(0,0),(-1,-1),0.5,HexColor("#EDE0D4")),
                    ("INNERGRID",(0,0),(-1,-1),0.3,HexColor("#F0E0D0")),
                    ("ROWBACKGROUNDS",(0,1),(-1,-1),[BC,WH]*200),
                ]))
                story.append(qtbl)

        def add_footer(canvas_obj, doc):
            canvas_obj.saveState()
            canvas_obj.setFont("Helvetica", 7)
            canvas_obj.setFillColor(BR)
            canvas_obj.drawCentredString(W/2, 1.5*cm,
                f"≡ TE CONNECTIVITY · STAMPING DEPT · TANGIER   |   "
                f"TPM KPI DASHBOARD   |   {datetime.now().strftime('%m/%d/%Y')}   |   "
                f"Page {doc.page}")
            canvas_obj.setStrokeColor(OR)
            canvas_obj.setLineWidth(1.5)
            canvas_obj.line(2*cm, 2*cm, W-2*cm, 2*cm)
            canvas_obj.restoreState()

        doc = SimpleDocTemplate(buf_pdf, pagesize=A4,
                                leftMargin=2*cm, rightMargin=2*cm,
                                topMargin=1.8*cm, bottomMargin=2.8*cm)
        doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
        return buf_pdf.getvalue()

    ec1, ec2, ec3 = st.columns(3)
    with ec1:
        try:
            _pdf_bytes = build_pdf(df, kpi, ma_disp, pareto)
            st.download_button(
                "⬇ DOWNLOAD PDF REPORT", data=_pdf_bytes,
                file_name=f"TE_TPM_{today_str}.pdf", mime="application/pdf",
                use_container_width=True)
        except Exception as _e:
            st.warning(f"PDF: `pip install reportlab` ({_e})")
    with ec2:
        try:
            _xl = export_excel(df, kpi)
            st.download_button(
                "⬇ EXCEL MULTI-SHEET", data=_xl,
                file_name=f"TE_KPI_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        except Exception as _e:
            st.warning(f"Excel: {_e}")
    with ec3:
        st.download_button(
            "⬇ CSV PARETO",
            data=pareto.to_csv(index=False, sep=";").encode("utf-8"),
            file_name=f"TE_pareto_{today_str}.csv",
            mime="text/csv", use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  TAB 2 — QUALIFICATION DES ARRÊTS  (filtres + User ID + persistance robuste)
# ═══════════════════════════════════════════════════════════════════════════════
#  TAB 2 — QUALIFICATION DES ARRÊTS
# ═══════════════════════════════════════════════════════════════════════════════
with tab_qual:

    # ─────────────────────────────────────────────────────────────────────────
    # Helper
    # ─────────────────────────────────────────────────────────────────────────
    def _is_qualified(r):
        return any(str(r.get(c, "")).strip() not in ("", "None", "nan")
                   for c in ["Shift", "Key Failure"])

    # Compteurs globaux (calculés sur df filtré sidebar)
    _qual_n = int(df[["Shift","Key Failure"]].apply(_is_qualified, axis=1).sum()) \
        if all(c in df.columns for c in ["Shift","Key Failure"]) else 0
    _stop_n = int((df["mttr_h"] > 0).sum())
    _pct_q  = (_qual_n / _stop_n * 100) if _stop_n > 0 else 0

    # ── Header avec compteur + barre de progression ───────────────────────────
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,{TE_BLACK} 0%,#2A1A0A 100%);
                border:2px solid {TE_ORANGE};border-radius:14px;
                padding:20px 28px;margin-bottom:18px">
      <div style="display:flex;align-items:center;justify-content:space-between;
                  flex-wrap:wrap;gap:16px">
        <div>
          <div style="font-family:'Barlow Condensed',sans-serif;font-size:20px;font-weight:800;
                      color:{TE_ORANGE};letter-spacing:2px;text-transform:uppercase;margin-bottom:6px">
            ✏️ Qualification des Arrêts
          </div>
          <div style="font-family:'JetBrains Mono',monospace;font-size:10px;
                      color:rgba(255,255,255,0.6);line-height:1.9">
            Renseignez votre <strong style="color:{TE_ORANGE2}">User ID</strong>,
            le <strong style="color:{TE_ORANGE2}">Shift</strong>
            et la <strong style="color:{TE_ORANGE2}">Key Failure</strong>
            · Cliquez sur <strong style="color:{TE_ORANGE2}">💾 Enregistrer</strong> pour valider
          </div>
        </div>
        <div style="text-align:right">
          <div style="font-family:'Barlow Condensed',sans-serif;font-size:34px;font-weight:800;
                      color:{TE_ORANGE};line-height:1">
            {_qual_n}<span style="font-size:16px;color:rgba(255,255,255,0.35)">/{_stop_n}</span>
          </div>
          <div style="font-size:9px;color:rgba(255,255,255,0.4);
                      font-family:'JetBrains Mono',monospace;letter-spacing:1px">
            STOPS QUALIFIED
          </div>
        </div>
      </div>
      <div style="margin-top:14px">
        <div style="display:flex;justify-content:space-between;font-size:9px;
                    color:rgba(255,255,255,0.4);font-family:'JetBrains Mono',monospace;
                    margin-bottom:5px">
          <span>Progression de la qualification</span>
          <span style="color:{TE_ORANGE}">{_qual_n} / {_stop_n} ({_pct_q:.0f}%)</span>
        </div>
        <div style="background:rgba(255,255,255,0.1);border-radius:4px;height:7px;overflow:hidden">
          <div style="background:linear-gradient(90deg,{TE_ORANGE},{TE_DARK});height:100%;
                      width:{min(_pct_q,100):.1f}%;border-radius:4px;
                      transition:width 0.4s ease"></div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Filtres : Machine ID + Date précise (filtre Statut supprimé) ──────────
    st.markdown(f"""
    <div style="background:{TE_WHITE};border:1px solid #EDE0D4;
                border-left:4px solid {TE_ORANGE};border-radius:10px;
                padding:14px 18px;margin-bottom:14px">
      <div style="font-family:'JetBrains Mono',monospace;font-size:9px;font-weight:700;
                  letter-spacing:2.5px;text-transform:uppercase;
                  color:{TE_ORANGE};margin-bottom:10px">
        🔍 Recherche Rapide
      </div>
    """, unsafe_allow_html=True)

    _df_stops = df[df["mttr_h"] > 0].copy()
    _fcol1, _fcol2 = st.columns(2)

    with _fcol1:
        _machines_avail = ["Toutes"] + sorted(
            _df_stops[COL_MACHINE].dropna().unique().tolist())
        _filter_machine = st.selectbox(
            "🏭 Machine ID",
            options=_machines_avail, index=0,
            key="q_filter_machine",
            help="Filtrer par machine spécifique")

    with _fcol2:
        _dates_raw    = pd.to_datetime(_df_stops[COL_DATE], errors="coerce").dropna()
        _dates_avail  = sorted(_dates_raw.dt.date.unique())
        _date_options = ["Toutes"] + [d.strftime("%m/%d/%Y") for d in _dates_avail]
        _filter_date_str = st.selectbox(
            "📅 Date précise",
            options=_date_options, index=0,
            key="q_filter_date",
            help="Filtrer par date de shift exacte")

    st.markdown("</div>", unsafe_allow_html=True)

    # ── Construction du df filtré ─────────────────────────────────────────────
    # Ordre colonnes : Hydra read-only | User ID | Shift | Key Failure
    _display_cols = [c for c in [
        COL_MACHINE, COL_DATE, COL_STATUS, "mttr_h",  # ← lecture seule
        "User ID", "Shift", "Key Failure"             # ← éditables dans cet ordre
    ] if c in df.columns]

    _df_base = _df_stops.copy()   # index = index original de df / session_state

    if _filter_machine != "Toutes":
        _df_base = _df_base[_df_base[COL_MACHINE] == _filter_machine]

    if _filter_date_str != "Toutes":
        _target_date = pd.to_datetime(_filter_date_str, format="%m/%d/%Y", errors="coerce")
        if pd.notna(_target_date):
            _df_base = _df_base[
                pd.to_datetime(_df_base[COL_DATE], errors="coerce").dt.date
                == _target_date.date()]

    # Capturer les indices AVANT de formater les dates
    _orig_idx = _df_base.index.values   # clés dans session_state.edited_df

    # Formatage date pour l'affichage
    _df_show = _df_base[_display_cols].copy()
    if COL_DATE in _df_show.columns:
        _df_show[COL_DATE] = (
            pd.to_datetime(_df_show[COL_DATE], errors="coerce")
            .dt.strftime("%m/%d/%Y").fillna("—"))
    _df_show = _df_show.reset_index(drop=True)

    # Compteur de lignes
    _n_shown    = len(_df_show)
    _is_filtered = _filter_machine != "Toutes" or _filter_date_str != "Toutes"
    st.markdown(
        f'<div style="font-family:\'JetBrains Mono\',monospace;font-size:9px;'
        f'color:#9A7A60;margin-bottom:8px;letter-spacing:1px">'
        f'Affichage : <strong style="color:{TE_ORANGE}">{_n_shown}</strong>'
        f' arrêt{"s" if _n_shown != 1 else ""}'
        f'{"  ·  filtre actif / " + str(_stop_n) + " stops totaux" if _is_filtered else "  ·  " + str(_stop_n) + " stops totaux"}'
        f'</div>',
        unsafe_allow_html=True)

    if _df_show.empty:
        st.info("Aucun arrêt ne correspond aux filtres sélectionnés.")
    else:
        # ── Tableau éditable ──────────────────────────────────────────────────
        _edited = st.data_editor(
            _df_show,
            use_container_width=True,
            height=min(600, max(200, _n_shown * 38 + 62)),
            num_rows="fixed",
            column_order=_display_cols,    # ordre explicite forcé
            column_config={
                # ── Lecture seule — données Hydra ──────────────────────────
                COL_MACHINE: st.column_config.TextColumn(
                    "🏭 Machine", disabled=True, width="small"),
                COL_DATE: st.column_config.TextColumn(
                    "📅 Date", disabled=True, width="small"),
                COL_STATUS: st.column_config.TextColumn(
                    "⚙ Status", disabled=True, width="medium"),
                "mttr_h": st.column_config.NumberColumn(
                    "⏱ MTTR (h)", format="%.4f",
                    disabled=True, width="small"),
                # ── Éditables : ordre 1→2→3 ────────────────────────────────
                "User ID": st.column_config.TextColumn(
                    "👤 User ID",
                    disabled=False, width="small", max_chars=20,
                    help="Votre matricule / identifiant technicien"),
                "Shift": st.column_config.SelectboxColumn(
                    "🔄 Shift",
                    options=SHIFTS, required=False, width="small",
                    help="A (6-14h) · B (14-22h) · C (22-6h)"),
                "Key Failure": st.column_config.SelectboxColumn(
                    "🔧 Key Failure",
                    options=KEY_FAILURES, required=False, width="large",
                    help="Cause racine de l'arrêt"),
            },
            key="qual_editor_v5"
        )

        # ── Bouton 💾 Enregistrer (centré, pleine largeur centrale) ──────────
        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
        _sb_left, _sb_mid, _sb_right = st.columns([1.5, 2, 1.5])
        with _sb_mid:
            _save_clicked = st.button(
                "💾  ENREGISTRER LES MODIFICATIONS",
                type="primary",
                use_container_width=True,
                key="btn_save_qual",
                help="Valider et sauvegarder les saisies du tableau ci-dessus")

        # ── Persistance : uniquement au clic Enregistrer ─────────────────────
        if _save_clicked and _edited is not None:
            _n_saved = 0
            for _col in ["User ID", "Shift", "Key Failure"]:
                if _col not in _edited.columns:
                    continue
                _new_vals = _edited[_col].values
                _old_vals = st.session_state.edited_df.loc[_orig_idx, _col].values
                try:
                    _diff = ~(
                        (pd.isnull(_new_vals) & pd.isnull(_old_vals)) |
                        (np.array(_new_vals, dtype=object) == np.array(_old_vals, dtype=object))
                    )
                    for _k, _orig_i in enumerate(_orig_idx):
                        if _diff[_k]:
                            st.session_state.edited_df.at[_orig_i, _col] = _new_vals[_k]
                            _n_saved += 1
                except Exception:
                    st.session_state.edited_df.loc[_orig_idx, _col] = _new_vals
                    _n_saved += int(_diff.sum()) if '_diff' in dir() else len(_orig_idx)

            # Stocker le résultat pour affichage après rerun
            st.session_state["_save_result"] = _n_saved
            st.rerun()

        # ── Message succès/info (rendu après rerun, avant le tableau) ─────────
        # Note : placé ici car Streamlit rend de haut en bas après rerun.
        # Le message s'affiche juste sous le bouton.
        if st.session_state.get("_save_result") is not None:
            _nr = st.session_state.pop("_save_result")
            if _nr > 0:
                st.markdown(f"""
                <div style="background:#eafaf1;border:1.5px solid #a9dfbf;
                            border-left:5px solid {TE_GREEN};border-radius:10px;
                            padding:14px 20px;margin-top:6px;
                            display:flex;align-items:center;gap:14px">
                  <span style="font-size:24px">✅</span>
                  <div>
                    <div style="font-family:'Barlow Condensed',sans-serif;font-size:15px;
                                font-weight:800;color:#1e8449;text-transform:uppercase;
                                letter-spacing:1px;margin-bottom:3px">
                      Enregistrement réussi
                    </div>
                    <div style="font-size:12px;color:#145a32;line-height:1.6">
                      <strong>{_nr}</strong> cellule{"s" if _nr != 1 else ""}
                      modifiée{"s" if _nr != 1 else ""} sauvegardée{"s" if _nr != 1 else ""}
                      dans la session.
                      Les données sont prêtes pour l'export PDF et CSV.
                    </div>
                  </div>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div style="background:#fff8e1;border:1.5px solid #ffe082;
                            border-left:5px solid {TE_AMBER};border-radius:10px;
                            padding:12px 18px;margin-top:6px;
                            display:flex;align-items:center;gap:10px">
                  <span style="font-size:18px">ℹ️</span>
                  <span style="font-size:12px;color:#5d4037">
                    Aucune modification détectée par rapport au dernier enregistrement.
                  </span>
                </div>
                """, unsafe_allow_html=True)

    # ── Export des arrêts qualifiés ───────────────────────────────────────────
    if _qual_n > 0:
        st.markdown(f'<div class="te-section">Export des Arrêts Qualifiés</div>',
                    unsafe_allow_html=True)
        _q_all = df[df[["Shift","Key Failure"]].apply(_is_qualified, axis=1)]
        _q_exp_cols = [c for c in [
            COL_MACHINE, COL_DATE, COL_STATUS,
            "mttr_h", "User ID", "Shift", "Key Failure"
        ] if c in _q_all.columns]
        _ts = datetime.now().strftime("%Y%m%d_%H%M")
        _xc1, _xc2 = st.columns(2)
        with _xc1:
            st.download_button(
                f"⬇ CSV — {_qual_n} ARRÊTS QUALIFIÉS",
                data=_q_all[_q_exp_cols].to_csv(index=False, sep=";").encode("utf-8"),
                file_name=f"TE_arrets_qualifies_{_ts}.csv",
                mime="text/csv", use_container_width=True)
        with _xc2:
            st.info(f"📄 **{_qual_n}** arrêts qualifiés — le rapport PDF complet "
                    f"(avec ce tableau) est disponible dans l'onglet **📊 KPIs**.")

# ─────────────────────────────────────────────────────────────────────────────
#  FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="text-align:center;font-family:'JetBrains Mono',monospace;font-size:9px;
            letter-spacing:2px;color:#C0A080;padding:24px 0 12px;
            border-top:1px solid #E0D0C0;margin-top:32px">
    ≡ TE CONNECTIVITY · STAMPING DEPARTMENT · TANGIER<br>
    TPM KPI DASHBOARD · {datetime.now().strftime('%m/%d/%Y')}
</div>
""", unsafe_allow_html=True)
