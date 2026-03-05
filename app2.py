# ----------------------------------------------------------------------------------
# CORPORATE RENEWAL TRACKER — C&P ONLY — MINET HEADER + BRIGHT GANTT (no oval caps)
# ----------------------------------------------------------------------------------

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from pathlib import Path
from base64 import b64encode

# --- PAGE CONFIG ---
st.set_page_config(page_title="Corporate Plan to Make", layout="wide")

SOURCE_FILE = "War Room (Plan To Make Plan) (1).xlsx"
OUTPUT_FILE = "War Room (Plan To Make Plan) (1) - UPDATED.xlsx"
# IMPORTANT: use the exact sheet name from Excel (no HTML entities)
SHEET_NAME  = "Corporates & Parastatals"
LOGO_FILE   = "logo.png"   # your logo in project folder

# --- BRAND COLOURS ---
MINET_RED   = "#e60000"
MINET_BLACK = "#000000"
MINET_WHITE = "#FFFFFF"
MINET_GREY  = "#404040"
CARD_BG     = "#f5f5f5"

# ========================================================================================
#  HEADER (logo left, title right, grey card, red accent bar)
# ========================================================================================

def render_minet_header(logo_path: str, title_text: str, tagline: str = ""):
    if Path(logo_path).exists():
        with open(logo_path, "rb") as f:
            logo_b64 = b64encode(f.read()).decode()
        logo_html = f"data:image/png;base64,{logo_b64}"
    else:
        logo_html = ""
        st.warning(f"Logo not found: {logo_path}")

    custom_css = f"""
    <style>
      .minet-card {{
        background: {CARD_BG};
        border-left: 12px solid {MINET_RED};
        border-radius: 10px;
        padding: 14px 22px;
        margin: 6px 0 16px 0;
        box-shadow: 0 1px 2px rgba(0,0,0,0.06);
      }}
      .minet-row {{
        display: flex;
        align-items: center;
        gap: 22px;
      }}
      .minet-logo {{
        height: 55px;
        width: auto;
      }}
      .minet-tagline {{
        font-size: 12px;
        color: {MINET_GREY};
        margin-left: 4px;
      }}
      .minet-title {{
        font-size: 32px;
        font-weight: 800;
        color: {MINET_BLACK} !important;
        margin: 0;
        text-shadow: none;
      }}
    </style>
    """

    html = f"""
    <div class="minet-card">
      <div class="minet-row">
        <div>
          <div class="minet-tagline">{tagline}</div>
          {'<img class="minet-logo" src="' + logo_html + '"/>' if logo_html else ''}
        </div>
        <div>
          <div class="minet-title">{title_text}</div>
        </div>
      </div>
    </div>
    """

    st.markdown(custom_css + html, unsafe_allow_html=True)

render_minet_header(LOGO_FILE, "Corporate Plan to Make", tagline="")

# ========================================================================================
#  HELPERS
# ========================================================================================

def clean_header_list(cols):
    return [str(c).replace("\n", " ").strip() for c in cols]

def pick_any(df, names):
    for n in names:
        if n in df.columns:
            return df[n]
    lower_map = {c.lower(): c for c in df.columns}
    for n in names:
        if n.lower() in lower_map:
            return df[lower_map[n.lower()]]
    return pd.Series([None]*len(df))

def to_number(series_like):
    s = pd.Series(series_like, copy=True).astype(str).str.strip()
    s = s.str.replace(",", "", regex=False)
    s = s.str.replace(" ", "", regex=False)
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)  # (1,234) -> -1234
    return pd.to_numeric(s, errors="coerce")

# ========================================================================================
#  LOAD C&P SHEET
# ========================================================================================

def load_cp_sheet(path: str):
    raw = pd.read_excel(path, sheet_name=SHEET_NAME, header=None, engine="openpyxl")
    raw = raw.dropna(how="all").dropna(axis=1, how="all")

    header = None
    for i in range(min(25, len(raw))):
        row = " ".join(raw.iloc[i].astype(str).str.lower().tolist())
        if "prospect" in row and ("by when" in row or "due" in row):
            header = i
            break
    if header is None:
        st.error("Could not find header row for the C&P sheet.")
        st.stop()

    df = raw.copy()
    df.columns = clean_header_list(df.iloc[header])
    df = df.iloc[header+1:].reset_index(drop=True)
    df.columns = clean_header_list(df.columns)

    out = pd.DataFrame({
        "Prospect": pick_any(df, ["Prospect"]),
        "Action": pick_any(df, ["Action"]),
        "Owner": pick_any(df, ["By Whom", "Owner"]),
        "Start Date": pd.to_datetime(pick_any(df, ["Start Date"]), errors="coerce"),
        "Due Date": pd.to_datetime(pick_any(df, ["By When","Due Date"]), errors="coerce"),
        "Warning Status": pick_any(df, ["Warning status","Warning Status"]).astype(str).str.strip(),
        "Status %": pd.to_numeric(pick_any(df, ["Status Percent","Status %"]), errors="coerce"),
        "Est. Income (100%)": to_number(pick_any(df, ["Estimated Converted Income (100%)"])),
        "Prob. Adj. Income": to_number(pick_any(df, ["Probability Adjusted Income"])),
        "Probability": pd.to_numeric(pick_any(df, ["Probability"]), errors="coerce"),
        "Current Milestone": pick_any(df, ["Current Milestone"]),
        "Conversion Status": pick_any(df, ["Conversion Status"]),
        "Comment": pick_any(df, ["Comment"])
    })

    out = out[(out["Prospect"].notna()) | (out["Start Date"].notna()) | (out["Due Date"].notna())]
    out["Warning Status"] = out["Warning Status"].replace({
        "ok":"OK", "complete":"Complete", "warning":"Warning"
    })

    return out.reset_index(drop=True)

df = load_cp_sheet(SOURCE_FILE)
if "cp_df" not in st.session_state:
    st.session_state["cp_df"] = df.copy()

# ========================================================================================
#  SIDEBAR FILTERS
# ========================================================================================

st.sidebar.header("Filters")

show_sections = st.sidebar.multiselect("Show Sections", ["Table","Gantt"], default=["Table","Gantt"])

color_by = st.sidebar.radio("Color Gantt By", ["Warning Status","Conversion Status"], index=0)

# ========================================================================================
#  TABLE (EDITABLE)
# ========================================================================================

if "Table" in show_sections:
    edited = st.data_editor(
        st.session_state["cp_df"],
        hide_index=True,
        num_rows="dynamic",
        use_container_width=True,
        key="cp_data_editor"
    )

    st.session_state["cp_df"] = edited.reset_index(drop=True)
    working_df = st.session_state["cp_df"].copy()

    st.download_button(
        "Download CSV",
        working_df.to_csv(index=False).encode(),
        "cp_updated.csv",
        "text/csv"
    )

    if st.button("Save updated Excel"):
        working_df.to_excel(OUTPUT_FILE, sheet_name=SHEET_NAME, index=False)
        st.success(f"Saved: {OUTPUT_FILE}")

# ========================================================================================
#  GANTT — BRIGHT RED/GREEN/BLUE (NO OVAL CAPS)
# ========================================================================================

if "Gantt" in show_sections:
    st.subheader("Gantt Chart")

    gantt_df = st.session_state["cp_df"].copy()
    gantt_df["Start Date"] = pd.to_datetime(gantt_df["Start Date"], errors="coerce")
    gantt_df["Due Date"] = pd.to_datetime(gantt_df["Due Date"], errors="coerce")
    bars = gantt_df.dropna(subset=["Start Date","Due Date"])

    # ---- Bright colours ----
    if color_by == "Warning Status":
        color_map = {
            "Complete": "#3498db",   # bright blue
            "OK": "#2ecc71",         # bright green
            "Warning": "#ff3030"     # bright red
        }
    else:
        uniq = gantt_df["Conversion Status"].fillna("Unknown").astype(str).unique()
        palette = ["#3498db","#2ecc71","#ff3030","#9b59b6","#f1c40f"]
        color_map = {u: palette[i%len(palette)] for i,u in enumerate(sorted(uniq))}

    # ---- Base Gantt (rectangle bars only) ----
    if not bars.empty:
        fig = px.timeline(
            bars,
            x_start="Start Date", x_end="Due Date",
            y="Prospect",
            color=color_by,
            color_discrete_map=color_map,
            hover_data={
                "Prospect":True,"Owner":True,"Warning Status":True,
                "Conversion Status":True,"Probability":":.0f",
                "Est. Income (100%)":":,.0f","Prob. Adj. Income":":,.0f",
                "Comment":True
            }
        )
    else:
        fig = go.Figure()

    # NOTE: No round end-caps added here (oval effect removed)

    # ---- Styling ----
    fig.update_yaxes(autorange="reversed")
    fig.update_layout(
        height=500,
        margin=dict(l=10,r=10,t=10,b=10),
        bargap=0.25,
        template="plotly_white",
        plot_bgcolor="#f3f6fb",
        paper_bgcolor="#f3f6fb",
        font=dict(color="black"),
        legend=dict(
            title=dict(font=dict(color="black")),
            font=dict(color="black"),
            bgcolor="rgba(0,0,0,0)"
        ),
        xaxis=dict(
            tickfont=dict(color="black"),
            title=dict(font=dict(color="black")),
            gridcolor="#d9e1ec",
            zerolinecolor="#d9e1ec"
        ),
        yaxis=dict(
            tickfont=dict(color="black"),
            title=dict(font=dict(color="black")),
            gridcolor="#d9e1ec",
            zerolinecolor="#d9e1ec"
        )
    )

    # Today marker
    today = datetime.now()
    fig.add_vline(x=today, line_dash="dash", line_color="black")
    fig.add_annotation(x=today, y=1.03, xref="x", yref="paper", text="Today",
                       showarrow=False, font=dict(color="black"))

    st.plotly_chart(fig, use_container_width=True)
