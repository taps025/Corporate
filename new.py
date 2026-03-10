import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
from pathlib import Path
from base64 import b64encode

st.set_page_config(page_title="Clients by Segment Dashboard", layout="wide")

EXCEL_FILE = "Clients by segment.xlsx"
SHEET_NAME = "2a"
LOGO_FILE = "logo.png"


def _txt(v) -> str:
    if pd.isna(v):
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def _to_amount(v) -> float:
    s = _txt(v).upper().replace("BWP", "")
    s = s.replace(",", "").replace(" ", "")
    if not s:
        return np.nan
    s = re.sub(r"^\((.*)\)$", r"-\1", s)
    return pd.to_numeric(s, errors="coerce")


def _to_client_status(v) -> str:
    s = _txt(v).lower()
    if s == "yes":
        return "Yes"
    if s == "no":
        return "No"
    return "Unknown"

def _normalize_business_name(raw_business: str) -> str:
    s = _txt(raw_business)
    if not s:
        return ""
    m = re.match(r"^\d+\.\s*(.+)$", s)
    if m:
        return m.group(1).strip()
    if re.fullmatch(r"\d+\.?", s):
        n = re.sub(r"\D", "", s)
        return f"Business {n} (name not provided in source)"
    return s


def render_header(logo_path: str, title_text: str):
    if Path(logo_path).exists():
        with open(logo_path, "rb") as f:
            logo_b64 = b64encode(f.read()).decode()
        logo_html = f"data:image/png;base64,{logo_b64}"
    else:
        logo_html = ""

    st.markdown(
        f"""
        <style>
          .minet-card {{
            background: #f0f0f0;
            border-left: 12px solid #e60000;
            border-radius: 12px;
            padding: 14px 22px;
            margin: 6px 0 16px 0;
          }}
          .minet-row {{
            display: flex;
            align-items: center;
            gap: 22px;
          }}
          .minet-logo {{
            height: 70px;
            width: auto;
          }}
          .minet-title {{
            font-size: 42px;
            font-weight: 800;
            color: #000000 !important;
            margin: 0;
            line-height: 1.05;
            white-space: nowrap;
          }}
        </style>
        <div class="minet-card">
          <div class="minet-row">
            <div>{'<img class="minet-logo" src="' + logo_html + '"/>' if logo_html else ''}</div>
            <div class="minet-title">{title_text}</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data
def parse_sheet_2a(path: str, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, engine="openpyxl")
    rows, cols = raw.shape
    records = []

    for r in range(rows):
        for c in range(cols):
            if _txt(raw.iat[r, c]).lower() != "segment name:":
                continue

            segment = "Unknown"
            for cc in range(c + 1, min(c + 7, cols)):
                cand = _txt(raw.iat[r, cc])
                if cand and cand.lower() not in {
                    "segment name:",
                    "name of business",
                    "aon client?",
                    "estimated premium",
                    "yes/no",
                }:
                    segment = cand
                    break

            header_row = None
            for rr in range(r + 1, min(r + 8, rows)):
                if "name of business" in _txt(raw.iat[rr, c]).lower():
                    header_row = rr
                    break
            if header_row is None:
                continue

            aon_col = None
            prem_col = None
            for cc in range(c, min(c + 8, cols)):
                h = _txt(raw.iat[header_row, cc]).lower()
                if "aon client" in h:
                    aon_col = cc
                if "estimated premium" in h:
                    prem_col = cc

            if aon_col is None and header_row + 1 < rows:
                for cc in range(c, min(c + 8, cols)):
                    if "yes/no" in _txt(raw.iat[header_row + 1, cc]).lower():
                        aon_col = cc
                        break

            start = header_row + 1
            if aon_col is not None and "yes/no" in _txt(raw.iat[start, aon_col]).lower():
                start += 1

            end = rows
            for rr in range(start, rows):
                neighborhood = [_txt(raw.iat[rr, cc]).lower() for cc in range(c, min(c + 8, cols))]
                if "segment name:" in neighborhood:
                    end = rr
                    break

            for rr in range(start, end):
                business_primary = _txt(raw.iat[rr, c])
                business_alt = _txt(raw.iat[rr, c + 1]) if c + 1 < cols else ""

                if re.fullmatch(r"\d+\.?", business_primary) and business_alt:
                    business = business_alt
                elif business_primary:
                    business = business_primary
                else:
                    business = business_alt

                if not business:
                    continue
                if "name of business" in business.lower():
                    continue

                aon_val = _txt(raw.iat[rr, aon_col]) if aon_col is not None else ""
                premium_val = raw.iat[rr, prem_col] if prem_col is not None else np.nan
                amount = _to_amount(premium_val)

                clean_business = _normalize_business_name(business)
                if not clean_business:
                    continue
                if clean_business.strip().upper() == "CAAB":
                    amount = 307772.26
                if clean_business.lower().startswith("business ") and not aon_val and pd.isna(amount):
                    continue

                records.append(
                    {
                        "Segment": segment,
                        "Name of Business": clean_business,
                        "Minet Client": _to_client_status(aon_val),
                        "Estimated Premium (BWP)": amount,
                    }
                )

    df = pd.DataFrame(records)
    if df.empty:
        return df

    df = df.drop_duplicates(subset=["Segment", "Name of Business", "Minet Client"]).reset_index(drop=True)
    return df


try:
    data = parse_sheet_2a(EXCEL_FILE, SHEET_NAME)
except Exception as e:
    st.error(f"Failed to load data: {e}")
    st.stop()

if data.empty:
    st.warning("No records found in sheet 2a.")
    st.stop()


st.sidebar.header("Filters")
view_mode = st.sidebar.selectbox("View", ["Revenue Landscape", "Dashboard"], index=0)
header_title = (
    "Corporate & Parastatal Dashboard"
    if view_mode == "Dashboard"
    else "Corporate & Parastatal Revenue Landscape"
)
render_header(LOGO_FILE, header_title)

segments = sorted(data["Segment"].dropna().unique().tolist())
statuses = sorted(data["Minet Client"].dropna().unique().tolist())
statuses = [s for s in statuses if s != "Unknown"]

selected_segment = st.sidebar.selectbox("Segment", ["All"] + segments, index=0)
selected_status = st.sidebar.selectbox("Minet Client", ["All"] + statuses, index=0)
search = st.sidebar.text_input("Search business")

filtered = data.copy()
filtered = filtered[filtered["Minet Client"] != "Unknown"]
if selected_segment != "All":
    filtered = filtered[filtered["Segment"] == selected_segment]
if selected_status != "All":
    filtered = filtered[filtered["Minet Client"] == selected_status]

if search.strip():
    filtered = filtered[filtered["Name of Business"].str.contains(search.strip(), case=False, na=False)]

if filtered.empty:
    st.warning("No records match the selected filters.")
    st.stop()

if view_mode == "Revenue Landscape":
    segment_order = sorted(filtered["Segment"].dropna().unique().tolist())
    tab_labels = ["All Segments"] + segment_order if selected_segment == "All" else segment_order
    tabs = st.tabs(tab_labels)

    for tab, label in zip(tabs, tab_labels):
        with tab:
            if label == "All Segments":
                seg_df = filtered.copy()
            else:
                seg_df = filtered[filtered["Segment"] == label].copy()
            seg_df = seg_df.sort_values(["Minet Client", "Estimated Premium (BWP)", "Name of Business"], ascending=[True, False, True])

            seg_businesses = len(seg_df)
            seg_minet = int((seg_df["Minet Client"] == "Yes").sum())
            seg_prospects = int((seg_df["Minet Client"] == "No").sum())
            seg_premium = float(seg_df["Estimated Premium (BWP)"].fillna(0).sum())
            seg_minet_income = float(
                seg_df.loc[seg_df["Minet Client"] == "Yes", "Estimated Premium (BWP)"].fillna(0).sum()
            )
            seg_non_minet_income = float(
                seg_df.loc[seg_df["Minet Client"] == "No", "Estimated Premium (BWP)"].fillna(0).sum()
            )

            m1, m2, m3, m5, m6 = st.columns(5)
            m1.metric("Businesses", f"{seg_businesses:,}")
            m2.metric("Minet Clients", f"{seg_minet:,}")
            m3.metric("Prospects", f"{seg_prospects:,}")
            m5.metric("Current Income", f"P {seg_minet_income:,.2f}")
            m6.metric("Estimated Income", f"P {seg_non_minet_income:,.2f}")

            display_df = seg_df.copy()
            if label == "All Segments":
                display_df = display_df[["Segment", "Name of Business", "Minet Client", "Estimated Premium (BWP)"]]
            else:
                display_df = display_df[["Name of Business", "Minet Client", "Estimated Premium (BWP)"]]
            display_df["Estimated Premium (BWP)"] = display_df["Estimated Premium (BWP)"].map(
                lambda x: f"{x:,.2f}" if pd.notna(x) else ""
            )
            display_df = display_df.rename(columns={"Estimated Premium (BWP)": "Estimated Income (Pula)"})
            display_df["Estimated Income (Pula)"] = display_df["Estimated Income (Pula)"].apply(
                lambda x: f"P {x}" if x else x
            )
            st.dataframe(display_df, use_container_width=True, hide_index=True)

            st.download_button(
                f"Download {label} (CSV)",
                seg_df.to_csv(index=False).encode("utf-8"),
                f"clients_by_segment_{label.lower().replace(' ', '_')}.csv",
                "text/csv",
                key=f"dl_{label}",
            )
else:
    total_businesses = len(filtered)
    total_minet = int((filtered["Minet Client"] == "Yes").sum())
    total_non_minet = int((filtered["Minet Client"] == "No").sum())
    total_premium = float(filtered["Estimated Premium (BWP)"].fillna(0).sum())
    total_minet_income = float(
        filtered.loc[filtered["Minet Client"] == "Yes", "Estimated Premium (BWP)"].fillna(0).sum()
    )
    total_non_minet_income = float(
        filtered.loc[filtered["Minet Client"] == "No", "Estimated Premium (BWP)"].fillna(0).sum()
    )

    m1, m2, m3, m5, m6 = st.columns(5)
    m1.metric("Businesses", f"{total_businesses:,}")
    m2.metric("Minet Clients", f"{total_minet:,}")
    m3.metric("Non-Minet", f"{total_non_minet:,}")
    m5.metric("Current Income", f"P {total_minet_income:,.2f}")
    m6.metric("Estimated Income", f"P {total_non_minet_income:,.2f}")

    summary = filtered.groupby("Segment", as_index=False).agg(
        Businesses=("Name of Business", "count"),
        Minet_Clients=("Minet Client", lambda s: (s == "Yes").sum()),
        Total_Premium_Pula=("Estimated Premium (BWP)", "sum"),
    )
    summary = summary.sort_values("Total_Premium_Pula", ascending=False)
    summary["Minet Rate %"] = (summary["Minet_Clients"] / summary["Businesses"] * 100).round(1)

    c1, c2 = st.columns(2)
    with c1:
        fig_premium = px.bar(
            summary,
            x="Segment",
            y="Total_Premium_Pula",
            color="Segment",
            title="Income by Segment",
            text_auto=".2s",
        )
        fig_premium.update_layout(
            showlegend=False,
            xaxis_title="",
            yaxis_title="Pula",
            plot_bgcolor="#f3f6fb",
            paper_bgcolor="#f3f6fb",
            font=dict(color="#111111", size=14),
            title_font=dict(color="#111111", size=28),
            xaxis=dict(
                tickfont=dict(color="#111111", size=12),
                title_font=dict(color="#111111"),
                gridcolor="#c9d3e0",
            ),
            yaxis=dict(
                tickfont=dict(color="#111111", size=12),
                title_font=dict(color="#111111"),
                gridcolor="#c9d3e0",
            ),
        )
        fig_premium.update_traces(textfont=dict(color="#111111", size=18))
        st.plotly_chart(fig_premium, use_container_width=True)
    with c2:
        mix_df = pd.DataFrame(
            {"Minet Client": ["Yes", "No"], "Count": [total_minet, total_non_minet]}
        )
        fig_mix = px.pie(
            mix_df,
            names="Minet Client",
            values="Count",
            title="Client Mix",
            color="Minet Client",
            color_discrete_map={"Yes": "#22c55e", "No": "#ef4444"},
            hole=0.45,
        )
        fig_mix.update_layout(
            plot_bgcolor="#f3f6fb",
            paper_bgcolor="#f3f6fb",
            font=dict(color="#111111", size=14),
            title_font=dict(color="#111111", size=28),
            legend=dict(font=dict(color="#111111", size=13)),
        )
        fig_mix.update_traces(textfont=dict(color="#111111", size=18))
        st.plotly_chart(fig_mix, use_container_width=True)

    d1, d2 = st.columns(2)
    with d1:
        top_businesses = (
            filtered.groupby(["Segment", "Name of Business"], as_index=False)["Estimated Premium (BWP)"]
            .sum()
            .sort_values("Estimated Premium (BWP)", ascending=False)
            .head(10)
        )
        fig_top = px.bar(
            top_businesses.sort_values("Estimated Premium (BWP)"),
            x="Estimated Premium (BWP)",
            y="Name of Business",
            orientation="h",
            color="Segment",
            title="Top 10 Businesses by Income",
            text_auto=".2s",
        )
        fig_top.update_layout(
            xaxis_title="Pula",
            yaxis_title="",
            plot_bgcolor="#f3f6fb",
            paper_bgcolor="#f3f6fb",
            font=dict(color="#111111", size=14),
            title_font=dict(color="#111111", size=28),
            xaxis=dict(
                tickfont=dict(color="#111111", size=12),
                title_font=dict(color="#111111"),
                gridcolor="#c9d3e0",
            ),
            yaxis=dict(
                tickfont=dict(color="#111111", size=12),
                title_font=dict(color="#111111"),
                gridcolor="#c9d3e0",
            ),
            legend=dict(font=dict(color="#111111", size=12)),
        )
        fig_top.update_traces(textfont=dict(color="#111111", size=14))
        st.plotly_chart(fig_top, use_container_width=True)
    with d2:
        fig_rate = px.bar(
            summary.sort_values("Minet Rate %", ascending=False),
            x="Segment",
            y="Minet Rate %",
            color="Minet Rate %",
            color_continuous_scale="Blues",
            title="Minet Conversion Rate by Segment",
            text_auto=True,
        )
        fig_rate.update_layout(
            xaxis_title="",
            yaxis_title="Minet Rate (%)",
            plot_bgcolor="#f3f6fb",
            paper_bgcolor="#f3f6fb",
            font=dict(color="#111111", size=14),
            title_font=dict(color="#111111", size=28),
            xaxis=dict(
                tickfont=dict(color="#111111", size=12),
                title_font=dict(color="#111111"),
                gridcolor="#c9d3e0",
            ),
            yaxis=dict(
                tickfont=dict(color="#111111", size=12),
                title_font=dict(color="#111111"),
                gridcolor="#c9d3e0",
            ),
        )
        fig_rate.update_traces(textfont=dict(color="#111111", size=14))
        st.plotly_chart(fig_rate, use_container_width=True)

