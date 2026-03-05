import base64
import json
import mimetypes
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

# ============================================================
#   PAGE CONFIG
# ============================================================
st.set_page_config(page_title="Corporate Renewal tracker", layout="wide")

# ---------- Brand colours ----------
MINET_RED = "#cc0000"
BLACK = "#111111"
WHITE = "#FFFFFF"

# ---------- Local data files ----------
EXCEL_FILE = "Jenn.xlsx"
EXCEL_SHEET = "Renewal book 2026"
STATUS_FILE = "status_store.json"

STATUS_OPTIONS = ["On going", "Renewed", "Lost", "Not renewing", "Awaiting POP"]
MONTHS = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]
MONTH_TO_NUM = {m: i + 1 for i, m in enumerate(MONTHS)}

# ============================================================
#   GLOBAL STYLES
# ============================================================
st.markdown(
    f"""
    <style>
      :root {{
        --minet-red: {MINET_RED};
        --minet-black: {BLACK};
        --minet-white: {WHITE};
      }}

      .block-container {{
        padding-top: 0.75rem;
        background: var(--background-color);
      }}

      .app-header {{
        display: flex; align-items: center; gap: 14px;
        margin: 0 0 12px 0; padding: 30px 12px;
        border-left: 6px solid var(--minet-red);
        background: #f8f8f8;
        border-radius: 6px;
      }}
      .app-header h1 {{
        font-size: 32px; font-weight: 800; line-height: 1.15;
        color: var(--minet-black); margin: 0;
      }}

      div[data-testid="stMetricLabel"] p {{
        color: var(--text-color) !important;
        opacity: 1 !important;
        font-weight: 700 !important;
      }}
      div[data-testid="stMetricValue"] {{
        color: var(--text-color) !important;
      }}
      div[data-testid="stMetricValue"] * {{
        color: var(--text-color) !important;
        font-weight: 800 !important;
      }}

      .stButton > button {{
        background: var(--minet-red); color: var(--minet-white);
        border-radius: 6px; border: 1px solid var(--minet-red);
        font-weight: 700;
      }}
      .stButton > button:hover {{
        filter: brightness(0.95);
        border: 1px solid #a60000;
      }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
#   HELPERS
# ============================================================
def format_pula(v: float) -> str:
    return "P" + f"{float(v):,.2f}".replace(",", " ")


def renewal_date(year, month):
    return date(int(year), MONTH_TO_NUM[month], 15)


def days_left(d):
    return (d - date.today()).days


def traffic_light(status, dleft):
    if status == "Renewed":
        return "✅"
    if status in ["Lost", "Not renewing"]:
        return "⚪"
    if dleft <= 0:
        return "🟥"
    if dleft <= 30:
        return "🟧"
    if dleft <= 60:
        return "🟨"
    return "🟩"


def to_iso_date(val) -> str:
    if val is None:
        return ""
    if isinstance(val, date):
        return val.isoformat()
    s = str(val).strip()
    if not s or s.lower() in ["nan", "nat"]:
        return ""
    try:
        return pd.to_datetime(s, errors="coerce").date().isoformat()
    except Exception:
        return ""


def find_logo_bytes() -> tuple[str | None, str | None]:
    app_dir = Path(__file__).parent.resolve()
    candidates = []

    folders = [app_dir / "projects", app_dir]
    names = ["logo", "Logo", "LOGO"]
    exts = ["png", "PNG", "jpg", "jpeg", "svg", "webp", "JPG", "JPEG", "SVG", "WEBP"]

    for folder in folders:
        for n in names:
            for ext in exts:
                candidates.append(folder / f"{n}.{ext}")

    for p in candidates:
        if p.exists():
            mime, _ = mimetypes.guess_type(str(p))
            if mime is None:
                mime = "image/png"
            try:
                data = p.read_bytes()
                b64 = base64.b64encode(data).decode("utf-8")
                return f"data:{mime};base64,{b64}", None
            except Exception as ex:
                return None, f"Found logo at {p} but failed to read: {ex}"

    dbg = "Searched for logo in:\n" + "\n".join([str(c) for c in candidates[:12]]) + "\n... (more paths skipped)"
    return None, dbg


def load_status_store() -> dict:
    try:
        with open(STATUS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_status_store(store: dict) -> None:
    with open(STATUS_FILE, "w", encoding="utf-8") as f:
        json.dump(store, f, indent=4)


@st.cache_data(ttl=30)
def load_events_from_excel() -> pd.DataFrame:
    src = pd.read_excel(EXCEL_FILE, sheet_name=EXCEL_SHEET, engine="openpyxl")
    src.columns = [str(c).strip() for c in src.columns]

    missing = [c for c in ["CLIENT NUMBER", "CLIENT NAME"] if c not in src.columns]
    if missing:
        raise ValueError(f"Missing required column(s) in {EXCEL_FILE}: {', '.join(missing)}")

    status_store = load_status_store()
    this_year = date.today().year
    rows = []

    for _, row in src.iterrows():
        client_number = str(row.get("CLIENT NUMBER", "")).strip()
        client_name = str(row.get("CLIENT NAME", "")).strip()
        if not client_number or client_number.lower() == "nan":
            continue

        for month in MONTHS:
            if month not in src.columns:
                continue

            amt = pd.to_numeric(row.get(month), errors="coerce")
            if pd.isna(amt) or float(amt) == 0:
                continue

            key = f"{client_number}_{this_year}_{month}"
            saved = status_store.get(key, {})

            rows.append(
                {
                    "CLIENT NUMBER": client_number,
                    "CLIENT NAME": client_name,
                    "Year": this_year,
                    "Month": month,
                    "Amount": float(amt),
                    "Status": saved.get("status", "On going"),
                    "Notes": saved.get("notes", ""),
                    "Renewed On": saved.get("renewed_on", ""),
                }
            )

    return pd.DataFrame(rows)


# ============================================================
#   HEADER
# ============================================================
logo_data_uri, logo_dbg = find_logo_bytes()

header_html = f"""
<div class='app-header'>
  {'<img src="'+logo_data_uri+'" alt="Logo" style="height:56px; width:auto; object-fit:contain;">' if logo_data_uri else ''}
  <h1>Corporate Renewal tracker</h1>
</div>
"""
st.markdown(header_html, unsafe_allow_html=True)

if not logo_data_uri:
    with st.expander("Logo not found - help", expanded=False):
        st.write("Place **logo.png** inside either `./projects/` or alongside `app.py`.")
        st.code(logo_dbg or "", language="text")

# ============================================================
#   LOAD DATA FROM EXCEL
# ============================================================
try:
    events = load_events_from_excel()
except Exception as e:
    st.error(f"Cannot load {EXCEL_FILE}: {e}")
    st.stop()

if events.empty:
    st.warning("No renewal events were found in the Excel source.")
    st.stop()

events["Amount"] = pd.to_numeric(events["Amount"], errors="coerce")
events["Year"] = pd.to_numeric(events["Year"], errors="coerce").astype("Int64")
events["Renewed On"] = pd.to_datetime(events["Renewed On"], errors="coerce").dt.date

events["Income"] = events["Amount"].apply(format_pula)
events["RenewalDate_internal"] = events.apply(lambda r: renewal_date(r["Year"], r["Month"]), axis=1)
events["DaysLeft_internal"] = events["RenewalDate_internal"].apply(days_left)
events["Light"] = events.apply(lambda r: traffic_light(r["Status"], r["DaysLeft_internal"]), axis=1)

# ============================================================
#   SIDEBAR FILTERS
# ============================================================
st.sidebar.header("Filters")
filter_month = st.sidebar.selectbox("Month", ["All"] + MONTHS)
filter_status = st.sidebar.selectbox("Status", ["All"] + STATUS_OPTIONS)
search = st.sidebar.text_input("Search client")

# ============================================================
#   APPLY FILTERS
# ============================================================
view = events.copy()
view = view[view["Amount"] > 0]

if filter_month != "All":
    view = view[view["Month"] == filter_month]

if filter_status != "All":
    view = view[view["Status"] == filter_status]

if search:
    q = search.lower()
    view = view[
        view["CLIENT NUMBER"].astype(str).str.lower().str.contains(q, na=False)
        | view["CLIENT NAME"].astype(str).str.lower().str.contains(q, na=False)
    ]

# ============================================================
#   KPI CARDS
# ============================================================
c1, c2, c3, c4 = st.columns(4)

total_clients_view = view["CLIENT NUMBER"].nunique()
c1.metric("👥 Total Clients", f"{total_clients_view:,}")

ongoing_view = (view["Status"] == "On going").sum()
c2.metric("📌 Ongoing", f"{ongoing_view:,}")

renewed_view = (view["Status"] == "Renewed").sum()
c3.metric("✅ Renewed", f"{renewed_view:,}")

total_income_view = view["Amount"].sum()
c4.metric("💰 Total Income", format_pula(total_income_view))

# ============================================================
#   EDITABLE TABLE
# ============================================================
cols = ["CLIENT NUMBER", "CLIENT NAME", "Year", "Month", "Income", "Light", "Status", "Notes", "Renewed On"]
tidy = view[cols].copy()

col_cfg = {
    "CLIENT NUMBER": st.column_config.TextColumn(label="Client Number"),
    "CLIENT NAME": st.column_config.TextColumn(label="Client Name"),
    "Income": st.column_config.TextColumn(label="Income (Pula)"),
    "Light": st.column_config.TextColumn(label="⚫ Traffic"),
    "Status": st.column_config.SelectboxColumn(options=STATUS_OPTIONS, label="Status"),
    "Notes": st.column_config.TextColumn(label="Notes"),
    "Renewed On": st.column_config.DateColumn(label="Renewed On", format="YYYY-MM-DD", help="Select the renewal date"),
}

edited = st.data_editor(
    tidy,
    column_config=col_cfg,
    disabled=["CLIENT NUMBER", "CLIENT NAME", "Income", "Light", "Year", "Month"],
    use_container_width=True,
)

# ============================================================
#   SAVE STATUS TO LOCAL JSON
# ============================================================
if st.button("Save Changes"):
    status_store = load_status_store()

    for _, r in edited.iterrows():
        renewed_on_val = r["Renewed On"]
        if isinstance(renewed_on_val, pd.Timestamp):
            renewed_on_val = renewed_on_val.date()
        renewed_on_iso = to_iso_date(renewed_on_val)

        if (r["Status"] == "Renewed") and (renewed_on_iso == ""):
            renewed_on_iso = date.today().isoformat()

        key = f"{r['CLIENT NUMBER']}_{int(r['Year'])}_{r['Month']}"
        status_store[key] = {
            "status": r["Status"],
            "notes": r["Notes"] if r["Notes"] is not None else "",
            "renewed_on": renewed_on_iso,
        }

    try:
        save_status_store(status_store)
        load_events_from_excel.clear()
        st.success("Saved! Changes are stored in status_store.json.")
        st.rerun()
    except Exception as e:
        st.error(f"Failed to save updates: {e}")
