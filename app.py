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
TRACKER_FILE = "Client Renewal & Budget Tracker.xlsx"
TRACKER_SHEET = "Corporates & Parastatals"

STATUS_OPTIONS = ["On going", "Renewed", "Organic growth", "Lost", "Not renewing", "Awaiting POP"]
TREND_OPTIONS = [
    "RATE REDUCTION",
    "SUM INSURED REDUCED",
    "COVER REDUCED",
    "HIGH DEDUCTIBLE",
    "RATE INCREASED",
    "COVERS UPSOLD",
    "LOST",
]
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


def normalize_match_key(val) -> str:
    s = str(val).strip().lower()
    if s in {"", "nan", "none"}:
        return ""
    return " ".join(s.split())


@st.cache_data(ttl=120)
def load_budget_tracker_data() -> pd.DataFrame:
    tracker = pd.read_excel(
        TRACKER_FILE,
        sheet_name=TRACKER_SHEET,
        header=9,
        usecols="A:L",
        engine="openpyxl",
    )
    tracker.columns = [str(c).strip() for c in tracker.columns]

    needed = [
        "CLIENT #",
        "CLIENT NAME",
        "Renewed Amount",
        "Budget & Renewed Amount Variance",
        "Comments on Variance",
        "Trends on Variance",
    ]
    missing = [c for c in needed if c not in tracker.columns]
    if missing:
        raise ValueError(
            f"Missing required column(s) in {TRACKER_FILE}/{TRACKER_SHEET}: {', '.join(missing)}"
        )

    out = tracker[needed].copy()
    out = out.dropna(subset=["CLIENT #", "CLIENT NAME"], how="all")
    out["CLIENT #"] = out["CLIENT #"].astype(str).str.strip()
    out["CLIENT NAME"] = out["CLIENT NAME"].astype(str).str.strip()
    out["client_number_key"] = out["CLIENT #"].apply(normalize_match_key)
    out["client_name_key"] = out["CLIENT NAME"].apply(normalize_match_key)

    out["Renewed Amount"] = pd.to_numeric(out["Renewed Amount"], errors="coerce")
    out["Budget & Renewed Amount Variance"] = pd.to_numeric(
        out["Budget & Renewed Amount Variance"], errors="coerce"
    )
    out["Comments on Variance"] = out["Comments on Variance"].fillna("").astype(str)
    out["Trends on Variance"] = out["Trends on Variance"].fillna("").astype(str)
    return out


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
                    "StoreKey": key,
                    "CLIENT NUMBER": client_number,
                    "CLIENT NAME": client_name,
                    "Year": this_year,
                    "Month": month,
                    "Amount": saved.get("amount", float(amt)),
                    "Status": saved.get("status", "On going"),
                                    }
            )

    events = pd.DataFrame(rows)
    events["client_number_key"] = events["CLIENT NUMBER"].apply(normalize_match_key)
    events["client_name_key"] = events["CLIENT NAME"].apply(normalize_match_key)

    tracker = load_budget_tracker_data()
    tracker_cols = [
        "Renewed Amount",
        "Budget & Renewed Amount Variance",
        "Comments on Variance",
        "Trends on Variance",
    ]

    by_number = (
        tracker[["client_number_key"] + tracker_cols]
        .dropna(subset=["client_number_key"])
        .drop_duplicates("client_number_key", keep="first")
    )
    merged = events.merge(by_number, on="client_number_key", how="left")

    by_name = (
        tracker[["client_name_key"] + tracker_cols]
        .dropna(subset=["client_name_key"])
        .drop_duplicates("client_name_key", keep="first")
    )
    need_name_fill = merged["Renewed Amount"].isna()
    if need_name_fill.any():
        fill_from_name = merged.loc[need_name_fill, ["client_name_key"]].merge(
            by_name, on="client_name_key", how="left"
        )
        for col in tracker_cols:
            merged.loc[need_name_fill, col] = fill_from_name[col].values

    # Apply manual overrides stored from previous edits.
    for idx, row in merged.iterrows():
        saved = status_store.get(row["StoreKey"], {})
        if "renewed_amount" in saved:
            merged.at[idx, "Renewed Amount"] = saved.get("renewed_amount")
        if "comments_on_variance" in saved:
            merged.at[idx, "Comments on Variance"] = saved.get("comments_on_variance", "")
        if "trends_on_variance" in saved:
            merged.at[idx, "Trends on Variance"] = saved.get("trends_on_variance", "")

    merged["Comments on Variance"] = merged["Comments on Variance"].fillna("")
    merged["Trends on Variance"] = merged["Trends on Variance"].fillna("")

    return merged.drop(columns=["client_number_key", "client_name_key", "StoreKey"])


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
events["Renewed Amount"] = pd.to_numeric(events["Renewed Amount"], errors="coerce")
events["Budget & Renewed Amount Variance"] = pd.to_numeric(
    events["Budget & Renewed Amount Variance"], errors="coerce"
)

# Auto-calculate variance when both amounts are available.
calc_mask = events["Amount"].notna() & events["Renewed Amount"].notna()
events.loc[calc_mask, "Budget & Renewed Amount Variance"] = (
    events.loc[calc_mask, "Renewed Amount"] - events.loc[calc_mask, "Amount"]
)

events["Year"] = pd.to_numeric(events["Year"], errors="coerce").astype("Int64")
events["Budget Amount"] = events["Amount"]
events["Renewed Amount (P)"] = events["Renewed Amount"]
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
c1, c2, c3, c4, c5 = st.columns(5)

total_clients_view = view["CLIENT NUMBER"].nunique()
c1.metric("👥 Total Clients", f"{total_clients_view:,}")

ongoing_view = (view["Status"] == "On going").sum()
c2.metric("📌 Ongoing", f"{ongoing_view:,}")

renewed_view = (view["Status"] == "Renewed").sum()
c3.metric("✅ Renewed", f"{renewed_view:,}")

total_budget_view = pd.to_numeric(view["Amount"], errors="coerce").sum()
c4.metric("💰 Total Budgeted Amount", format_pula(total_budget_view))

total_renewed_amount_view = pd.to_numeric(view["Renewed Amount"], errors="coerce").sum()
c5.metric("💳 Total Renewed Amount", format_pula(total_renewed_amount_view))

# ============================================================
#   EDITABLE TABLE
# ============================================================
cols = [
    "CLIENT NUMBER",
    "CLIENT NAME",
    "Year",
    "Month",
    "Light",
    "Status",
    "Budget Amount",
    "Renewed Amount (P)",
    "Budget & Renewed Amount Variance",
    "Comments on Variance",
    "Trends on Variance",
]
tidy = view[cols].copy()

col_cfg = {
    "CLIENT NUMBER": st.column_config.TextColumn(label="Client Number"),
    "CLIENT NAME": st.column_config.TextColumn(label="Client Name"),
    "Budget Amount": st.column_config.NumberColumn(label="Budget Amount", format="P %.2f"),
    "Renewed Amount (P)": st.column_config.NumberColumn(label="Renewed Amount (P)", format="P %.2f"),
    "Budget & Renewed Amount Variance": st.column_config.NumberColumn(
        label="Budget & Renewed Amount Variance", format="P %.2f"
    ),
    "Comments on Variance": st.column_config.TextColumn(label="Comments on Variance"),
    "Trends on Variance": st.column_config.SelectboxColumn(
        label="Trends on Variance", options=TREND_OPTIONS
    ),
    "Light": st.column_config.TextColumn(label="⚫ Traffic"),
    "Status": st.column_config.SelectboxColumn(options=STATUS_OPTIONS, label="Status"),
}

edited = st.data_editor(
    tidy,
    column_config=col_cfg,
    disabled=[
        "CLIENT NUMBER",
        "CLIENT NAME",
        "Budget & Renewed Amount Variance",
        "Light",
        "Year",
        "Month",
    ],
    use_container_width=True,
)

# ============================================================
#   SAVE STATUS TO LOCAL JSON
# ============================================================
if st.button("Save Changes"):
    status_store = load_status_store()

    for _, r in edited.iterrows():
        budget_amount_val = pd.to_numeric(r["Budget Amount"], errors="coerce")
        renewed_amount_val = pd.to_numeric(r["Renewed Amount (P)"], errors="coerce")

        key = f"{r['CLIENT NUMBER']}_{int(r['Year'])}_{r['Month']}"
        status_store[key] = {
            "status": r["Status"],
            "amount": None if pd.isna(budget_amount_val) else float(budget_amount_val),
            "renewed_amount": None if pd.isna(renewed_amount_val) else float(renewed_amount_val),
            "comments_on_variance": r["Comments on Variance"] if r["Comments on Variance"] is not None else "",
            "trends_on_variance": r["Trends on Variance"] if r["Trends on Variance"] is not None else "",
        }

    try:
        save_status_store(status_store)
        load_events_from_excel.clear()
        st.success("Saved! Changes are stored in status_store.json.")
        st.rerun()
    except Exception as e:
        st.error(f"Failed to save updates: {e}")


