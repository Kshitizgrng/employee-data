import io
import os
from datetime import date
from typing import Dict, List

import numpy as np
import pandas as pd
import streamlit as st
import altair as alt

# ----------------------- Page Setup -----------------------
st.set_page_config(
    page_title="📊 People Analytics – Payroll & Headcount Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📊 People Analytics – Payroll & Headcount Dashboard")
st.caption("Upload or auto-load your Excel file (`01-Employee Database- Original.xlsx`). "
           "The app automatically cleans header spacing and casing.")

# ----------------------- Helpers --------------------------
REQUIRED_COLUMNS_CANON = {
    "employee_id": ["employee id", "employee  id", "emp id", "employee_id"],
    "first_name": ["first name", "first  name", "first_name", "fname"],
    "last_name": ["last name", "last  name", "last_name", "lname"],
    "division": ["division"],
    "department": ["department", "dept"],
    "date_joined": ["date joined", "date  joined", "start date", "hire date", "date_joined"],
    "employment_status": ["employment status", "employment  status", "status", "employment_status"],
    "years_of_service": ["years of service", "years  of  service", "tenure", "years_of_service"],
    "email": ["email", "work email"],
    "hourly_rate": ["hourly rate", "hourly  rate", "hourly_rate"],
    "bonus_rate": ["bonus rate", "bonus  rate", "bonus_rate"],
    "salary": ["salary", "base pay", "base salary"],
    "bonus_paid": ["bonus paid", "bonus  paid", "bonus_paid"],
    "overtime_paid": ["overtime paid", "overtime  paid", "overtime_paid"],
    "total_bonus": ["total bonus", "total  bonus", "total_bonus"],
}

NUMERIC_COLS = [
    "years_of_service", "hourly_rate", "bonus_rate",
    "salary", "bonus_paid", "overtime_paid", "total_bonus"
]


def normalize(s: str) -> str:
    if s is None:
        return ""
    s = s.strip().lower()
    return " ".join(s.split())


def auto_map_columns(cols: List[str]) -> Dict[str, str]:
    raw_norm = {c: normalize(c) for c in cols}
    mapping = {}
    used = set()
    for canon, aliases in REQUIRED_COLUMNS_CANON.items():
        for raw, norm in raw_norm.items():
            if raw in used:
                continue
            if norm in [normalize(a) for a in aliases]:
                mapping[canon] = raw
                used.add(raw)
                break
    return mapping


def to_number(s):
    if pd.isna(s):
        return np.nan
    if isinstance(s, (int, float, np.integer, np.floating)):
        return float(s)
    try:
        return float(str(s).replace("$", "").replace(",", "").strip())
    except Exception:
        return np.nan


def excel_download(df: pd.DataFrame, filename: str = "export.xlsx") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
        worksheet = writer.sheets["data"]
        for i, col in enumerate(df.columns):
            max_len = max(10, min(50, df[col].astype(str).str.len().max() if not df.empty else 10))
            worksheet.set_column(i, i, max_len + 2)
    return output.getvalue()


def calc_years_from_date(date_joined: pd.Timestamp) -> float:
    if pd.isna(date_joined):
        return np.nan
    today = pd.Timestamp(date.today())
    delta = (today - date_joined).days
    return round(delta / 365.25, 2)


def zscore(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce")
    return (s - s.mean()) / s.std(ddof=0)


# ----------------------- Sidebar --------------------------
st.sidebar.header("⚙️ Controls")

DEFAULT_FILE = "01-Employee Database- Original.xlsx"
uploaded = None
df_raw = None

# Auto-load if file exists in repo
if os.path.exists(DEFAULT_FILE):
    try:
        df_raw = pd.read_excel(DEFAULT_FILE, dtype=str)
        st.sidebar.success(f"Auto-loaded `{DEFAULT_FILE}` from repository ✅")
    except Exception as e:
        st.sidebar.error(f"Error reading default file: {e}")
else:
    uploaded = st.sidebar.file_uploader(
        "📤 Upload your Excel file (01-Employee Database- Original.xlsx)",
        type=["xlsx", "xls"]
    )

st.sidebar.markdown("---")
search_q = st.sidebar.text_input("🔎 Search name/email/department", "")
st.sidebar.caption("Tip: Type partial matches, e.g. `mar` → Mark, Martha, etc.")

# ----------------------- Load / Clean ----------------------
if df_raw is None and uploaded is None:
    st.info("Please upload or include `01-Employee Database- Original.xlsx` to begin.")
    st.stop()

if uploaded is not None:
    try:
        df_raw = pd.read_excel(uploaded, dtype=str)
    except Exception as e:
        st.error(f"Could not read the Excel file: {e}")
        st.stop()

col_map = auto_map_columns(df_raw.columns.tolist())
missing = [c for c in REQUIRED_COLUMNS_CANON.keys() if c not in col_map]
if missing:
    st.error(f"Missing required columns: {', '.join(missing)}")
    st.stop()

df = df_raw.rename(columns={orig: canon for canon, orig in col_map.items()})
df["date_joined"] = pd.to_datetime(df["date_joined"], errors="coerce")

for col in NUMERIC_COLS:
    df[col] = df[col].apply(to_number)

if df["years_of_service"].isna().all() or (df["years_of_service"].median(skipna=True) <= 0.01):
    df["years_of_service"] = df["date_joined"].apply(calc_years_from_date)

df["full_name"] = (df["first_name"].fillna("") + " " + df["last_name"].fillna("")).str.strip()
df["total_comp"] = df[["salary", "bonus_paid", "overtime_paid"]].sum(axis=1, min_count=1)

# ----------------------- Filters --------------------------
divisions = sorted(df["division"].dropna().unique())
departments = sorted(df["department"].dropna().unique())
statuses = sorted(df["employment_status"].dropna().unique())

sel_div = st.multiselect("Division filter", options=divisions, default=divisions)
sel_dept = st.multiselect("Department filter", options=departments, default=departments)
sel_status = st.multiselect("Employment Status filter", options=statuses, default=statuses)

mask = df["division"].isin(sel_div) & df["department"].isin(sel_dept) & df["employment_status"].isin(sel_status)
q = normalize(search_q)
if q:
    mask &= (
        df["full_name"].str.lower().str.contains(q, na=False)
        | df["email"].str.lower().str.contains(q, na=False)
        | df["department"].str.lower().str.contains(q, na=False)
        | df["division"].str.lower().str.contains(q, na=False)
    )

df_f = df.loc[mask].copy()

# ----------------------- KPIs -----------------------------
c1, c2, c3, c4 = st.columns(4)
headcount = int(df_f["employee_id"].nunique())
avg_tenure = round(df_f["years_of_service"].mean(skipna=True), 2) if headcount else 0.0
active_ratio = df_f["employment_status"].str.lower().eq("active").mean() if headcount else 0.0
total_payroll = df_f["total_comp"].sum(skipna=True)

c1.metric("Headcount", f"{headcount:,}")
c2.metric("Avg. Tenure (yrs)", f"{avg_tenure:.2f}")
c3.metric("Active %", f"{active_ratio*100:,.1f}%")
c4.metric("Total Payroll (Salary + Bonus + OT)", f"${total_payroll:,.0f}")

# ----------------------- Charts ---------------------------
st.subheader("📈 Visuals")

left, right = st.columns(2)
with left:
    grp_div = df_f.groupby("division")["employee_id"].nunique().reset_index(name="headcount")
    st.altair_chart(
        alt.Chart(grp_div).mark_bar().encode(
            x="headcount:Q", y=alt.Y("division:N", sort="-x"), tooltip=["division", "headcount"]
        ),
        use_container_width=True,
    )

with right:
    payroll_dept = df_f.groupby("department")["total_comp"].sum().reset_index()
    st.altair_chart(
        alt.Chart(payroll_dept).mark_bar().encode(
            x="total_comp:Q", y=alt.Y("department:N", sort="-x"),
            tooltip=["department", alt.Tooltip("total_comp:Q", format="$.2f")]
        ),
        use_container_width=True,
    )

# ----------------------- Data Quality ---------------------
st.subheader("🕵️ Data Quality")

issues = []
for col in ["employee_id", "first_name", "last_name", "division", "department", "employment_status", "salary"]:
    n_miss = df_f[col].isna().sum()
    if n_miss:
        issues.append(f"Missing `{col}`: {n_miss} rows")

for col in ["salary", "bonus_paid", "overtime_paid", "hourly_rate", "bonus_rate", "total_bonus"]:
    n_neg = (df_f[col] < 0).sum(skipna=True)
    if n_neg:
        issues.append(f"Negative values in `{col}`: {n_neg} rows")

if issues:
    st.warning("Potential issues detected:")
    for i in issues:
        st.write("•", i)
else:
    st.success("No obvious data issues detected.")

# ----------------------- Data Table -----------------------
st.subheader("📑 Employee Data")
st.dataframe(
    df_f[
        ["employee_id", "full_name", "email", "division", "department", "employment_status",
         "date_joined", "years_of_service", "salary", "bonus_paid", "overtime_paid",
         "bonus_rate", "hourly_rate", "total_bonus", "total_comp"]
    ].sort_values(["division", "department", "full_name"], na_position="last"),
    use_container_width=True,
)

# ----------------------- Exports --------------------------
st.subheader("⬇️ Export")
csv_bytes = df_f.to_csv(index=False).encode("utf-8")
st.download_button("Download CSV", csv_bytes, file_name="filtered.csv", mime="text/csv")
xlsx_bytes = excel_download(df_f)
st.download_button("Download Excel", xlsx_bytes, file_name="filtered.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("All charts and tables reflect the current filters and search query.")


