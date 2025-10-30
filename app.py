import io
import math
from datetime import datetime, date
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
st.caption("Upload your Excel file (`01-Employee Database- Original.xlsx`) with the columns described. "
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
    s = " ".join(s.split())
    return s

def auto_map_columns(cols: List[str]) -> Dict[str, str]:
    raw_norm = {c: normalize(c) for c in cols}
    mapping = {}
    used = set()

    for canon, aliases in REQUIRED_COLUMNS_CANON.items():
        matched = None
        for raw, norm in raw_norm.items():
            if raw in used:
                continue
            if norm in [normalize(a) for a in aliases]:
                matched = raw
                break
        if matched is not None:
            mapping[canon] = matched
            used.add(matched)
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

template_df = pd.DataFrame(
    {k: [] for k in [
        "Employee  ID","First  Name","Last  Name","Division","Department","Date  Joined",
        "Employment Status","Years of  Service","Email","Hourly  Rate","Bonus Rate","Salary",
        "Bonus Paid","Overtime  Paid","total bonus"
    ]}
)
template_bytes = excel_download(template_df)
st.sidebar.download_button("Download empty template (.xlsx)", template_bytes, file_name="template.xlsx")

uploaded = st.sidebar.file_uploader(
    "📤 Upload your Excel file (01-Employee Database- Original.xlsx)",
    type=["xlsx", "xls"]
)

st.sidebar.markdown("---")
search_q = st.sidebar.text_input("🔎 Search name/email/department", "")
st.sidebar.caption("Tip: Type partial matches, e.g. `mar` will match *Mark*, *Martha*, etc.")

# ----------------------- Load / Clean ----------------------
if uploaded is None:
    st.info("Please upload `01-Employee Database- Original.xlsx` to begin. "
            "You can rename your file, but keep the same column layout.")
    st.stop()

try:
    df_raw = pd.read_excel(uploaded, dtype=str)
except Exception as e:
    st.error(f"Could not read the Excel file: {e}")
    st.stop()

col_map = auto_map_columns(df_raw.columns.tolist())

missing = [c for c in REQUIRED_COLUMNS_CANON.keys() if c not in col_map]
if missing:
    with st.expander("❗ Column mapping issues (click for details)"):
        st.write("Missing required columns:")
        st.write(", ".join(missing))
        st.write("Original columns detected:", list(df_raw.columns))
    st.error("Fix column names (or download the template) and re-upload.")
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
divisions = sorted([d for d in df["division"].dropna().unique() if str(d).strip() != ""])
departments = sorted([d for d in df["department"].dropna().unique() if str(d).strip() != ""])
statuses = sorted([s for s in df["employment_status"].dropna().unique() if str(s).strip() != ""])

sel_div = st.multiselect("Division filter", options=divisions, default=divisions)
sel_dept = st.multiselect("Department filter", options=departments, default=departments)
sel_status = st.multiselect("Employment Status filter", options=statuses, default=statuses)

q = normalize(search_q)
mask = df["division"].isin(sel_div) & df["department"].isin(sel_dept) & df["employment_status"].isin(sel_status)
if q:
    mask = mask & (
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
    grp_div = df_f.groupby("division", dropna=False)["employee_id"].nunique().reset_index(name="headcount")
    chart_div = alt.Chart(grp_div).mark_bar().encode(
        x=alt.X("headcount:Q", title="Headcount"),
        y=alt.Y("division:N", sort="-x", title="Division"),
        tooltip=["division", "headcount"]
    )
    st.altair_chart(chart_div, use_container_width=True)

with right:
    payroll_dept = df_f.groupby("department", dropna=False)["total_comp"].sum().reset_index()
    chart_payroll = alt.Chart(payroll_dept).mark_bar().encode(
        x=alt.X("total_comp:Q", title="Total Payroll"),
        y=alt.Y("department:N", sort="-x", title="Department"),
        tooltip=["department", alt.Tooltip("total_comp:Q", format="$.2f")]
    )
    st.altair_chart(chart_payroll, use_container_width=True)

left2, right2 = st.columns(2)

with left2:
    tenure_hist = alt.Chart(df_f.dropna(subset=["years_of_service"])).mark_bar().encode(
        x=alt.X("years_of_service:Q", bin=alt.Bin(maxbins=20), title="Years of Service"),
        y=alt.Y("count()", title="Employees"),
        tooltip=[alt.Tooltip("count()", title="Employees")]
    )
    st.altair_chart(tenure_hist, use_container_width=True)

with right2:
    status_counts = df_f["employment_status"].fillna("Unknown").value_counts().reset_index()
    status_counts.columns = ["employment_status", "count"]
    donut = alt.Chart(status_counts).mark_arc(innerRadius=60).encode(
        theta="count:Q",
        color=alt.Color("employment_status:N", legend=alt.Legend(title="Status")),
        tooltip=["employment_status", "count"]
    )
    st.altair_chart(donut, use_container_width=True)

# ----------------------- Data Quality ---------------------
st.subheader("🕵️ Data Quality")

issues = []
for col in ["employee_id", "first_name", "last_name", "division", "department", "employment_status", "salary"]:
    n_miss = df_f[col].isna().sum()
    if n_miss > 0:
        issues.append(f"Missing `{col}`: {n_miss} rows")

for col in ["salary", "bonus_paid", "overtime_paid", "hourly_rate", "bonus_rate", "total_bonus"]:
    n_neg = (df_f[col] < 0).sum(skipna=True)
    if n_neg > 0:
        issues.append(f"Negative values in `{col}`: {int(n_neg)} rows")

if not df_f.empty:
    df_out = df_f.copy()
    df_out["salary_z"] = df_out.groupby("division")["salary"].transform(zscore)
    flagged = df_out[abs(df_out["salary_z"]) >= 3].sort_values("salary_z", ascending=False)
    if not flagged.empty:
        with st.expander("⚠️ Salary outliers (|z| ≥ 3)"):
            st.dataframe(flagged[["employee_id", "full_name", "division", "department", "salary", "salary_z"]], use_container_width=True)
    else:
        st.caption("No extreme salary outliers detected.")

if issues:
    st.warning("Potential issues detected:")
    for i in issues:
        st.write("- " + i)
else:
    st.success("No obvious data quality issues detected.")

# ----------------------- Pivots & Table -------------------
st.subheader("📑 Pivots & Table")

pivot_type = st.selectbox("Pivot preset", ["Headcount by Division/Department", "Payroll by Division/Department", "None"])

if pivot_type == "Headcount by Division/Department":
    pivot = pd.pivot_table(df_f, index="division", columns="department",
                           values="employee_id", aggfunc=pd.Series.nunique, fill_value=0)
    st.dataframe(pivot, use_container_width=True)
    st.download_button("Download pivot (CSV)", pivot.to_csv().encode("utf-8"), file_name="pivot_headcount.csv", mime="text/csv")

elif pivot_type == "Payroll by Division/Department":
    pivot = pd.pivot_table(df_f, index="division", columns="department",
                           values="total_comp", aggfunc="sum", fill_value=0)
    st.dataframe(pivot.style.format("${:,.0f}"), use_container_width=True)
    st.download_button("Download pivot (CSV)", pivot.to_csv().encode("utf-8"), file_name="pivot_payroll.csv", mime="text/csv")

st.dataframe(
    df_f[
        ["employee_id","full_name","email","division","department","employment_status",
         "date_joined","years_of_service","salary","bonus_paid","overtime_paid",
         "bonus_rate","hourly_rate","total_bonus","total_comp"]
    ].sort_values(["division","department","full_name"], na_position="last"),
    use_container_width=True,
)

# ----------------------- Exports --------------------------
st.subheader("⬇️ Export")
colx, coly = st.columns(2)

csv_bytes = df_f.to_csv(index=False).encode("utf-8")
colx.download_button("Download filtered data (CSV)", csv_bytes, file_name="filtered.csv", mime="text/csv")

xlsx_bytes = excel_download(df_f)
coly.download_button("Download filtered data (Excel)", xlsx_bytes, file_name="filtered.xlsx",
                     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("All charts/tables reflect your current filters and search.")

