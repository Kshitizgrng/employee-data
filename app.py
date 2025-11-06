import io
import os
from datetime import date
from typing import Dict, List

import numpy as np
import pandas as pd
import streamlit as st
import altair as alt

st.set_page_config(
    page_title="📊 People Analytics – Payroll & Headcount Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📊 People Analytics – Payroll & Headcount Dashboard")
st.caption("Upload or auto-load `01-Employee Database- Original.xlsx`. Headers are auto-normalized and cleaned.")

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

st.sidebar.header("⚙️ Controls")
DEFAULT_FILE = "01-Employee Database- Original.xlsx"
uploaded = None
df_raw = None

if os.path.exists(DEFAULT_FILE):
    try:
        df_raw = pd.read_excel(DEFAULT_FILE, dtype=str)
        st.sidebar.success(f"Auto-loaded `{DEFAULT_FILE}` ✅")
    except Exception as e:
        st.sidebar.error(f"Error reading default file: {e}")
else:
    uploaded = st.sidebar.file_uploader("📤 Upload Excel", type=["xlsx", "xls"])

st.sidebar.markdown("---")
search_q = st.sidebar.text_input("🔎 Search name/email/department", "")
st.sidebar.caption("Tip: partial matches work, e.g. `mar` → Mark, Martha")
hide_names = st.sidebar.toggle("🙈 Anonymize names", value=False)
show_outliers_only = st.sidebar.toggle("🚨 Focus on outliers (±2.5σ)", value=False)

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
df = df_raw.rename(columns={orig: canon for canon, orig in col_map.items()})

for col in REQUIRED_COLUMNS_CANON.keys():
    if col not in df.columns:
        df[col] = np.nan
        st.warning(f"Missing column added blank: '{col}'")

df["date_joined"] = pd.to_datetime(df["date_joined"], errors="coerce")
for col in NUMERIC_COLS:
    df[col] = df[col].apply(to_number)

if df["years_of_service"].isna().all() or (df["years_of_service"].median(skipna=True) <= 0.01):
    df["years_of_service"] = df["date_joined"].apply(calc_years_from_date)

if df["total_bonus"].isna().all() and "bonus_paid" in df.columns:
    df["total_bonus"] = df["bonus_paid"]

df["full_name"] = (df["first_name"].fillna("") + " " + df["last_name"].fillna("")).str.strip()
if hide_names:
    df["full_name"] = "Emp " + pd.util.hash_pandas_object(df["employee_id"].astype(str)).astype(str).str[-6:]
df["total_comp"] = df[["salary", "bonus_paid", "overtime_paid"]].sum(axis=1, min_count=1)
df["join_year"] = df["date_joined"].dt.year
tenure_bins = [-0.1, 1, 3, 5, 10, 20, 100]
tenure_labels = ["<1", "1–3", "3–5", "5–10", "10–20", "20+"]
df["tenure_band"] = pd.cut(df["years_of_service"], bins=tenure_bins, labels=tenure_labels)
df["salary_z"] = zscore(df["salary"])
df["totalcomp_z"] = zscore(df["total_comp"])
if show_outliers_only:
    df = df[(df["salary_z"].abs() >= 2.5) | (df["totalcomp_z"].abs() >= 2.5)]

divisions = sorted(df["division"].dropna().unique())
departments = sorted(df["department"].dropna().unique())
statuses = sorted(df["employment_status"].dropna().unique())
sel_div = st.multiselect("Division", options=divisions, default=divisions)
sel_dept = st.multiselect("Department", options=departments, default=departments)
sel_status = st.multiselect("Employment Status", options=statuses, default=statuses)

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

hc = int(df_f["employee_id"].nunique())
avg_ten = round(df_f["years_of_service"].mean(skipna=True), 2) if hc else 0.0
active_ratio = df_f["employment_status"].str.lower().eq("active").mean() if hc else 0.0
total_payroll = df_f["total_comp"].sum(skipna=True)
med_salary = float(df_f["salary"].median(skipna=True)) if hc else 0.0
p90_salary = float(df_f["salary"].quantile(0.9)) if hc else 0.0
avg_totalcomp = float(df_f["total_comp"].mean(skipna=True)) if hc else 0.0

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Headcount", f"{hc:,}")
k2.metric("Avg Tenure", f"{avg_ten:.2f} yrs")
k3.metric("Active", f"{active_ratio*100:,.1f}%")
k4.metric("Total Payroll", f"${total_payroll:,.0f}")
k5.metric("Median Salary", f"${med_salary:,.0f}")
k6.metric("Avg Total Comp / FTE", f"${avg_totalcomp:,.0f}")

tab_overview, tab_comp, tab_org, tab_quality, tab_data = st.tabs(
    ["📈 Overview", "💸 Compensation & Tenure", "🏢 Org Structure", "🧹 Data Quality", "📑 Data & Export"]
)

with tab_overview:
    sel = alt.selection_point(fields=["division"], bind="legend")
    grp_div = df_f.groupby("division", dropna=False)["employee_id"].nunique().reset_index(name="headcount")
    chart_div = (
        alt.Chart(grp_div)
        .mark_bar()
        .encode(
            x=alt.X("headcount:Q", title="Headcount"),
            y=alt.Y("division:N", sort="-x", title="Division"),
            tooltip=["division", "headcount"],
            color=alt.Color("division:N", legend=alt.Legend(title="Division")),
        )
        .add_params(sel)
        .properties(height=350)
    )
    payroll_dept = df_f.groupby("department", dropna=False)["total_comp"].sum().reset_index()
    chart_payroll = (
        alt.Chart(payroll_dept)
        .mark_bar()
        .encode(
            x=alt.X("total_comp:Q", title="Total Compensation"),
            y=alt.Y("department:N", sort="-x", title="Department"),
            tooltip=["department", alt.Tooltip("total_comp:Q", format="$.2f")],
            color=alt.Color("department:N", legend=None),
        )
        .properties(height=350)
    )
    left, right = st.columns(2)
    with left:
        st.altair_chart(chart_div, use_container_width=True)
    with right:
        st.altair_chart(chart_payroll, use_container_width=True)
    hc_status = (
        df_f.assign(cnt=1)
        .groupby(["employment_status"], dropna=False)["cnt"]
        .sum()
        .reset_index()
    )
    chart_status = (
        alt.Chart(hc_status)
        .mark_arc(innerRadius=60)
        .encode(
            theta="cnt:Q",
            color=alt.Color("employment_status:N", legend=alt.Legend(title="Status")),
            tooltip=["employment_status", "cnt"]
        )
        .properties(height=360)
    )
    st.altair_chart(chart_status, use_container_width=True)
    hires_ts = (
        df_f.dropna(subset=["date_joined"])
        .sort_values("date_joined")
        .assign(headcount=lambda d: np.arange(1, len(d) + 1))
        .rename(columns={"date_joined": "date"})
    )
    chart_hires = (
        alt.Chart(hires_ts)
        .mark_line(point=True)
        .encode(
            x=alt.X("date:T", title="Date Joined"),
            y=alt.Y("headcount:Q", title="Cumulative Hires"),
            tooltip=[alt.Tooltip("date:T", title="Date"), alt.Tooltip("headcount:Q", title="Cumulative Hires")]
        )
        .properties(height=300)
    )
    st.altair_chart(chart_hires, use_container_width=True)

with tab_comp:
    scatter = (
        alt.Chart(df_f)
        .mark_circle()
        .encode(
            x=alt.X("years_of_service:Q", title="Years of Service"),
            y=alt.Y("salary:Q", title="Base Salary"),
            color=alt.Color("division:N", legend=alt.Legend(title="Division")),
            tooltip=["full_name", "division", "department", alt.Tooltip("salary:Q", format="$.0f"), "years_of_service"]
        )
        .interactive()
        .properties(height=380)
    )
    reg = (
        scatter.transform_regression("years_of_service", "salary").mark_line()
        .encode(color=alt.value("#666"))
    )
    st.altair_chart(scatter + reg, use_container_width=True)
    comp_box = (
        alt.Chart(df_f)
        .mark_boxplot(extent="min-max")
        .encode(
            x=alt.X("division:N", title="Division", sort="-y"),
            y=alt.Y("salary:Q", title="Salary"),
            color="division:N",
            tooltip=[alt.Tooltip("salary:Q", format="$.0f")]
        )
        .properties(height=380)
    )
    st.altair_chart(comp_box, use_container_width=True)
    pay_mix = (
        df_f.melt(
            id_vars=["employee_id", "division", "department", "full_name"],
            value_vars=["salary", "bonus_paid", "overtime_paid"],
            var_name="component",
            value_name="amount",
        )
        .dropna(subset=["amount"])
    )
    pay_mix_grp = pay_mix.groupby(["division", "component"], dropna=False)["amount"].sum().reset_index()
    stack = (
        alt.Chart(pay_mix_grp)
        .mark_bar()
        .encode(
            x=alt.X("sum(amount):Q", stack="normalize", title="Share of Total Comp"),
            y=alt.Y("division:N", sort="-x", title="Division"),
            color=alt.Color("component:N", legend=alt.Legend(title="Component")),
            tooltip=[alt.Tooltip("sum(amount):Q", format="$.0f"), "division", "component"]
        )
        .properties(height=380)
    )
    st.altair_chart(stack, use_container_width=True)
    tenure_hist = (
        alt.Chart(df_f.dropna(subset=["years_of_service"]))
        .mark_bar()
        .encode(
            x=alt.X("years_of_service:Q", bin=alt.Bin(maxbins=25), title="Years of Service"),
            y=alt.Y("count():Q", title="Employees"),
            tooltip=["count()"]
        )
        .properties(height=300)
    )
    st.altair_chart(tenure_hist, use_container_width=True)

with tab_org:
    heat_data = df_f.assign(cnt=1).groupby(["division", "department"], dropna=False)["cnt"].sum().reset_index()
    heat = (
        alt.Chart(heat_data)
        .mark_rect()
        .encode(
            x=alt.X("department:N", title="Department"),
            y=alt.Y("division:N", title="Division"),
            color=alt.Color("cnt:Q", title="Headcount"),
            tooltip=["division", "department", "cnt"]
        )
        .properties(height=420)
    )
    st.altair_chart(heat, use_container_width=True)
    band = (
        alt.Chart(df_f.dropna(subset=["tenure_band"]))
        .mark_bar()
        .encode(
            x=alt.X("tenure_band:N", sort=tenure_labels, title="Tenure Band (yrs)"),
            y=alt.Y("count():Q", title="Employees"),
            color=alt.Color("tenure_band:N", legend=None),
            tooltip=["tenure_band", "count()"]
        )
        .properties(height=300)
    )
    st.altair_chart(band, use_container_width=True)
    top_depts = (
        df_f.groupby("department", dropna=False)["total_comp"]
        .sum()
        .reset_index()
        .sort_values("total_comp", ascending=False)
        .head(15)
    )
    pareto = (
        alt.Chart(top_depts.assign(rank=lambda d: np.arange(1, len(d) + 1)))
        .mark_bar()
        .encode(
            x=alt.X("rank:O", title="Department Rank by Payroll"),
            y=alt.Y("total_comp:Q", title="Total Compensation"),
            tooltip=["department", alt.Tooltip("total_comp:Q", format="$.0f")],
            color=alt.Color("department:N", legend=None)
        )
        .properties(height=300)
    )
    st.altair_chart(pareto, use_container_width=True)

with tab_quality:
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
        st.caption("No obvious data issues detected.")

    outliers = df_f[(df_f["salary_z"].abs() >= 2.5) | (df_f["totalcomp_z"].abs() >= 2.5)][
        ["employee_id", "full_name", "division", "department", "salary", "total_comp", "salary_z", "totalcomp_z"]
    ].sort_values("salary_z", key=lambda s: s.abs(), ascending=False)

    st.write("Extreme values (|z| ≥ 2.5):")
    st.dataframe(outliers, use_container_width=True, hide_index=True)

with tab_data:
    st.dataframe(
        df_f[
            ["employee_id", "full_name", "email", "division", "department", "employment_status",
             "date_joined", "join_year", "years_of_service", "tenure_band",
             "salary", "bonus_paid", "overtime_paid", "bonus_rate", "hourly_rate",
             "total_bonus", "total_comp"]
        ].sort_values(["division", "department", "full_name"], na_position="last"),
        use_container_width=True,
        hide_index=True,
    )
    st.markdown("### ⬇️ Export")
    csv_bytes = df_f.to_csv(index=False).encode("utf-8")
    st.download_button("Download CSV", csv_bytes, file_name="filtered.csv", mime="text/csv")
    xlsx_bytes = excel_download(df_f)
    st.download_button(
        "Download Excel",
        xlsx_bytes,
        file_name="filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.caption("All charts and tables reflect the current filters, search, anonymity, and outlier settings.")
