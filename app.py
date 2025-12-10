import os
import re
import sqlite3
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

# ============== PAGE CONFIG + THEME ==============
st.set_page_config(
    page_title="Customer Purchase Dashboard",
    layout="wide",
    page_icon="📊"
)

st.markdown(
    """
    <style>
        body {background-color: #0E1117; color: #FAFAFA;}
        .stApp {background-color: #0E1117;}
        div[data-testid="stSidebar"] {background-color: #1E1E1E;}
        .block-container {padding-top: 1rem; padding-bottom: 1rem;}
        h1, h2, h3, h4, h5 {color: #00B4D8;}
        table {color: #FAFAFA !important;}
    </style>
    """,
    unsafe_allow_html=True
)

# ============== FULL PAGE LOADER OVERLAY ==============
loader_html = """
<style>
#overlay-loader {
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(14, 17, 23, 0.85);
    backdrop-filter: blur(2px);
    z-index: 999999;
    display: none;
}
.loader-spinner {
    border: 10px solid #2c2c2c;
    border-top: 10px solid #00B4D8;
    border-radius: 50%;
    width: 120px;
    height: 120px;
    animation: spin 0.9s linear infinite;
    position: absolute;
    top: 45%;
    left: 45%;
}
@keyframes spin {
    0% { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
}
</style>

<div id="overlay-loader">
    <div class="loader-spinner"></div>
</div>

<script>
function showLoader() {
    const el = document.getElementById("overlay-loader");
    if (el) { el.style.display = "block"; }
}
function hideLoader() {
    const el = document.getElementById("overlay-loader");
    if (el) { el.style.display = "none"; }
}
</script>
"""
st.markdown(loader_html, unsafe_allow_html=True)

# ============== PATHS + DB INIT ==============
DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, "dashboard.db")


@st.cache_resource
def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        create table if not exists sales (
            id integer primary key autoincrement,
            product_name text,
            qty real,
            free_qty real,
            rate real,
            value real,
            inv_no text,
            inv_date text,
            ledger_account text,
            area text,
            city text,
            mobile text,
            parent_manufacturer text,
            manufacturer_division text,
            supplier_name text
        )
        """
    )
    cur.execute(
        """
        create table if not exists companies (
            id integer primary key autoincrement,
            parent_manufacturer text unique
        )
        """
    )
    conn.commit()


init_db()

# ============== EXCEL PARSING HELPERS ==============


def detect_header_row(file):
    if hasattr(file, "seek"):
        file.seek(0)
    preview = pd.read_excel(file, header=None, nrows=8)
    for i in range(len(preview)):
        row = preview.iloc[i].astype(str).str.lower()
        rowtext = " ".join(row.tolist())
        if any(
            k in rowtext
            for k in ["product", "qty", "ledger", "invdate", "inv date", "invno", "invoice", "inv no"]
        ):
            if hasattr(file, "seek"):
                file.seek(0)
            return i
    if hasattr(file, "seek"):
        file.seek(0)
    return 0


def normalize_columns(df):
    col_map = {}
    for col in df.columns:
        c = str(col).strip().lower()

        if "product" in c:
            col_map[col] = "Product Name"
            continue
        if re.search(r"\b(qty|quantity)\b", c):
            col_map[col] = "Qty"
            continue
        if "free" in c:
            col_map[col] = "Free"
            continue
        if "rate" in c:
            col_map[col] = "Rate"
            continue
        if re.search(r"grs|gross|grsamt|amount|value", c):
            col_map[col] = "Value"
            continue
        if "invdate" in c or "inv date" in c or "invoice date" in c:
            col_map[col] = "InvDate"
            continue
        if ("inv" in c or "invoice" in c) and "date" not in c:
            col_map[col] = "InvNo"
            continue
        if "ledger" in c or "customer" in c or "party" in c:
            col_map[col] = "Ledger Account"
            continue
        if "area" in c:
            col_map[col] = "Area"
            continue
        if "city" in c:
            col_map[col] = "City"
            continue
        if "mobile" in c or "phone" in c or "contact" in c:
            col_map[col] = "Mobile"
            continue
        if "parent" in c and "manufact" in c:
            col_map[col] = "Parent Manufacturer"
            continue
        if "manufacturer / division" in c or ("manufacturer" in c and "division" in c):
            col_map[col] = "Manufacturer / Division"
            continue
        if "supplier" in c:
            col_map[col] = "Supplier Name"
            continue

    return df.rename(columns=col_map)


def parse_sales_excel(uploaded_file):
    header_row = detect_header_row(uploaded_file)
    df = pd.read_excel(uploaded_file, header=header_row)

    df.columns = [str(c).strip() for c in df.columns]
    df = normalize_columns(df)

    required = [
        "Product Name",
        "Qty",
        "Free",
        "Rate",
        "Value",
        "InvNo",
        "InvDate",
        "Ledger Account",
        "Area",
        "City",
        "Mobile",
        "Parent Manufacturer",
        "Manufacturer / Division",
        "Supplier Name",
    ]
    for c in required:
        if c not in df.columns:
            df[c] = pd.NA

    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
    df["Free"] = pd.to_numeric(df["Free"], errors="coerce").fillna(0)
    df["Rate"] = pd.to_numeric(df["Rate"], errors="coerce").fillna(0)
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce").fillna(0)

    df["InvDate"] = pd.to_datetime(df["InvDate"], errors="coerce", dayfirst=True)

    for c in required:
        if c not in ["Qty", "Free", "Rate", "Value", "InvDate"]:
            df[c] = df[c].astype(str).str.strip()

    return df


def insert_sales_df(df):
    if df.empty:
        return
    conn = get_conn()
    cur = conn.cursor()

    rows = []
    for _, row in df.iterrows():
        rows.append(
            (
                row["Product Name"],
                float(row["Qty"]),
                float(row["Free"]),
                float(row["Rate"]),
                float(row["Value"]),
                row["InvNo"],
                row["InvDate"].strftime("%Y-%m-%d") if pd.notna(row["InvDate"]) else None,
                row["Ledger Account"],
                row["Area"],
                row["City"],
                row["Mobile"],
                row["Parent Manufacturer"],
                row["Manufacturer / Division"],
                row["Supplier Name"],
            )
        )

    cur.executemany(
        """
        insert into sales (
            product_name,
            qty,
            free_qty,
            rate,
            value,
            inv_no,
            inv_date,
            ledger_account,
            area,
            city,
            mobile,
            parent_manufacturer,
            manufacturer_division,
            supplier_name
        ) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        rows,
    )
    conn.commit()


def parse_company_excel(uploaded_file):
    header_row = detect_header_row(uploaded_file)
    df = pd.read_excel(uploaded_file, header=header_row)
    companies = []
    for col in df.columns[:2]:
        companies.extend(df[col].dropna().astype(str).str.strip().tolist())
    companies = [c.upper() for c in companies if c]
    companies = sorted(list(dict.fromkeys(companies)))
    return companies


def replace_companies(companies):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("delete from companies")
    cur.executemany("insert into companies(parent_manufacturer) values (?)", [(c,) for c in companies])
    conn.commit()

# ============== LOAD DATA FROM DB ==============


def load_sales():
    conn = get_conn()
    df = pd.read_sql_query("select * from sales", conn)
    if df.empty:
        return df
    df["InvDate"] = pd.to_datetime(df["inv_date"], errors="coerce")
    df["Month"] = df["InvDate"].dt.to_period("M").astype(str)
    df["qty_eff"] = df["qty"].fillna(0) + df["free_qty"].fillna(0)
    return df


def load_companies():
    conn = get_conn()
    df = pd.read_sql_query("select parent_manufacturer from companies", conn)
    return df["parent_manufacturer"].tolist()

# ============== REPORTS ==============


def report1(df):
    if df.empty:
        return pd.DataFrame(
            columns=[
                "Sr. No",
                "Product Name",
                "Parent Manufacturer",
                "No. of Instances",
                "Max Purchases",
                "Avg Purchases",
            ]
        )

    g = df.groupby(["product_name", "parent_manufacturer"]).agg(
        No_of_Instances=("inv_no", lambda x: x.nunique()),
        Max_Purchases=("qty_eff", "max"),
        Avg_Purchases=("qty_eff", "mean"),
    ).reset_index()

    g["Avg_Purchases"] = g["Avg_Purchases"].round(2)
    g = g.sort_values("Avg_Purchases", ascending=False).reset_index(drop=True)
    g.insert(0, "Sr. No", g.index + 1)

    g = g.rename(
        columns={
            "product_name": "Product Name",
            "parent_manufacturer": "Parent Manufacturer",
        }
    )

    return g


def report2(df, companies):
    base = pd.DataFrame({"Parent Manufacturer": companies})

    if df.empty:
        base["Total Qty"] = 0
        base["Total Value"] = 0.0
        base.insert(0, "Sr. No", base.index + 1)
        return base

    g = df.groupby("parent_manufacturer").agg(
        Total_Qty=("qty_eff", "sum"),
        Total_Value=("value", "sum"),
    ).reset_index()

    g["parent_manufacturer"] = g["parent_manufacturer"].astype(str).str.upper()

    merged = base.merge(
        g.rename(columns={"parent_manufacturer": "Parent Manufacturer"}),
        on="Parent Manufacturer",
        how="left",
    ).fillna(0)

    merged["Total_Qty"] = merged["Total_Qty"].astype(float)
    merged["Total_Value"] = merged["Total_Value"].astype(float)

    merged = merged.rename(
        columns={
            "Total_Qty": "Total Qty",
            "Total_Value": "Total Value",
        }
    )

    if "Total Value" in merged.columns:
        merged = merged.sort_values("Total Value", ascending=False).reset_index(drop=True)
    merged.insert(0, "Sr. No", merged.index + 1)

    return merged

# ============== UI ==============

st.title("Customer Purchase Dashboard")

tab_dash, tab_upload = st.tabs(["Dashboard", "Upload Data"])

# ---------- TAB: DASHBOARD ----------
with tab_dash:
    st.markdown("<script>showLoader()</script>", unsafe_allow_html=True)
    df = load_sales()
    companies = load_companies()
    st.markdown("<script>hideLoader()</script>", unsafe_allow_html=True)

    if df.empty:
        st.warning("No sales data available. Upload sales files in the Upload Data tab.")
        
    else:
        st.subheader("Filters")

        min_d = df["InvDate"].min()
        max_d = df["InvDate"].max()

        if pd.isna(min_d) or pd.isna(max_d):
            st.warning("InvDate missing in data. Date filter disabled.")
            date_filtered_df = df.copy()
        else:
            today = datetime.today().date()
            one_month_ago = today - timedelta(days=30)

            # Raw desired range
            raw_from = one_month_ago
            raw_to = today

            # Clamp inside data range
            default_from = max(min_d.date(), raw_from)
            default_from = min(default_from, max_d.date())

            default_to = max(min_d.date(), raw_to)
            default_to = min(default_to, max_d.date())

            # Safety: if from > to, fall back to full data range
            if default_from > default_to:
                default_from = min_d.date()
                default_to = max_d.date()

            col1, col2 = st.columns(2)
            with col1:
                from_date = st.date_input(
                    "From Date",
                    value=default_from,
                    min_value=min_d.date(),
                    max_value=max_d.date(),
                )
            with col2:
                to_date = st.date_input(
                    "To Date",
                    value=default_to,
                    min_value=min_d.date(),
                    max_value=max_d.date(),
                )

            st.markdown("<script>showLoader()</script>", unsafe_allow_html=True)
            mask = (df["InvDate"].dt.date >= from_date) & (df["InvDate"].dt.date <= to_date)
            date_filtered_df = df[mask]
            st.markdown("<script>hideLoader()</script>", unsafe_allow_html=True)

        if date_filtered_df.empty:
            st.warning("No data in this date range.")
            st.stop()

        for c in ["ledger_account", "area", "city", "mobile"]:
            if c not in date_filtered_df.columns:
                date_filtered_df[c] = ""
            date_filtered_df[c] = date_filtered_df[c].fillna("").astype(str)

        date_filtered_df["Customer_Display"] = (
            date_filtered_df["ledger_account"]
            + " | "
            + date_filtered_df["area"]
            + " | "
            + date_filtered_df["city"]
            + " | "
            + date_filtered_df["mobile"]
        )
        date_filtered_df["Customer_ID"] = (
            date_filtered_df["ledger_account"]
            + "|"
            + date_filtered_df["area"]
            + "|"
            + date_filtered_df["city"]
            + "|"
            + date_filtered_df["mobile"]
        )

        cust_map = (
            date_filtered_df[["Customer_ID", "Customer_Display"]]
            .drop_duplicates()
            .sort_values("Customer_Display")
        )

        if cust_map.empty:
            st.warning("No customers in this filtered data.")
            st.stop()

        selected_display = st.selectbox("Select Customer", cust_map["Customer_Display"])

        st.markdown("<script>showLoader()</script>", unsafe_allow_html=True)
        selected_id = cust_map.loc[
            cust_map["Customer_Display"] == selected_display, "Customer_ID"
        ].values[0]
        ledger, area, city, mobile = selected_id.split("|")

        df_cust = date_filtered_df[
            (date_filtered_df["ledger_account"] == ledger)
            & (date_filtered_df["area"] == area)
            & (date_filtered_df["city"] == city)
            & (date_filtered_df["mobile"] == mobile)
        ]
        st.markdown("<script>hideLoader()</script>", unsafe_allow_html=True)

        st.subheader("Report 1: Frequent Purchases")
        st.markdown("<script>showLoader()</script>", unsafe_allow_html=True)
        r1 = report1(df_cust)
        st.markdown("<script>hideLoader()</script>", unsafe_allow_html=True)
        st.dataframe(r1, use_container_width=True, hide_index=True)

        st.subheader("Report 2: Company wise Purchases")
        st.markdown("<script>showLoader()</script>", unsafe_allow_html=True)
        if companies:
            r2 = report2(df_cust, companies)
        else:
            r2 = pd.DataFrame(columns=["Sr. No", "Parent Manufacturer", "Total Qty", "Total Value"])
        st.markdown("<script>hideLoader()</script>", unsafe_allow_html=True)
        st.dataframe(r2, use_container_width=True, hide_index=True)

    st.caption("Created by Jinesh | Sanjay Distributors")

# ---------- TAB: UPLOAD DATA ----------
with tab_upload:
    st.header("Upload Sales Files (data is appended to database)")
    files = st.file_uploader("Upload Excel Files", type=["xlsx"], accept_multiple_files=True)

    if files:
        st.markdown("<script>showLoader()</script>", unsafe_allow_html=True)
        for f in files:
            df_new = parse_sales_excel(f)
            insert_sales_df(df_new)
        st.markdown("<script>hideLoader()</script>", unsafe_allow_html=True)
        st.success("All sales files processed and added to database.")

    st.divider()
    st.header("Upload Company List")
    comp_file = st.file_uploader("Upload Company List", type=["xlsx"])

    if comp_file:
        st.markdown("<script>showLoader()</script>", unsafe_allow_html=True)
        comps = parse_company_excel(comp_file)
        replace_companies(comps)
        st.markdown("<script>hideLoader()</script>", unsafe_allow_html=True)
        st.success(f"{len(comps)} companies saved to database.")

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("select count(*) from sales")
    total_sales = cur.fetchone()[0]
    cur.execute("select count(*) from companies")
    total_comp = cur.fetchone()[0]

    if total_sales == 0 and total_comp == 0:
        st.warning("No files uploaded yet. Please upload sales data and the company list to start.")
    elif total_sales == 0:
        st.warning("No sales data uploaded yet. Please upload one or more sales Excel files.")
    elif total_comp == 0:
        st.warning("Sales data is present, but no company list uploaded yet. Please upload the company list.")
    else:
        st.info(f"Database status: {total_sales} sales rows, {total_comp} companies.")