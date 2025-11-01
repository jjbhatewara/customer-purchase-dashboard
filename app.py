import os
import re
from datetime import datetime

import pandas as pd
import streamlit as st

# ---------- CONFIG ----------
st.set_page_config(page_title="Customer Purchase Dashboard", layout="wide", page_icon="📊")
st.markdown("""
    <style>
        body {background-color: #0E1117; color: #FAFAFA;}
        .stApp {background-color: #0E1117;}
        div[data-testid="stSidebar"] {background-color: #1E1E1E;}
        .block-container {padding-top: 1rem; padding-bottom: 1rem;}
        h1, h2, h3, h4, h5 {color: #00B4D8;}
        table {color: #FAFAFA !important;}
    </style>
""", unsafe_allow_html=True)

# ---------- PATHS ----------
DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)
SALES_FILE = os.path.join(DATA_DIR, "sales_data.xlsx")
COMPANY_FILE = os.path.join(DATA_DIR, "company_list.xlsx")

# ---------- HELPERS ----------
def detect_header_row(path_or_buffer):
    """Return row index (0-based) that looks like the header (has keywords)."""
    preview = pd.read_excel(path_or_buffer, header=None, nrows=8)
    for i in range(len(preview)):
        row = preview.iloc[i].astype(str).str.lower()
        rowtext = " ".join(row.tolist())
        if any(k in rowtext for k in ["product", "qty", "ledger", "invoice", "inv no", "invno"]):
            return i
    return 0

def normalize_columns(df):
    """Map likely column name variants to a standard set used by the app."""
    col_map = {}
    for col in df.columns:
        c = str(col).strip().lower()
        # product
        if re.search(r'product', c):
            col_map[col] = "Product Name"
            continue
        # qty / quantity
        if re.search(r'\b(qty|quantity|qty\.)\b', c):
            col_map[col] = "Qty"
            continue
        # free
        if re.search(r'\bfree\b', c):
            col_map[col] = "Free"
            continue
        # rate
        if re.search(r'\brate\b', c):
            col_map[col] = "Rate"
            continue
        # gross amount / grs / value / amount
        if re.search(r'grs|gross|grsamt|amount|value|amount\s*\(|gross\s*amt', c):
            col_map[col] = "Value"
            continue
        # invoice number
        if re.search(r'inv|invoice', c):
            col_map[col] = "InvNo"
            continue
        # ledger / customer
        if re.search(r'ledger|customer|retailer|party|name of customer', c):
            col_map[col] = "Ledger Account"
            continue
        # parent manufacturer
        if re.search(r'parent|parent\s*manufacturer|parentmanufacturer|parent manufact', c):
            col_map[col] = "Parent Manufacturer"
            continue
        # manufacturer / division
        if re.search(r'manufacturer|division', c):
            col_map[col] = "Manufacturer / Division"
            continue
        # city / area
        if re.search(r'\barea\b', c):
            col_map[col] = "Area"
            continue
        if re.search(r'\bcity\b', c):
            col_map[col] = "City"
            continue

    df = df.rename(columns=col_map)
    return df

def read_sales_file(path):
    """Read sales Excel and return cleaned DataFrame and inferred period string."""
    header_row = detect_header_row(path)
    raw = pd.read_excel(path, header=None)
    # try to parse period text from first 4 rows (if present)
    period_text = " ".join(raw.iloc[0:4, :].astype(str).fillna("").apply(lambda r: " ".join(r), axis=1).tolist())
    period = None
    m = re.search(r'from\s*[:\-]?\s*(\d{1,2}[\/\-\.\s]\d{1,2}[\/\-\.\s]\d{2,4}).*upto\s*[:\-]?\s*(\d{1,2}[\/\-\.\s]\d{1,2}[\/\-\.\s]\d{2,4})', period_text, re.IGNORECASE)
    if m:
        period = f"{m.group(1).strip()} → {m.group(2).strip()}"

    # read with detected header row
    df = pd.read_excel(path, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    df = normalize_columns(df)

    # Guarantee columns exist (create if missing)
    for needed in ["Product Name", "Qty", "Value", "InvNo", "Ledger Account", "Parent Manufacturer"]:
        if needed not in df.columns:
            df[needed] = pd.NA

    # numeric coercion
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce").fillna(0)

    # ensure strings
    df["Ledger Account"] = df["Ledger Account"].astype(str).str.strip()
    df["Parent Manufacturer"] = df["Parent Manufacturer"].astype(str).str.strip()

    # if InvNo is numeric-ish, make string (to count unique invoices)
    df["InvNo"] = df["InvNo"].astype(str).str.strip()

    return df, period

def read_company_list(path):
    header_row = detect_header_row(path)
    df = pd.read_excel(path, header=header_row)
    # flatten first two columns if present
    cols = df.columns.tolist()
    companies = []
    for col in cols[:2]:
        companies.extend(df[col].dropna().astype(str).str.strip().tolist())
    # uppercase normalization for matching
    companies = [c.upper() for c in companies if str(c).strip() != ""]
    # unique and sorted
    companies = sorted(list(dict.fromkeys(companies)))
    return companies

def ensure_required_columns_present(df, required):
    missing = [c for c in required if c not in df.columns]
    return missing

# ---------- REPORT FUNCTIONS ----------
def generate_report1(df_cust):
    # If InvNo exists, No of Instances should be nunique invno, else count rows
    if df_cust.empty:
        return pd.DataFrame(columns=["Sr. No", "Product Name", "Parent Manufacturer", "No. of Instances", "Max Purchases", "Avg Purchases"])
    grp = df_cust.groupby(["Product Name", "Parent Manufacturer"]).agg(
        No_of_Instances = ("InvNo", lambda s: s.nunique() if s.notna().any() else s.shape[0]),
        Max_Purchases = ("Qty", "max"),
        Avg_Purchases = ("Qty", "mean")
    ).reset_index()
    grp["Avg_Purchases"] = grp["Avg_Purchases"].round(2)
    grp = grp.sort_values("Avg_Purchases", ascending=False).reset_index(drop=True)
    grp.insert(0, "Sr. No", grp.index + 1)
    # rename columns nicely
    grp = grp.rename(columns={
        "Product Name": "Product Name",
        "Parent Manufacturer": "Parent Manufacturer",
        "No_of_Instances": "No. of Instances",
        "Max_Purchases": "Max Purchases",
        "Avg_Purchases": "Avg Purchases"
    })
    return grp

def generate_report2(df_cust, company_list):
    # compute sums by Parent Manufacturer using customer-filtered df (df_cust)
    summary = (df_cust.groupby("Parent Manufacturer")
               .agg(Total_Qty=("Qty", "sum"), Total_Value=("Value", "sum"))
               .reset_index())
    # uppercase Parent Manufacturer for matching
    summary["Parent Manufacturer"] = summary["Parent Manufacturer"].astype(str).str.upper()
    # build full list dataframe
    full = pd.DataFrame({"Parent Manufacturer": [c.upper() for c in company_list]})
    merged = full.merge(summary, on="Parent Manufacturer", how="left").fillna(0)
    merged["Total_Qty"] = merged["Total_Qty"].astype(int)
    merged["Total_Value"] = merged["Total_Value"].astype(float).round(2)
    merged = merged.sort_values("Total_Value", ascending=False).reset_index(drop=True)
    merged.insert(0, "Sr. No", merged.index + 1)
    # rename for display
    merged = merged.rename(columns={"Parent Manufacturer": "Parent Manufacturer", "Total_Qty": "Total Qty", "Total_Value": "Total Value"})
    return merged

# ---------- UI ----------
st.title("📊 Customer Purchase Dashboard (Persistent Data)")

tab1, tab2 = st.tabs(["🗂️ Upload Files", "📈 Dashboard"])

with tab1:
    st.header("Upload / Replace Files (Persisted for all users)")
    st.markdown("Upload Sales Data and All Company List. Files are saved in the app `data/` folder and will be used by everyone until replaced.")

    uploaded_sales = st.file_uploader("Upload Sales Data (.xlsx)", type=["xlsx"], key="u_sales")
    uploaded_company = st.file_uploader("Upload Company List (.xlsx)", type=["xlsx"], key="u_company")

    if uploaded_sales is not None:
        with open(SALES_FILE, "wb") as f:
            f.write(uploaded_sales.getbuffer())
        st.success("✅ Sales file saved to server data folder.")

    if uploaded_company is not None:
        with open(COMPANY_FILE, "wb") as f:
            f.write(uploaded_company.getbuffer())
        st.success("✅ Company list file saved to server data folder.")

    # Show current status
    if os.path.exists(SALES_FILE) and os.path.exists(COMPANY_FILE):
        t_sales = datetime.fromtimestamp(os.path.getmtime(SALES_FILE)).strftime("%Y-%m-%d %H:%M:%S")
        t_comp = datetime.fromtimestamp(os.path.getmtime(COMPANY_FILE)).strftime("%Y-%m-%d %H:%M:%S")
        st.info(f"Current files present. Sales last updated: {t_sales} | Company list last updated: {t_comp}")
    else:
        st.warning("Please upload both files to enable the Dashboard tab.")

with tab2:
    # Dashboard loads persisted files
    if not (os.path.exists(SALES_FILE) and os.path.exists(COMPANY_FILE)):
        st.warning("No persisted files found. Upload files first in the 'Upload Files' tab.")
        st.stop()

    # Read/prepare data
    try:
        sales_df, period = read_sales_file(SALES_FILE)
    except Exception as e:
        st.error(f"Failed to read sales file: {e}")
        st.stop()

    try:
        company_list = read_company_list(COMPANY_FILE)
    except Exception as e:
        st.error(f"Failed to read company list file: {e}")
        st.stop()

    # Check that essential columns exist and/or have data
    required = ["Ledger Account", "Parent Manufacturer", "Qty", "Value"]
    missing = ensure_required_columns_present(sales_df, required)
    if missing:
        st.error(f"The uploaded sales file is missing required columns after normalization: {missing}. Please check your file headers.")
        st.write("Detected columns:", sales_df.columns.tolist())
        st.stop()

    # Show period (if found)
    if period:
        st.markdown(f"### 📅 Period: {period}")

    # Filters
    st.subheader("Filters")
    customers = sales_df["Ledger Account"].replace("nan", pd.NA).dropna().unique().tolist()
    if not customers:
        st.warning("No customers (Ledger Account) found in the sales data.")
        st.stop()

    selected_customer = st.selectbox("Select Customer (Ledger Account)", sorted(customers))

    # Filter data for customer
    cust_df = sales_df[sales_df["Ledger Account"] == selected_customer]

    # Report 1
    st.subheader("Report 1 — Frequent Purchases (sorted by Avg Purchases ↓)")
    rpt1 = generate_report1(cust_df)
    st.dataframe(rpt1, use_container_width=True, hide_index=True)

    # Report 2
    st.subheader("Report 2 — Company-wise Purchases (includes 0-sales companies, sorted by Total Value ↓)")
    # << FIXED: use cust_df (customer-filtered) instead of whole sales_df >>
    rpt2 = generate_report2(cust_df, company_list)
    st.dataframe(rpt2, use_container_width=True, hide_index=True)

    st.caption("Created by Jinesh | Sanjay Distributors")
