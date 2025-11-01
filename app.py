
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
import re
from datetime import datetime
from dateutil import parser

st.set_page_config(page_title="Customer Purchase Reports Dashboard", layout="wide")

st.title("Customer Purchase Reports Dashboard")

# Utility functions
@st.cache_data
def load_excel(file):
    # Attempt to read data starting from row 4 (skip first 3 rows)
    preview = pd.read_excel(uploaded_file, nrows=5, header=None)

    # Find the row index that contains "Product Name" or "Qty"
    header_row = None
    for i, row in preview.iterrows():
        if row.astype(str).str.contains(r'(?i)product|qty').any():
            header_row = i
            break

    # Now read the full file with the detected header
    if header_row is not None:
        df = pd.read_excel(uploaded_file, header=header_row)
    else:
        df = pd.read_excel(uploaded_file)
    return df

def infer_date_range_from_header(file):
    # Read first few rows as raw text to find patterns like 'From : dd/mm/yy Upto : dd/mm/yy'
    try:
        raw = pd.read_excel(file, nrows=3, header=None).astype(str).fillna('')
        joined = " ".join(raw.stack().tolist())
        m = re.search(r'From\s*:?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})\s*Upto\s*:?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})', joined, re.IGNORECASE)
        if m:
            d1 = parser.parse(m.group(1), dayfirst=True)
            d2 = parser.parse(m.group(2), dayfirst=True)
            return d1, d2
    except Exception:
        pass
    return None, None

def safe_to_datetime(s):
    try:
        return parser.parse(str(s), dayfirst=True)
    except Exception:
        return pd.NaT

def prepare_data(df, uploaded_file):
    # Normalize column names by stripping and replacing newlines
    df.columns = [str(c).strip() for c in df.columns]
    # Common expected column name mapping
    col_map_candidates = {
        'Product Name': ['Product Name','ProductName','PRODUCT','PRODUCT NAME'],
        'Qty': ['Qty','QTY','Quantity','QTY '],
        'Free': ['Free','FREE'],
        'Rate': ['Rate','RATE'],
        'GrsAmt': ['GrsAmt','Grs Amt','Gross Amount','GrsAmount','Grs Amt'],
        'InvNo': ['InvNo','Inv No','Invoice No','InvoiceNumber','Inv #'],
        'Ledger Account': ['Ledger Account','LedgerAccount','Customer','Name of Customer','Ledger'],
        'Area': ['Area','AREA','Town'],
        'City': ['City','CITY'],
        'Manufacturer / Division': ['Manufacturer / Division','Manufacturer','Division'],
        'Parent Manufacturer': ['Parent Manufacturer','ParentManufacture','Parent Manufacturer '],
        'Supplier Name': ['Supplier Name','SupplierName','Supplier']
    }
    # Create a reverse lookup
    col_map = {}
    for standard, candidates in col_map_candidates.items():
        for c in df.columns:
            for cand in candidates:
                if str(c).strip().lower() == cand.strip().lower():
                    col_map[c] = standard

    df = df.rename(columns=col_map)
    # Keep only relevant columns if present
    needed = ['Product Name','Qty','Free','Rate','GrsAmt','InvNo','Ledger Account','Area','City','Manufacturer / Division','Parent Manufacturer','Supplier Name']
    for col in needed:
        if col not in df.columns:
            df[col] = np.nan

    # Convert Qty and GrsAmt to numeric
    df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').fillna(0)
    df['GrsAmt'] = pd.to_numeric(df['GrsAmt'], errors='coerce').fillna(0)
    # Try to detect a date column
    date_cols = [c for c in df.columns if 'date' in str(c).lower() or 'dt' in str(c).lower() or 'invdate' in str(c).lower()]
    if date_cols:
        df['__Date'] = pd.to_datetime(df[date_cols[0]], errors='coerce', dayfirst=True)
    else:
        # Try to infer from header
        d1, d2 = infer_date_range_from_header(uploaded_file)
        if d1 is not None:
            # put the start date as the date for all rows (best-effort)
            df['__Date'] = pd.NaT
            df['__InferredStart'] = d1
            df['__InferredEnd'] = d2
        else:
            df['__Date'] = pd.NaT

    # Create Month column if Date available
    df['Month'] = df['__Date'].dt.strftime('%Y-%m').fillna('Unknown')
    return df

def generate_report1(df_cust):
    # Group by product and parent manufacturer
    grp = df_cust.groupby(['Product Name','Parent Manufacturer']).agg(
        No_of_Instances = ('InvNo','nunique'),
        Max_Purchases = ('Qty','max'),
        Avg_Purchases = ('Qty','mean')
    ).reset_index()
    grp = grp.sort_values(['No_of_Instances','Avg_Purchases'], ascending=False).reset_index(drop=True)
    grp.insert(0, 'Sr. No', grp.index+1)
    return grp

def generate_report2(df_cust, start_month=None, end_month=None, value_col='Qty'):
    # If Month is Unknown, then monthly split won't be possible per row
    if df_cust['Month'].eq('Unknown').all():
        # fallback: single column for period total
        total = df_cust.groupby(['Parent Manufacturer']).agg(Total=(value_col,'sum')).reset_index()
        total.insert(0, 'Sr. No', total.index+1)
        return total, False
    else:
        # Filter by date range if provided (start_month/end_month in 'YYYY-MM' format)
        df_m = df_cust.copy()
        if start_month:
            df_m = df_m[df_m['Month'] >= start_month]
        if end_month:
            df_m = df_m[df_m['Month'] <= end_month]
        pivot = pd.pivot_table(df_m, index=['Parent Manufacturer'], columns='Month', values=value_col, aggfunc='sum', fill_value=0)
        pivot = pivot.reset_index().rename_axis(None, axis=1)
        pivot.insert(0, 'Sr. No', range(1, len(pivot)+1))
        return pivot, True

def download_button_df(df, file_name="report.csv"):
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(label=f"Download {file_name}", data=csv, file_name=file_name, mime='text/csv')

# UI - Upload
uploaded_file = st.file_uploader("Upload your Excel (.xlsx) file (upload once every 15 days)", type=['xlsx'])

if uploaded_file is not None:
    df_raw = load_excel(uploaded_file)
    st.success("✅ Data uploaded successfully (Last updated: {})".format(datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    with st.expander("Preview raw data (first 10 rows)"):
        st.dataframe(df_raw.head(10))

    df = prepare_data(df_raw, uploaded_file)

    # Sidebar filters
    st.sidebar.header("Filters")
    customers = df['Ledger Account'].dropna().unique().tolist()
    if not customers:
        st.warning("No customers found in the uploaded file. Please check column names.")
    selected_customer = st.sidebar.selectbox("Select Customer (Ledger Account)", options=["-- Select --"] + customers)
    if selected_customer and selected_customer != "-- Select --":
        df_cust = df[df['Ledger Account'] == selected_customer].copy()
        st.sidebar.write("Area:", df_cust['Area'].dropna().unique().tolist())
        st.write(f"## Reports for customer: {selected_customer}")

        # Report 1
        st.write("### Report 1 — Frequent Purchases")
        rpt1 = generate_report1(df_cust)
        st.dataframe(rpt1)
        download_button_df(rpt1, f"{selected_customer}_frequent_purchases.csv")

        # Report 2
        st.write("### Report 2 — Company-wise Monthly Purchases")
        # Month selectors: get unique months sorted
        months = sorted(df['Month'].dropna().unique())
        # If months include 'Unknown', exclude for selector but keep for fallback
        month_options = [m for m in months if m != 'Unknown']
        if month_options:
            col1, col2 = st.columns(2)
            with col1:
                start = st.selectbox("From (YYYY-MM)", options=["--"] + month_options, index=0)
            with col2:
                end = st.selectbox("To (YYYY-MM)", options=["--"] + month_options, index=len(month_options)-1 if month_options else 0)
            start_sel = None if start == "--" else start
            end_sel = None if end == "--" else end
        else:
            start_sel = end_sel = None

        pivot, monthly_available = generate_report2(df_cust, start_sel, end_sel, value_col='Qty')
        if not monthly_available:
            st.info("Monthly breakdown unavailable: no per-row date column found. Showing totals for the uploaded period.")
        st.dataframe(pivot)
        download_button_df(pivot, f"{selected_customer}_company_monthly.csv")

        # Optional: a simple bar chart for top products
        if not df_cust.empty:
            st.write("### Top products by total quantity")
            top_products = df_cust.groupby('Product Name').agg(TotalQty=('Qty','sum')).reset_index().sort_values('TotalQty', ascending=False).head(10)
            fig = px.bar(top_products, x='Product Name', y='TotalQty', title='Top 10 Products (by Qty)')
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Please select a customer from the sidebar to view the reports.")
else:
    st.info("Upload a distributor Excel file to begin. Data is expected to start at row 4 (skip first three rows).")
