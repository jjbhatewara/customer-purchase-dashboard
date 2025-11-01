import streamlit as st
import pandas as pd
import numpy as np
import re

# -------------------------
# Streamlit Page Config
# -------------------------
st.set_page_config(
    page_title="Customer Purchase Dashboard",
    layout="wide",
    page_icon="📊"
)

# -------------------------
# Global Styling (Dark UI)
# -------------------------
st.markdown("""
    <style>
        .reportview-container, .stApp {
            background-color: #0E1117;
            color: #FAFAFA;
        }
        h1, h2, h3, h4 {
            color: #00B4D8;
        }
        div[data-testid="stExpander"] div[role="button"] {
            color: #00B4D8;
        }
        .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
        }
        .stTabs [data-baseweb="tab"] {
            color: #FAFAFA;
            background-color: #1E1E1E;
        }
        .stTabs [aria-selected="true"] {
            color: #00B4D8;
            background-color: #2B2B2B;
        }
    </style>
""", unsafe_allow_html=True)


# -------------------------
# Helper Functions
# -------------------------
def detect_header_row(file):
    """Detects which row contains the header (looks for Product or Qty)."""
    preview = pd.read_excel(file, nrows=6, header=None)
    for i, row in preview.iterrows():
        if row.astype(str).str.contains(r'(?i)product|qty').any():
            return i
    return 0


def load_sales_data(uploaded_file):
    """Reads and cleans sales data."""
    header_row = detect_header_row(uploaded_file)
    df = pd.read_excel(uploaded_file, header=header_row)

    # Extract period info
    uploaded_file.seek(0)
    preview_lines = pd.read_excel(uploaded_file, nrows=3, header=None).astype(str)
    text = " ".join(preview_lines.iloc[:, 0].dropna())
    date_match = re.search(r"From\s*:\s*(\d{1,2}/\d{1,2}/\d{2,4}).*Upto\s*:\s*(\d{1,2}/\d{1,2}/\d{2,4})", text)
    period = None
    if date_match:
        period = f"{date_match.group(1)} → {date_match.group(2)}"

    # Normalize column names
    df.columns = df.columns.str.strip().str.replace(r'\n', ' ', regex=True)
    df.rename(columns=lambda x: str(x).strip().title(), inplace=True)

    expected_cols = [
        "Product Name", "Qty", "Free", "Rate", "Grsamt",
        "Invno", "Ledger Account", "Parent Manufacture"
    ]
    for c in expected_cols:
        if c not in df.columns:
            matches = [col for col in df.columns if c.split()[0].lower() in col.lower()]
            if matches:
                df.rename(columns={matches[0]: c}, inplace=True)

    numeric_cols = ["Qty", "Free", "Rate", "Grsamt"]
    for c in numeric_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    return df, period


def load_company_list(uploaded_file):
    """Loads the All Company List file (2 columns, both Parent Manufacturers)."""
    header_row = detect_header_row(uploaded_file)
    df = pd.read_excel(uploaded_file, header=header_row)
    all_companies = pd.unique(df.iloc[:, 0].dropna().astype(str).tolist() +
                              df.iloc[:, 1].dropna().astype(str).tolist())
    all_companies = [x.strip().upper() for x in all_companies]
    return sorted(set(all_companies))


def generate_reports(df, company_list, selected_customer):
    """Generates both reports based on selected customer."""
    cust_df = df[df["Ledger Account"] == selected_customer]

    # --- Report 1: Frequent Purchases ---
    report1 = (cust_df.groupby(["Product Name", "Parent Manufacture"])
               .agg({"Invno": "nunique", "Qty": ["max", "mean"]})
               .reset_index())
    report1.columns = ["Product Name", "Parent Manufacture",
                       "No. of Instances", "Max Purchases", "Avg Purchases"]
    report1["Avg Purchases"] = report1["Avg Purchases"].round(2)
    report1 = report1.sort_values("Avg Purchases", ascending=False)

    # --- Report 2: Company-wise Purchases ---
    if "Parent Manufacture" not in cust_df.columns:
        report2 = pd.DataFrame(columns=["Parent Manufacture", "Total Qty", "Total Value"])
    else:
        comp_df = cust_df.copy()
        comp_df["Parent Manufacture"] = comp_df["Parent Manufacture"].astype(str).str.upper()

        summary = (comp_df.groupby("Parent Manufacture")
                   .agg({"Qty": "sum", "Grsamt": "sum"})
                   .reset_index())
        summary.columns = ["Parent Manufacture", "Total Qty", "Total Value"]

        # Merge with full company list to include 0-sales companies
        full_df = pd.DataFrame({"Parent Manufacture": company_list})
        report2 = full_df.merge(summary, on="Parent Manufacture", how="left").fillna(0)
        report2["Total Qty"] = report2["Total Qty"].astype(int)
        report2["Total Value"] = report2["Total Value"].round(2)
        report2 = report2.sort_values("Total Value", ascending=False)

    return report1, report2


# -------------------------
# Main App Layout
# -------------------------
tab1, tab2 = st.tabs(["📂 Upload Files", "📊 Dashboard"])

with tab1:
    st.header("Upload Excel Files")
    st.markdown("Upload your **Sales Data** and **All Company List** files below.")
    col1, col2 = st.columns(2)
    with col1:
        sales_file = st.file_uploader("Upload Sales Data Excel", type=["xlsx"], key="sales")
    with col2:
        company_file = st.file_uploader("Upload All Company List Excel", type=["xlsx"], key="company")

    if sales_file and company_file:
        try:
            df, period = load_sales_data(sales_file)
            company_list = load_company_list(company_file)

            # Save to session
            st.session_state["sales_data"] = df
            st.session_state["period"] = period
            st.session_state["company_list"] = company_list

            st.success("✅ Files uploaded and processed successfully! Switch to the 'Dashboard' tab to view reports.")
        except Exception as e:
            st.error(f"Error reading files: {e}")
    else:
        st.info("Please upload both files to continue.")

with tab2:
    if "sales_data" not in st.session_state or "company_list" not in st.session_state:
        st.warning("⚠️ Please upload files first in the 'Upload Files' tab.")
        st.stop()

    df = st.session_state["sales_data"]
    company_list = st.session_state["company_list"]
    period = st.session_state.get("period", None)

    st.header("Customer Purchase Dashboard")

    if period:
        st.markdown(f"### 📅 Period: {period}")

    # ---- Filters ----
    st.subheader("🎯 Filters")
    customer_names = sorted(df["Ledger Account"].dropna().unique())
    selected_customer = st.selectbox("Select Customer Name (Ledger Account)", customer_names)

    if not selected_customer:
        st.info("Please select a customer to view reports.")
        st.stop()

    # ---- Reports ----
    report1, report2 = generate_reports(df, company_list, selected_customer)

    st.subheader("📈 Report 1️⃣: Frequent Purchases (Sorted by Avg Purchases ↓)")
    st.dataframe(report1, use_container_width=True, hide_index=True)

    st.subheader("🏢 Report 2️⃣: Company-wise Purchases (Includes Zero-Sales Companies ↓)")
    st.dataframe(report2, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.caption("Created by Jinesh • Powered by Streamlit ✨")
