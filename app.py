import os
import pandas as pd
import streamlit as st

# ---------- CONFIG ----------
st.set_page_config(
    page_title="Customer Purchase Dashboard",
    layout="wide",
    page_icon="📊",
)

# ---------- DARK THEME ----------
st.markdown("""
    <style>
        body {background-color: #0E1117; color: #FAFAFA;}
        .stApp {background-color: #0E1117;}
        div[data-testid="stSidebar"] {background-color: #1E1E1E;}
        .block-container {padding-top: 1rem; padding-bottom: 1rem;}
        h1, h2, h3, h4, h5 {color: #FAFAFA;}
        table {color: #FAFAFA !important;}
    </style>
""", unsafe_allow_html=True)

# ---------- DATA PATH ----------
DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)
SALES_FILE = os.path.join(DATA_DIR, "sales_data.xlsx")
COMPANY_FILE = os.path.join(DATA_DIR, "company_list.xlsx")

# ---------- HELPER FUNCTIONS ----------
def clean_sales_excel(file_path):
    """Skip unwanted rows and parse proper data + extract period."""
    df_raw = pd.read_excel(file_path, header=None)
    start_idx = None
    for i in range(len(df_raw)):
        row = df_raw.iloc[i].astype(str).str.lower()
        if "ledger account" in " ".join(row):
            start_idx = i
            break
    df = pd.read_excel(file_path, header=start_idx)
    df.columns = df.columns.str.strip()

    # Extract date period (from first few rows)
    period_line = " ".join(df_raw.iloc[0:3, 0].astype(str).tolist())
    from_date, upto_date = None, None
    if "from" in period_line.lower() and "upto" in period_line.lower():
        try:
            part = period_line.split("From")[-1].split("Upto")
            from_date = part[0].replace(":", "").strip()
            upto_date = part[1].replace(":", "").strip()
        except Exception:
            pass
    return df, from_date, upto_date


def load_company_list(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()
    companies = set()
    for col in df.columns:
        companies.update(df[col].dropna().astype(str).tolist())
    return list(companies)


def generate_report_1(df, selected_customer):
    df_cust = df[df["Ledger Account"] == selected_customer]
    report = (
        df_cust.groupby("Parent Manufacturer")
        .agg({"Qty": "sum", "Value": "sum"})
        .reset_index()
    )
    report["Avg Purchase"] = (report["Value"] / report["Qty"]).round(2)
    report = report.sort_values(by="Avg Purchase", ascending=False)
    return report


def generate_report_2(df, companies):
    report = (
        df.groupby("Parent Manufacturer")
        .agg({"Qty": "sum", "Value": "sum"})
        .reset_index()
    )
    # Ensure all companies appear (even with 0s)
    all_df = pd.DataFrame({"Parent Manufacturer": companies})
    merged = pd.merge(all_df, report, on="Parent Manufacturer", how="left").fillna(0)
    merged["Qty"] = merged["Qty"].astype(int)
    merged["Value"] = merged["Value"].astype(float).round(2)
    merged = merged.sort_values(by="Value", ascending=False)
    return merged

# ---------- APP UI ----------
st.title("📊 Customer Purchase Dashboard")

tab1, tab2 = st.tabs(["🗂️ Upload Files", "📈 Dashboard"])

# ---------- TAB 1: UPLOAD ----------
with tab1:
    st.header("Upload Data Files")
    st.markdown("Upload or replace the latest Excel sheets below. These files persist for all users until replaced.")
    uploaded_sales = st.file_uploader("Upload Sales Data (.xlsx)", type=["xlsx"])
    uploaded_companies = st.file_uploader("Upload Company List (.xlsx)", type=["xlsx"])

    if uploaded_sales is not None:
        with open(SALES_FILE, "wb") as f:
            f.write(uploaded_sales.getbuffer())
        st.success("✅ Sales data uploaded and saved globally!")

    if uploaded_companies is not None:
        with open(COMPANY_FILE, "wb") as f:
            f.write(uploaded_companies.getbuffer())
        st.success("✅ Company list uploaded and saved globally!")

    if os.path.exists(SALES_FILE) and os.path.exists(COMPANY_FILE):
        st.info("✅ Both files are present. You can view reports in the Dashboard tab.")
    else:
        st.warning("Please upload both required files to activate the dashboard.")

# ---------- TAB 2: DASHBOARD ----------
with tab2:
    if not (os.path.exists(SALES_FILE) and os.path.exists(COMPANY_FILE)):
        st.warning("⚠️ Please upload both files first from the Upload tab.")
    else:
        sales_df, from_date, upto_date = clean_sales_excel(SALES_FILE)
        company_list = load_company_list(COMPANY_FILE)

        # Filter and Reports
        st.header("Dashboard")
        if from_date and upto_date:
            st.markdown(f"**🗓️ Period:** {from_date} → {upto_date}")

        customers = sales_df["Ledger Account"].dropna().unique().tolist()
        selected_customer = st.selectbox("Select Customer (Ledger Account)", customers)

        if selected_customer:
            st.subheader("Report 1: Customer Purchase Summary")
            rep1 = generate_report_1(sales_df, selected_customer)
            st.dataframe(rep1, use_container_width=True, hide_index=True)

            st.subheader("Report 2: Company-wise Monthly Summary")
            rep2 = generate_report_2(sales_df, company_list)
            st.dataframe(rep2, use_container_width=True, hide_index=True)
