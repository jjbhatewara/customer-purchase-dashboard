# Customer Purchase Reports Dashboard

This is a Streamlit app that allows uploading a distributor Excel file (data starting from row 4) and generates two reports per selected customer:
- Report 1: Frequent Purchases (product-level counts, max and average quantity)
- Report 2: Company-wise Monthly Purchases (pivot table by Parent Manufacturer and month)

## Run locally

1. Create a virtual environment and activate it.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the app:
   ```bash
   streamlit run app.py
   ```

## Deploy to Streamlit Cloud

1. Push this folder to a GitHub repository.
2. Log in to https://share.streamlit.io and connect your GitHub repo.
3. Select the repository and branch, set `app.py` as the main file, and deploy.# customer-purchase-dashboard
