import streamlit as st
import pandas as pd
import re
import os
from io import BytesIO
import logging
import xlsxwriter
import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="S9 Bank Statement Processor", layout="wide")
logging.basicConfig(level=logging.INFO)

@st.cache_data(ttl=3600)
def load_known_names():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_dict = json.loads(st.secrets["GOOGLE_CREDS"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1n7Uifi2bxK8Fz_arVxHvdeoL6m_d_8UR8QDpbNgsozs/edit#gid=0")
    worksheet = sheet.get_worksheet(0)
    names = worksheet.col_values(1)
    names = [name.strip() for name in names if name.strip() and name.lower() != 'name']
    return names, [n.lower() for n in names]

KNOWN_NAMES, KNOWN_NAMES_LOWER = load_known_names()

with st.sidebar:
    selected_menu = st.selectbox("📂 Select Menu", ["Upload & Extract Names", "Bank Reconciliation"])

if selected_menu == "Upload & Extract Names":
    st.title("📄 Upload Bank Statement")
    uploaded_file = st.file_uploader("Upload Excel or CSV File", type=["xlsx", "xls", "csv"])

    def clean_entity(name):
        name = re.sub(r'[^A-Z\\s\\-]', '', name.upper())
        name = re.sub(r'\\b(LIMITED|LTD|PLC|ENTERPRISE|ACCOUNT|AC|USD FOREX PURCHASE TRANSACTION|NIP|WILLOW)\\b', '', name)
        name = re.sub(r'\\s+', ' ', name).strip()
        return name.title()

    def extract_transaction_name(description):
        if not isinstance(description, str):
            return ''
        desc_lower = description.lower()
        for idx, name in enumerate(KNOWN_NAMES_LOWER):
            if name in desc_lower:
                return KNOWN_NAMES[idx]
        parts = re.split(r'\\||/', description)
        for part in reversed(parts):
            cleaned = clean_entity(part)
            if len(cleaned.split()) >= 2:
                return cleaned
        return clean_entity(description)

    def export_summary(df):
        return df.groupby('Extracted Name').size().reset_index(name='Count')

    if uploaded_file:
        ext = os.path.splitext(uploaded_file.name)[-1].lower()
        output_excel = BytesIO()
        processed_sheets = {}

        try:
            if ext == '.csv':
                df = pd.read_csv(uploaded_file)
                processed_sheets['Sheet1'] = df
            else:
                processed_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        except Exception as e:
            st.error(f"Error reading file: {e}")
            st.stop()

        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            for sheet_name, df in processed_sheets.items():
                df.columns = [c.lower() for c in df.columns]
                for col in df.columns:
                    if any(k in col.lower() for k in ['narration', 'description', 'details']):
                        df.rename(columns={col: 'description'}, inplace=True)
                    if 'debit' in col.lower():
                        df.rename(columns={col: 'debit'}, inplace=True)
                    if 'credit' in col.lower():
                        df.rename(columns={col: 'credit'}, inplace=True)

                if 'description' not in df.columns:
                    st.error(f"Missing 'description' column in {sheet_name}")
                    continue

                df['Extracted Name'] = df['description'].apply(extract_transaction_name)

                for col in ['debit', 'credit', 'balance']:
                    if col in df.columns:
                        df[col] = df[col].astype(str).str.replace(",", "").str.strip()
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

                df.to_excel(writer, sheet_name=f"Processed_{sheet_name}", index=False)

                st.success(f"✅ Processed: {sheet_name}")
                st.subheader(f"🔍 Preview: {sheet_name}")
                st.dataframe(df[['description', 'Extracted Name', 'debit', 'credit']].head(20))

                summary = export_summary(df)
                st.subheader(f"📊 Summary: {sheet_name}")
                st.dataframe(summary)

        output_excel.seek(0)
        st.download_button("📥 Download Processed Excel File", output_excel, file_name=f"processed_{uploaded_file.name}")

elif selected_menu == "Bank Reconciliation":
    st.title("🏦 Bank Reconciliation")
    recon_file = st.file_uploader("Upload Processed Excel File", type=["xlsx"], key="recon_file")

    if recon_file:
        xl = pd.ExcelFile(recon_file)
        sheet_names = [s for s in xl.sheet_names if s.startswith("Processed_")]
        total_usd = 0.0
        total_ngn = 0.0
        balance_summary = []
        output_excel = BytesIO()

        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            workbook = writer.book
            head_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
            money_fmt = workbook.add_format({'num_format': '#,##0.00'})

            for sheet in sheet_names:
                df = xl.parse(sheet)
                df.columns = [c.lower() for c in df.columns]

                extracted = next((c for c in df.columns if 'extracted name' in c), None)
                debit = next((c for c in df.columns if 'debit' in c), None)
                credit = next((c for c in df.columns if 'credit' in c), None)
                balance = next((c for c in df.columns if 'balance' in c), None)

                for col in [debit, credit, balance]:
                    if col:
                        df[col] = pd.to_numeric(df[col], errors='coerce')

                if extracted:
                    pivot = df.groupby(df[extracted]).agg({
                        debit: 'sum',
                        credit: 'sum'
                    }).reset_index()
                    safe_sheet_name = f"Pivot_{sheet}"[:31]
                    safe_sheet_name = re.sub(r'[\\/*?:\\[\\]]', '', safe_sheet_name)
                    pivot.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                    ws = writer.sheets[safe_sheet_name]  # ✅ consistent name

                    for i, col in enumerate(pivot.columns):
                        ws.write(0, i, col, head_fmt)
                        ws.set_column(i, i, 22, money_fmt)

                    st.markdown(f"### 📊 Pivot Table: {sheet.replace('Processed_', '')}")
                    st.dataframe(pivot)

                if balance and not df[balance].dropna().empty:
                    closing = df[balance].dropna().iloc[-1]
                    currency = "USD" if "usd" in sheet.lower() else "NGN"
                    balance_summary.append((sheet.replace("Processed_", ""), closing, currency))
                    if currency == "USD":
                        total_usd += closing
                    else:
                        total_ngn += closing

                st.divider()

            summary_df = pd.DataFrame(balance_summary, columns=["Sheet", "Closing Balance", "Currency"])
            summary_df.to_excel(writer, sheet_name="Closing Balances", index=False)
            ws = writer.sheets["Closing Balances"]
            for i, col in enumerate(summary_df.columns):
                ws.write(0, i, col, head_fmt)
                ws.set_column(i, i, 20, money_fmt)

        st.subheader("💼 Closing Balance Summary")
        st.dataframe(summary_df)
        st.markdown("---")
        col1, col2 = st.columns(2)
        col1.metric("USD Total", f"${total_usd:,.2f}")
        col2.metric("NGN Total", f"₦{total_ngn:,.2f}")
        output_excel.seek(0)
        st.download_button("📥 Download Reconciliation Report", output_excel, file_name="reconciliation_output.xlsx")
'''
)
