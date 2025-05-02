import streamlit as st
import pandas as pd
import re
import os
from io import BytesIO
import logging
import xlsxwriter

st.set_page_config(page_title="Bank Statement Processor", layout="wide")

# Sidebar navigation
with st.sidebar:
    selected_menu = st.selectbox("📂 Select Menu", ["Upload & Extract Names", "Bank Reconciliation"])

# -------------------- MENU 1: Upload & Extract Names --------------------
if selected_menu == "Upload & Extract Names":
    st.title("📄 Upload Bank Statement")
    st.markdown("Upload Excel with tabs. Each tab should have: Date, Description, Debit, Credit, Balance.")

    uploaded_file = st.file_uploader("Upload Bank Statement Excel", type=["xlsx"])

    def clean_entity(name):
        name = re.sub(r'[^A-Z\s\-]', '', name.upper())
        name = re.sub(r'\b(LIMITED|LTD|PLC|ENTERPRISE|TRANSFER|ACCOUNT|AC|USD|NIP)\b', '', name)
        name = re.sub(r'\s+', ' ', name).strip()
        return name.title()

    def extract_transaction_name(description):
        if not isinstance(description, str):
            return ''
        description = description.upper()
        parts = re.split(r'\||/', description)
        for part in reversed(parts):
            part = clean_entity(part)
            if len(part.split()) >= 2:
                return part
        return clean_entity(description)

    if uploaded_file:
        xl = pd.ExcelFile(uploaded_file)
        all_sheets = xl.sheet_names
        output_excel = BytesIO()

        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            workbook = writer.book
            money_fmt = workbook.add_format({'num_format': '#,##0.00'})
            head_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})

            for sheet in all_sheets:
                df = xl.parse(sheet)
                df.columns = [c.lower() for c in df.columns]

                desc_col = next((c for c in df.columns if 'narration' in c or 'description' in c or 'details' in c), None)
                if not desc_col:
                    st.error(f"Missing description column in {sheet}")
                    continue

                df['description'] = df[desc_col]
                df['Extracted Name'] = df['description'].apply(extract_transaction_name)

                for col in ['debit', 'credit', 'balance']:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce')

                pivot = df.groupby('Extracted Name').agg({
                    'debit': 'sum',
                    'credit': 'sum'
                }).reset_index()

                df.to_excel(writer, sheet_name=f"Processed_{sheet}", index=False)
                pivot.to_excel(writer, sheet_name=f"Pivot_{sheet}", index=False)

                ws = writer.sheets[f"Pivot_{sheet}"]
                for i, col in enumerate(pivot.columns):
                    ws.write(0, i, col, head_fmt)
                    ws.set_column(i, i, 22, money_fmt)

                st.subheader(f"📊 Pivot Table for {sheet}")
                st.dataframe(pivot)

        output_excel.seek(0)
        st.download_button("📥 Download Processed Excel", output_excel, file_name="processed_output.xlsx")

# -------------------- MENU 2: Bank Reconciliation --------------------
elif selected_menu == "Bank Reconciliation":
    st.title("🏦 Bank Reconciliation")

    recon_file = st.file_uploader("Upload Reconciliation Excel File", type=["xlsx"], key="recon_upload")

    if recon_file:
        xl = pd.ExcelFile(recon_file)
        all_sheets = xl.sheet_names
        total_usd = 0.0
        total_ngn = 0.0
        balance_data = []
        pivot_data = {}

        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            workbook = writer.book
            money_fmt = workbook.add_format({'num_format': '#,##0.00'})
            head_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})

            for sheet in all_sheets:
                df = xl.parse(sheet)
                df.columns = [col.lower() for col in df.columns]

                debit_col = next((col for col in df.columns if 'debit' in col), None)
                credit_col = next((col for col in df.columns if 'credit' in col), None)
                balance_col = next((col for col in df.columns if 'balance' in col), None)
                extracted_col = next((col for col in df.columns if 'extracted name' in col), None)

                for col in [debit_col, credit_col, balance_col]:
                    if col:
                        df[col] = pd.to_numeric(df[col], errors='coerce')

                if extracted_col:
                    pivot = df.groupby(df[extracted_col]).agg({
                        debit_col: 'sum',
                        credit_col: 'sum'
                    }).reset_index()
                    pivot_data[sheet] = pivot
                    pivot.to_excel(writer, sheet_name=f"Pivot_{sheet}", index=False)
                    ws = writer.sheets[f"Pivot_{sheet}"]
                    for i, col in enumerate(pivot.columns):
                        ws.write(0, i, col, head_fmt)
                        ws.set_column(i, i, 22, money_fmt)

                    st.markdown(f"### 🔍 Pivot Table - **{sheet}**")
                    st.dataframe(pivot)

                if balance_col and not df[balance_col].dropna().empty:
                    closing = df[balance_col].dropna().iloc[-1]
                    currency = 'USD' if 'usd' in sheet.lower() else 'NGN'
                    balance_data.append((sheet, closing, currency))
                    if currency == 'USD':
                        total_usd += closing
                    else:
                        total_ngn += closing

                st.divider()

            # Summary
            closing_df = pd.DataFrame(balance_data, columns=["Sheet", "Closing Balance", "Currency"])
            closing_df.to_excel(writer, sheet_name="Closing Balances", index=False)
            ws = writer.sheets["Closing Balances"]
            for i, col in enumerate(closing_df.columns):
                ws.write(0, i, col, head_fmt)
                ws.set_column(i, i, 20, money_fmt)

        output_excel.seek(0)

        st.subheader("📈 Summary of Closing Balances")
        st.dataframe(closing_df)

        st.markdown("---")
        col1, col2 = st.columns(2)
        col1.metric("💵 Total USD", f"${total_usd:,.2f}")
        col2.metric("💰 Total NGN", f"₦{total_ngn:,.2f}")

        st.download_button("📥 Download Reconciliation Excel", output_excel, file_name="reconciliation_output.xlsx")
