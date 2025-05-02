# Path: /extract_providus_transaction_name.py

import streamlit as st
import pandas as pd
import re
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="S9 Bank Statement Name Extractor", layout="wide")
st.title("S9 Bank Statement Name Extractor")

st.markdown("Upload your Excel or CSV file. Use the downloadable template below with the following headers: **Date, Transaction Details/Narration, Debit, Credit, Balance**")

with st.expander("📥 Download Sample Template"):
    from io import BytesIO
sample_excel = BytesIO()
sample_df = pd.DataFrame(columns=["Date", "Transaction Details/Narration", "Debit", "Credit", "Balance"])
sample_df.to_excel(sample_excel, index=False)
sample_excel.seek(0)
with sample_excel as f:
        st.download_button("Download sample_statement.xlsx", f, file_name="sample_statement.xlsx")

uploaded_files = st.file_uploader("Choose one or more files", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

# Known name list (pre-filled)
KNOWN_NAMES = [
    "DANGOTE CEMENT", "Agrited Nigeria Ltd", "Max air",
    "IDEOS TECHNOLOGY", "BIBAGE TECHNOLOGY",
    "Balance B/F", "Emmantexlicon Global Services", "Avut Reiche Innovations",
    "Alma Beta Agro", "Brightfield Solutions", "Nicdus Resources",
    "Blemaur Oil And Gas", "Westtech Limited", "Adekoya Adeniyi Ayobami",
    "Venture Garden Nig", "Greensource Insignia", "Coupons Retail",
    "Morgan Adebowale Omotayo", "Alade Ibijoke", "Venur Nigeria Technology",
    "Ifunanya Chinenye Igboanugo", "Nyerhovwo Alex Urhude", "Rebecca Ogochukwu Ekwueme",
    "Dealmakers Energy", "AGK ENERGIES LIMITED", "Adisa Moshood Abiola",
    "Classmobile Technologies", "Adino Partners", "Bluebulb Energy",
    "Tanout Technologies", "Alasan Muhammad Nakofa Ventures", "Ardor Innovations",
    "Tradepot", "Salmnine Investment", "Denero Global Services",
    "Ifeoluwa Damola Kuponiyi", "Obi Ernest", "Proost Integrated",
    "Faltas Innovation", "Starkraft-Nordic", "Gilbert Igweka",
    "SDR AGRO", "Zydox Oil", "Captus Consilium", "STAMP DUTY",
    "VAT", "ELECTRONIC MONEY TRANSFER LEVY", "Saravan Energy", "Kordax",
    "Gruges Energy", "Texas Multinational Resources-Web", "To Hofstede Essentials",
    "Tekwanet Business Solutions-Ft Ifo C", "Putsherd Integrated Solutions-Ft Ifo C",
    "Tmdk Terminal - Fidelity", "Chizoba Eunice Ezeonyido", "Lender Tech Solutions",
    "Dania Oluwaseun Nurudeen", "Privolt Oil And Gas Services", "Tradedepot",
    "Transactworld"
]  # keep same full name list here
KNOWN_NAMES_LOWER = [n.lower() for n in KNOWN_NAMES]

def clean_entity(name):
    name = re.sub(r'[^A-Z\s\-]', '', name.upper())
    name = re.sub(r'\b(LIMITED|LTD|PLC|ENTERPRISE|COMPANY|TRANSFER|ACCOUNT|AC|CIB|USD FOREX PURCHASE TRANSACTION|NIP|WILLOW)\b', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name.title()

def extract_transaction_name(description):
    if not isinstance(description, str): return ''
    desc_lower = description.lower()
    for idx, name in enumerate(KNOWN_NAMES_LOWER):
        if name in desc_lower:
            return KNOWN_NAMES[idx]
    description = description.upper()
    description = re.sub(r'(\*+\d+|\d{8,}|\d{4,}\/\d+|\d{5,})', '', description)
    description = re.sub(r'WILLOW COMMERCIAL( LIMITED)?', '', description)
    description = re.sub(r'(PAYT TO|RIB:PP|PRB TRSF TO WILOW|TO WILLOW)', '', description)
    description = re.sub(r'A/C TO A/C TRANSFER THROUGH IBS INTERNET\s+TRANSFER FROM', '', description)
    description = re.sub(r'-?\s*FT\s+GTL\s+WILLOW\s+COM.*$', '', description)
    description = re.sub(r'-\s*NIP$', '', description)
    description = re.sub(r'-\s*USD FOREX PURCHASE TRANSACTION$', '', description)
    description = re.sub(r'-\s*CIB$', '', description)
    description = re.sub(r'-\s*WILLOW$', '', description)
    description = re.sub(r'\bTO PAYMENT\b', '', description)
    match = re.search(r'FROM\s+[A-Z\s]+/\s*([A-Z\s\-\.]+)', description)
    if match: return clean_entity(match.group(1))
    match = re.search(r'TO\s+[A-Z\s]+/\s*([A-Z\s\-\.]+)', description)
    if match: return clean_entity(match.group(1))
    parts = re.split(r'\||/', description)
    for part in reversed(parts):
        part = clean_entity(part)
        if len(part.split()) >= 2:
            return part
    return clean_entity(description)

def categorize_transaction(row):
    desc_lower = str(row['description']).lower()
    if 'vat' in desc_lower:
        return 'VAT'
    if 'charges' in desc_lower or 'charge' in desc_lower or 'offshore charges' in desc_lower:
        return 'Bank Charges'
    if 'sms alert' in desc_lower:
        return 'Bank Charges'
    try:
        credit = float(str(row.get('Credit', 0)).replace(',', '').strip())
        debit = float(str(row.get('Debit', 0)).replace(',', '').strip())
    except (ValueError, TypeError):
        credit = 0.0
        debit = 0.0
    if debit < 50 or 'ELECTRONIC MONEY TRANSFER LEVY' in desc_lower or 'STAMP DUTY' in desc_lower:
        return 'Bank Charges'
    elif credit > 0:
        return 'In'
    elif debit > 0:
        return 'Out'
    return 'Unknown'

def highlight_and_save_excel(df, output_path):
    df.to_excel(output_path, index=False)
    wb = load_workbook(output_path)
    ws = wb.active
    fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    header = [cell.value for cell in ws[1]]
    if 'Transaction Type' in header:
        trans_type_idx = header.index('Transaction Type') + 1
        for row in ws.iter_rows(min_row=2):
            if row[trans_type_idx - 1].value == 'Bank Charges':
                for cell in row:
                    cell.fill = fill
    wb.save(output_path)

def export_summary(df):
    return df.groupby(['Extracted Name', 'Transaction Type']).size().unstack(fill_value=0)

if uploaded_files:
    for uploaded_file in uploaded_files:
        ext = os.path.splitext(uploaded_file.name)[-1].lower()
        if ext == '.csv':
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, sheet_name=0)

        for col in df.columns:
            if any(k in col.lower() for k in ['narration', 'description', 'details']):
                df.rename(columns={col: 'description'}, inplace=True)
            if 'debit' in col.lower():
                df.rename(columns={col: 'debit'}, inplace=True)
            if 'credit' in col.lower():
                df.rename(columns={col: 'credit'}, inplace=True)

        if 'description' in df.columns:
            df['Extracted Name'] = df['description'].apply(extract_transaction_name)
            df['Transaction Type'] = df.apply(categorize_transaction, axis=1)

            st.success(f"✅ Processed: {uploaded_file.name}")
            st.dataframe(df[['description', 'Extracted Name', 'Transaction Type']].head(20))

            summary = export_summary(df)
            st.subheader(f"📊 Summary: {uploaded_file.name}")
            st.dataframe(summary)

            cleaned_name = f"cleaned_{uploaded_file.name.replace('.xlsx','').replace('.csv','')}.xlsx"
            highlight_and_save_excel(df, cleaned_name)
            if 'cleaned_files' not in st.session_state:
                st.session_state.cleaned_files = []
            st.session_state.cleaned_files.append(cleaned_name)
    ext = os.path.splitext(uploaded_file.name)[-1].lower()
    if ext == '.csv':
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file, sheet_name=0)

    # Normalize column names
    for col in df.columns:
        if any(k in col.lower() for k in ['narration', 'description', 'details']):
            df.rename(columns={col: 'description'}, inplace=True)
        if 'debit' in col.lower():
            df.rename(columns={col: 'debit'}, inplace=True)
        if 'credit' in col.lower():
            df.rename(columns={col: 'credit'}, inplace=True)

    if 'description' in df.columns:
        df['Extracted Name'] = df['description'].apply(extract_transaction_name)
        df['Transaction Type'] = df.apply(categorize_transaction, axis=1)

        st.success("✅ Extraction completed!")
        st.dataframe(df[['description', 'Extracted Name', 'Transaction Type']].head(20))

        summary = export_summary(df)
        st.subheader("📊 Summary by Name & Type")
        st.dataframe(summary)

        with st.expander("⬇ Download All Processed Files as ZIP"):
    import zipfile, tempfile
    zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    with zipfile.ZipFile(zip_buffer.name, "w") as zf:
        for uploaded_file in uploaded_files:
            ext = os.path.splitext(uploaded_file.name)[-1].lower()
            if ext == '.csv':
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file, sheet_name=0)

            for col in df.columns:
                if any(k in col.lower() for k in ['narration', 'description', 'details']):
                    df.rename(columns={col: 'description'}, inplace=True)
                if 'debit' in col.lower():
                    df.rename(columns={col: 'debit'}, inplace=True)
                if 'credit' in col.lower():
                    df.rename(columns={col: 'credit'}, inplace=True)

            if 'description' in df.columns:
                df['Extracted Name'] = df['description'].apply(extract_transaction_name)
                df['Transaction Type'] = df.apply(categorize_transaction, axis=1)

                cleaned_name = f"cleaned_{uploaded_file.name.replace('.xlsx','').replace('.csv','')}.xlsx"
                highlight_and_save_excel(df, cleaned_name)
                zf.write(cleaned_name, arcname=cleaned_name)

    with open(zip_buffer.name, "rb") as f:
        st.download_button("Download All as ZIP", data=f, file_name="cleaned_statements.zip", mime="application/zip")
