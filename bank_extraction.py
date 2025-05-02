# Path: /extract_providus_transaction_name.py

import streamlit as st
import pandas as pd
import re
import os
import logging
from openpyxl import load_workbook
from io import BytesIO
import zipfile
import tempfile

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

st.set_page_config(page_title="S9 Bank Statement Name Extractor", layout="wide")
st.title("S9 Bank Statement Name Extractor")

st.markdown("Upload your Excel or CSV file. Use the downloadable template below with the following headers: *Date, Transaction Details/Narration, Debit, Credit, Balance*")

# Downloadable sample template
with st.expander("📥 Download Sample Template"):
    sample_excel = BytesIO()
    sample_df = pd.DataFrame(columns=["Date", "Transaction Details/Narration", "Debit", "Credit", "Balance"])
    sample_df.to_excel(sample_excel, index=False)
    sample_excel.seek(0)
    st.download_button(
        "Download sample_statement.xlsx",
        sample_excel,
        file_name="sample_statement.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# File uploader for multiple Excel/CSV files
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
]
KNOWN_NAMES_LOWER = [n.lower() for n in KNOWN_NAMES]

def clean_entity(name):
    """
    Cleans a name by removing unwanted terms and formatting it to title case.
    
    Args:
        name (str): The raw name to clean.
    
    Returns:
        str: The cleaned name in title case.
    """
    name = re.sub(r'[^A-Z\s\-]', '', name.upper())
    name = re.sub(r'\b(LIMITED|LTD|PLC|ENTERPRISE|COMPANY|TRANSFER|ACCOUNT|AC|CIB|USD FOREX PURCHASE TRANSACTION|NIP|WILLOW)\b', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name.title()

def extract_transaction_name(description):
    """
    Extracts a clean transaction name from a bank statement description.
    
    Args:
        description (str): The transaction description/narration from the bank statement.
    
    Returns:
        str: The extracted and cleaned transaction name, or empty string if input is invalid.
    
    Logic:
        1. Checks for specific patterns like 'pp_fr_chg' (Bank Charges) or 'pp_fr_vat' (VAT).
        2. Matches against a list of known names (KNOWN_NAMES).
        3. Splits description by '/' and checks parts, skipping 'contractor payment'.
        4. Applies regex to remove unwanted patterns (e.g., numbers, 'WILLOW COMMERCIAL').
        5. Extracts names from 'FROM' or 'TO' patterns.
        6. Falls back to cleaned description if no match found.
    """
    logging.debug(f"Processing description: {description}")
    if not isinstance(description, str):
        logging.warning("Description is not a string")
        return ''

    desc_lower = description.lower()
    if 'pp_fr_chg' in desc_lower:
        logging.debug("Matched 'pp_fr_chg' -> Bank Charges")
        return 'Bank Charges'
    if 'pp_fr_vat' in desc_lower:
        logging.debug("Matched 'pp_fr_vat' -> VAT")
        return 'VAT'

    for idx, name in enumerate(KNOWN_NAMES_LOWER):
        if name in desc_lower:
            logging.debug(f"Matched known name: {KNOWN_NAMES[idx]}")
            return KNOWN_NAMES[idx]

    parts = description.split('/')
    for part in parts:
        if 'contractor payment' in part.lower():
            logging.debug("Skipping part with 'contractor payment'")
            continue
        cleaned = clean_entity(part)
        if cleaned.lower() in KNOWN_NAMES_LOWER:
            logging.debug(f"Matched cleaned part: {cleaned}")
            return cleaned

    description = description.upper()
    description = re.sub(r'(\*+\d+|\d{8,}|\d{4,}/\d+|\d{5,})', '', description)
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
    if match:
        result = clean_entity(match.group(1))
        logging.debug(f"Matched FROM pattern: {result}")
        return result
    match = re.search(r'TO\s+[A-Z\s]+/\s*([A-Z\s\-\.]+)', description)
    if match:
        result = clean_entity(match.group(1))
        logging.debug(f"Matched TO pattern: {result}")
        return result

    parts = re.split(r'\||/', description)
    for part in reversed(parts):
        part = clean_entity(part)
        if len(part.split()) >= 2:
            logging.debug(f"Matched multi-word part: {part}")
            return part

    result = clean_entity(description)
    logging.debug(f"Fallback to cleaned description: {result}")
    return result

def save_excel(df, output_path):
    """
    Saves a DataFrame to Excel without highlighting.
    
    Args:
        df (pd.DataFrame): The DataFrame to save.
        output_path (str): The path to save the Excel file.
    """
    df.to_excel(output_path, index=False)
    logging.info(f"Saved Excel file: {output_path}")

def export_summary(df):
    """
    Creates a summary of transactions grouped by Extracted Name.
    
    Args:
        df (pd.DataFrame): The DataFrame containing transaction data.
    
    Returns:
        pd.DataFrame: A DataFrame summarizing transaction counts by Extracted Name.
    """
    return df.groupby('Extracted Name').size().reset_index(name='Count')

# Process uploaded files
if uploaded_files:
    st.session_state.cleaned_files = []  # Initialize session state for cleaned files

    for uploaded_file in uploaded_files:
        logging.info(f"Processing file: {uploaded_file.name}")
        ext = os.path.splitext(uploaded_file.name)[-1].lower()
        try:
            if ext == '.csv':
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file, sheet_name=0)
        except Exception as e:
            st.error(f"Error reading file {uploaded_file.name}: {str(e)}")
            logging.error(f"Failed to read file {uploaded_file.name}: {str(e)}")
            continue

        # Normalize column names
        for col in df.columns:
            if any(k in col.lower() for k in ['narration', 'description', 'details']):
                df.rename(columns={col: 'description'}, inplace=True)
            if 'debit' in col.lower():
                df.rename(columns={col: 'debit'}, inplace=True)
            if 'credit' in col.lower():
                df.rename(columns={col: 'credit'}, inplace=True)

        # Check for required columns
        if 'description' not in df.columns:
            st.error(f"Error: File {uploaded_file.name} must contain a column for 'Narration', 'Description', or 'Details'.")
            logging.error(f"Missing description column in {uploaded_file.name}")
            continue

        # Process transactions with progress bar
        if 'description' in df.columns:
            progress_bar = st.progress(0)
            total_rows = len(df)
            df['Extracted Name'] = df['description'].apply(extract_transaction_name)
            progress_bar.progress(1.0)
            progress_bar.empty()

            st.success(f"✅ Processed: {uploaded_file.name}")
            st.dataframe(df[['description', 'Extracted Name']].head(20))

            # Display summary
            summary = export_summary(df)
            st.subheader(f"📊 Summary: {uploaded_file.name}")
            st.dataframe(summary)

            # Save cleaned file
            cleaned_name = f"cleaned_{uploaded_file.name.replace('.xlsx','').replace('.csv','')}.xlsx"
            save_excel(df, cleaned_name)
            st.session_state.cleaned_files.append(cleaned_name)

    # Provide ZIP download for all processed files
    if st.session_state.cleaned_files:
        with st.expander("⬇ Download All Processed Files as ZIP"):
            zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
            with zipfile.ZipFile(zip_buffer.name, "w") as zf:
                for cleaned_name in st.session_state.cleaned_files:
                    zf.write(cleaned_name, arcname=cleaned_name)
                    logging.info(f"Added to ZIP: {cleaned_name}")

            with open(zip_buffer.name, "rb") as f:
                st.download_button(
                    "Download All as ZIP",
                    data=f,
                    file_name="cleaned_statements.zip",
                    mime="application/zip"
                )
            logging.info("Prepared ZIP file for download")
