# Path: /extract_providus_transaction_name.py

import streamlit as st
import pandas as pd
import re
import os
import logging
from io import BytesIO
from openpyxl import load_workbook

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

st.set_page_config(page_title="S9 Bank Statement Name Extractor", layout="wide")
st.title("S9 Bank Statement Name Extractor")

# Sidebar menu
menu = st.sidebar.selectbox("Select Menu", ["Upload", "Bank Reconciliation"])

# --- Upload Menu ---
if menu == "Upload":
    st.markdown("Upload an Excel file with multiple sheets or a CSV file. Each sheet should have headers: *Date, Transaction Details/Narration, Debit, Credit, Balance*. Use the downloadable template below.")

    # Downloadable sample template
    with st.expander("📥 Download Sample Template"):
        sample_excel = BytesIO()
        sample_df = pd.DataFrame(columns=["Date", "Transaction Details/Narration", "Debit", "Credit", "Balance"])
        with pd.ExcelWriter(sample_excel, engine='openpyxl') as writer:
            sample_df.to_excel(writer, sheet_name='Sheet1', index=False)
            sample_df.to_excel(writer, sheet_name='Sheet2', index=False)  # Example with multiple sheets
        sample_excel.seek(0)
        st.download_button(
            "Download sample_statement.xlsx",
            sample_excel,
            file_name="sample_statement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # File uploader for Excel/CSV files (single file for simplicity)
    uploaded_file = st.file_uploader("Choose an Excel or CSV file", type=["xlsx", "xls", "csv"])

    # Known name list (from provided code)
    KNOWN_NAMES = [
        "Dangote Cement", "Agrited Nigeria Ltd", "Max air",
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
        "Gruges Energy", "Texas Multinational Resources-Web", "Hofstede Essentials",
        "Tekwanet Business Solutions", "Putsherd Integrated Solutions",
        "Tmdk Terminal - Fidelity", "Chizoba Eunice Ezeonyido", "Lender Tech Solutions",
        "Dania Oluwaseun Nurudeen", "Privolt Oil And Gas Services", "Tradedepot",
        "Transactworld",
        "Airpeace", "Willow Commercial Limited", "Payaza", "Flutterwave",
        "Code Crafter", "Bridge Building", "Angel exports",
        "Trzl - Westtech Limited", "Trzl - Willow Commercial Limited", "Trzl - Techcore Limited",
        "Fidelity/Westtech", "Putsherd Integrated", "Beverly Trust",
        "Gilbert", "Rukib Heritage", "Aduroja Temilade",
        "Telex Charge", "Transfer Charge", "Ocrativane Integrated",
        "WESTTECH:FIDELITY", "Adino Global", "TRZL-TECHCORE LIMITED", "Swift Charge"
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
        # Only remove 'WILLOW COMMERCIAL' if not a known name
        if 'WILLOW COMMERCIAL LIMITED' not in desc_lower:
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

    def export_summary(df):
        """
        Creates a summary of transactions grouped by Extracted Name.
        
        Args:
            df (pd.DataFrame): The DataFrame containing transaction data.
        
        Returns:
            pd.DataFrame: A DataFrame summarizing transaction counts by Extracted Name.
        """
        return df.groupby('Extracted Name').size().reset_index(name='Count')

    # Process uploaded file
    if uploaded_file:
        logging.info(f"Processing file: {uploaded_file.name}")
        ext = os.path.splitext(uploaded_file.name)[-1].lower()
        output_excel = BytesIO()
        processed_sheets = {}

        try:
            if ext == '.csv':
                # Treat CSV as a single sheet
                df = pd.read_csv(uploaded_file)
                processed_sheets['Sheet1'] = df
            else:
                # Read all sheets from Excel
                processed_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        except Exception as e:
            st.error(f"Error reading file {uploaded_file.name}: {str(e)}")
            logging.error(f"Failed to read file {uploaded_file.name}: {str(e)}")
            st.stop()

        # Initialize progress bar
        total_sheets = len(processed_sheets)
        progress_bar = st.progress(0)
        current_progress = 0

        # Process each sheet
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            for sheet_name, df in processed_sheets.items():
                logging.info(f"Processing sheet: {sheet_name}")
                
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
                    st.error(f"Error: Sheet '{sheet_name}' in {uploaded_file.name} must contain a column for 'Narration', 'Description', or 'Details'.")
                    logging.error(f"Missing description column in sheet {sheet_name}")
                    continue

                # Process transactions
                df['Extracted Name'] = df['description'].apply(extract_transaction_name)

                # Write processed DataFrame to output Excel
                output_sheet_name = f"Processed_{sheet_name}"
                df.to_excel(writer, sheet_name=output_sheet_name, index=False)
                logging.info(f"Wrote processed data to sheet: {output_sheet_name}")

                # Display results
                st.success(f"✅ Processed sheet: {sheet_name}")
                st.subheader(f"Preview: {sheet_name}")
                st.dataframe(df[['description', 'Extracted Name']].head(20))

                if df.empty or 'Extracted Name
                    continue

                # Display summary
                summary = export_summary(df)
                st.subheader(f"📊 Summary: {sheet_name}")
                st.dataframe(summary)

                # Update progress
                current_progress += 1
                progress_bar.progress(current_progress / total_sheets)

        progress_bar.empty()
        output_excel.seek(0)

        # Provide download button for the processed Excel file
        st.download_button(
            "Download Processed Excel File",
            output_excel,
            file_name=f"processed_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        logging.info(f"Prepared output Excel file for download: processed_{uploaded_file.name}")

# --- Bank Reconciliation Menu ---
elif menu == "Bank Reconciliation":
    st.markdown("Enter your opening balance, add transactions, and reconcile with the bank statement balance.")

    # Initialize session state for transactions
    if 'transactions' not in st.session_state:
        st.session_state.transactions = []

    # Input for opening balance
    opening_balance = st.number_input("Opening Balance", value=0.0, format="%.2f")

    # Form for adding transactions
    with st.form(key='transaction_form'):
        st.subheader("Add Transaction")
        trans_type = st.selectbox("Transaction Type", ["Addition", "Subtraction"])
        trans_amount = st.number_input("Amount", min_value=0.0, value=0.0, format="%.2f")
        trans_desc = st.text_input("Description (Optional)")
        submit_button = st.form_submit_button("Add Transaction")

        if submit_button:
            if trans_amount > 0:
                st.session_state.transactions.append({
                    'Type': trans_type,
                    'Amount': trans_amount,
                    'Description': trans_desc
                })
                st.success("Transaction added!")
            else:
                st.error("Amount must be greater than 0.")

    # Input for bank statement balance
    bank_statement_balance = st.number_input("Bank Statement Balance", value=0.0, format="%.2f")

    # Calculate reconciliation
    if st.button("Reconcile"):
        if not st.session_state.transactions:
            st.warning("No transactions added.")
        else:
            closing_balance = opening_balance
            for trans in st.session_state.transactions:
                if trans['Type'] == "Addition":
                    closing_balance += trans['Amount']
                else:
                    closing_balance -= trans['Amount']

            discrepancy = closing_balance - bank_statement_balance

            # Prepare reconciliation table
            recon_data = [
                {'Type': 'Opening Balance', 'Amount': opening_balance, 'Description': 'Starting balance'}
            ] + st.session_state.transactions + [
                {'Type': 'Closing Balance', 'Amount': closing_balance, 'Description': 'Calculated balance'},
                {'Type': 'Bank Statement Balance', 'Amount': bank_statement_balance, 'Description': 'Per bank statement'},
                {'Type': 'Discrepancy', 'Amount': discrepancy, 'Description': 'Difference (Closing - Bank)'}
            ]
            recon_df = pd.DataFrame(recon_data)

            # Display reconciliation table
            st.subheader("Reconciliation Summary")
            st.dataframe(recon_df)

            # Download reconciliation as Excel
            output_recon = BytesIO()
            recon_df.to_excel(output_recon, index=False, sheet_name='Reconciliation')
            output_recon.seek(0)
            st.download_button(
                "Download Reconciliation Excel",
                output_recon,
                file_name="bank_reconciliation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # Option to clear transactions
    if st.button("Clear Transactions"):
        st.session_state.transactions = []
        st.success("Transactions cleared.")
