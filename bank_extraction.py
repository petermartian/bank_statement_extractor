import streamlit as st
import pandas as pd
import re
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import zipfile
import tempfile
from io import BytesIO

st.set_page_config(page_title="S9 Bank Statement Name Extractor", layout="wide")
st.title("S9 Bank Statement Name Extractor")

st.markdown(
    "Upload your Excel or CSV file. Use the downloadable template below with the following headers: **Date, Transaction Details/Narration, Debit, Credit, Balance**"
)

with st.expander("📥 Download Sample Template"):
    sample_excel = BytesIO()
    sample_df = pd.DataFrame(
        columns=["Date", "Transaction Details/Narration", "Debit", "Credit", "Balance"]
    )
    sample_df.to_excel(sample_excel, index=False)
    sample_excel.seek(0)
    st.download_button(
        "Download sample_statement.xlsx",
        sample_excel,
        file_name="sample_statement.xlsx",
    )

uploaded_files = st.file_uploader(
    "Choose one or more files", type=["xlsx", "xls", "csv"], accept_multiple_files=True
)

# Known name list (pre-filled)
KNOWN_NAMES = [
    "ABRAHAM O.", "ABRAHAM OGBONNAYA", "ABRAHAM OGBONNAYA & SONS",
    "AGU EMEKA", "AGU EMEKA & SONS", "AJAH ANTHONY", "AJAH ANTHONY & SONS",
    "AKPAN EMMANUEL", "AKPAN EMMANUEL & SONS", "AKPAN JOSHUA", "AKPAN JOSHUA & SONS",
    "ALEXANDER OKON", "ALEXANDER OKON & SONS", "ALIYU DANJUMA", "ALIYU DANJUMA & SONS",
    "AMADI CHINEDU", "AMADI CHINEDU & SONS", "AMAECHI PETER", "AMAECHI PETER & SONS",
    "AMINU GARBA", "AMINU GARBA & SONS", "ANIEKWE CHIBUZOR", "ANIEKWE CHIBUZOR & SONS",
    "ANYA CHUKWU", "ANYA CHUKWU & SONS", "ANYANWU KINGSLEY", "ANYANWU KINGSLEY & SONS",
    "ARTHUR IBEKWE", "ARTHUR IBEKWE & SONS", "ATUOBI CHIKA", "ATUOBI CHIKA & SONS",
    "AYODELE SAMUEL", "AYODELE SAMUEL & SONS", "AZUBUIKE EKENE", "AZUBUIKE EKENE & SONS",
    "BABATUNDE ADEKUNLE", "BABATUNDE ADEKUNLE & SONS", "BALOGUN IBRAHIM", "BALOGUN IBRAHIM & SONS",
    "BASSEY EFFIONG", "BASSEY EFFIONG & SONS", "BELLO MUSA", "BELLO MUSA & SONS",
    "CHIBUIKE DAVID", "CHIBUIKE DAVID & SONS", "CHIDERA IKECHUKWU", "CHIDERA IKECHUKWU & SONS",
    "CHIKWENDU EMEKA", "CHIKWENDU EMEKA & SONS", "CHIMEZIE OBINNA", "CHIMEZIE OBINNA & SONS",
    "CHINWEOKE KINGSLEY", "CHINWEOKE KINGSLEY & SONS", "CHRISTOPHER OKORO", "CHRISTOPHER OKORO & SONS",
    "DANIEL OLAJIDE", "DANIEL OLAJIDE & SONS", "DAVID EMEKA", "DAVID EMEKA & SONS",
    "EBERECHI CHUKWUEMEKA", "EBERECHI CHUKWUEMEKA & SONS", "EDET EKPENYONG", "EDET EKPENYONG & SONS",
    "EDWARD OGBONNA", "EDWARD OGBONNA & SONS", "EGWUATU IKENNA", "EGWUATU IKENNA & SONS",
    "EKEMEZIE CHUKWUDI", "EKEMEZIE CHUKWUDI & SONS", "EKERE SUNDAY", "EKERE SUNDAY & SONS",
    "EKEZIE KINGSLEY", "EKEZIE KINGSLEY & SONS", "EMEKA CHUKWUEMEKA", "EMEKA CHUKWUEMEKA & SONS",
    "EMMANUEL OBI", "EMMANUEL OBI & SONS", "ESSIEN ETIM", "ESSIEN ETIM & SONS",
    "ETUKUDO OKON", "ETUKUDO OKON & SONS", "EZEKIEL UDUAK", "EZEKIEL UDUAK & SONS",
    "FAGBEMI OLUWASEUN", "FAGBEMI OLUWASEUN & SONS", "FELIX CHUKWUEMEKA", "FELIX CHUKWUEMEKA & SONS",
    "GABRIEL EMMANUEL", "GABRIEL EMMANUEL & SONS", "HABIB YUSUF", "HABIB YUSUF & SONS",
    "HARRISON IKECHUKWU", "HARRISON IKECHUKWU & SONS", "HASSAN IBRAHIM", "HASSAN IBRAHIM & SONS",
    "IBRAHIM USMAN", "IBRAHIM USMAN & SONS", "IDRIS GARBA", "IDRIS GARBA & SONS",
    "IKENNA EZE", "IKENNA EZE & SONS", "IKPEH EMMANUEL", "IKPEH EMMANUEL & SONS",
    "ISAAC CHUKWUEMEKA", "ISAAC CHUKWUEMEKA & SONS", "ISHIAKU GARBA", "ISHIAKU GARBA & SONS",
    "ISRAEL EMMANUEL", "ISRAEL EMMANUEL & SONS", "JAPHETH EMMANUEL", "JAPHETH EMMANUEL & SONS",
    "JOHNSON UCHECHUKWU", "JOHNSON UCHECHUKWU & SONS", "JOSEPH EMMANUEL", "JOSEPH EMMANUEL & SONS",
    "KABIRU MUSA", "KABIRU MUSA & SONS", "KOLAWOLE OLADIPO", "KOLAWOLE OLADIPO & SONS",
    "LAWRENCE OBI", "LAWRENCE OBI & SONS", "MABEL OKORO", "MABEL OKORO & SONS",
    "MADUABUCHI IKECHUKWU", "MADUABUCHI IKECHUKWU & SONS", "MAHMUD ABUBAKAR", "MAHMUD ABUBAKAR & SONS",
    "MAMMAN GARBA", "MAMMAN GARBA & SONS", "MARK EMEKA", "MARK EMEKA & SONS",
    "MARTINS CHUKWUEMEKA", "MARTINS CHUKWUEMEKA & SONS", "MATTHEW OGBONNA", "MATTHEW OGBONNA & SONS",
    "MICHAEL EZE", "MICHAEL EZE & SONS", "MONDAY EMMANUEL", "MONDAY EMMANUEL & SONS",
    "MOSES CHUKWUEMEKA", "MOSES CHUKWUEMEKA & SONS", "MUSA IBRAHIM", "MUSA IBRAHIM & SONS",
    "MUSTAPHA GARBA", "MUSTAPHA GARBA & SONS", "NDUBUISI CHUKWUEMEKA", "NDUBUISI CHUKWUEMEKA & SONS",
    "NNAMDI KINGSLEY", "NNAMDI KINGSLEY & SONS", "NWABUEZE CHUKWUEMEKA", "NWABUEZE CHUKWUEMEKA & SONS",
    "NWANKWO EMEKA", "NWANKWO EMEKA & SONS", "OBASI EMEKA", "OBASI EMEKA & SONS",
    "OBINNA CHUKWUEMEKA", "OBINNA CHUKWUEMEKA & SONS", "ODIONYENFELE EMMANUEL", "ODIONYENFELE EMMANUEL & SONS",
    "OGBONNAYA CHUKWUEMEKA", "OGBONNAYA CHUKWUEMEKA & SONS", "OKAFOR CHUKWUEMEKA", "OKAFOR CHUKWUEMEKA & SONS",
    "OKECHUKWU EMMANUEL", "OKECHUKWU EMMANUEL & SONS", "OKONKWO CHUKWUEMEKA", "OKONKWO CHUKWUEMEKA & SONS",
    "OKORO CHUKWUEMEKA", "OKORO CHUKWUEMEKA & SONS", "OKORO EMMANUEL", "OKORO EMMANUEL & SONS",
    "OLADAPO OLADIMEJI", "OLADAPO OLADIMEJI & SONS", "OLALEKAN ADEBAYO", "OLALEKAN ADEBAYO & SONS",
    "OLUWAFEMI DAVID", "OLUWAFEMI DAVID & SONS", "OMENAZU CHUKWUEMEKA", "OMENAZU CHUKWUEMEKA & SONS",
    "ONWUASOANYA CHUKWUEMEKA", "ONWUASOANYA CHUKWUEMEKA & SONS", "ONYEBUCHI CHUKWUEMEKA", "ONYEBUCHI CHUKWUEMEKA & SONS",
    "ORJI CHUKWUEMEKA", "ORJI CHUKWUEMEKA & SONS", "OSITA CHUKWUEMEKA", "OSITA CHUKWUEMEKA & SONS",
    "OYEBANJI ADEKUNLE", "OYEBANJI ADEKUNLE & SONS", "PATRICK CHUKWUEMEKA", "PATRICK CHUKWUEMEKA & SONS",
    "PAUL CHUKWUEMEKA", "PAUL CHUKWUEMEKA & SONS", "PETER CHUKWUEMEKA", "PETER CHUKWUEMEKA & SONS",
    "RAPHAEL CHUKWUEMEKA", "RAPHAEL CHUKWUEMEKA & SONS", "RICHARD CHUKWUEMEKA", "RICHARD CHUKWUEMEKA & SONS",
    "ROBERT CHUKWUEMEKA", "ROBERT CHUKWUEMEKA & SONS", "SAMUEL CHUKWUEMEKA", "SAMUEL CHUKWUEMEKA & SONS",
    "SUNDAY EMMANUEL", "SUNDAY EMMANUEL & SONS", "TAIWO ADEKUNLE", "TAIWO ADEKUNLE & SONS",
    "TIMOTHY CHUKWUEMEKA", "TIMOTHY CHUKWUEMEKA & SONS", "TITUS CHUKWUEMEKA", "TITUS CHUKWUEMEKA & SONS",
    "UCHECHUKWU EMMANUEL", "UCHECHUKWU EMMANUEL & SONS", "UDO EKPENYONG", "UDO EKPENYONG & SONS",
    "UMAR GARBA", "UMAR GARBA & SONS", "VICTOR CHUKWUEMEKA", "VICTOR CHUKWUEMEKA & SONS",
    "VINCENT CHUKWUEMEKA", "VINCENT CHUKWUEMEKA & SONS", "YAKUBU GARBA", "YAKUBU GARBA & SONS",
    "YUSUF AHMED", "YUSUF AHMED & SONS", "ZAKARIYA USMAN", "ZAKARIYA USMAN & SONS",
    "ZENITH BANK PLC", "ACCESS BANK PLC", "FIRST BANK OF NIGERIA LTD", "GUARANTY TRUST BANK PLC",
    "UNITED BANK FOR AFRICA PLC", "STERLING BANK PLC", "FIDELITY BANK PLC", "ECOBANK NIGERIA LTD",
    "KEYSTONE BANK LTD", "STANBIC IBTC BANK PLC", "SKYE BANK PLC", "UNITY BANK PLC",
    "WEMA BANK PLC", "HERITAGE BANK PLC", "PROVIDUS BANK PLC", "POLARIS BANK LTD",
    "JAIZ BANK PLC", "TITAN TRUST BANK LTD", "SUNTRUST BANK NIGERIA LTD",
    "FCMB GROUP PLC", "UNION BANK OF NIGERIA PLC", "STANDARD CHARTERED BANK NIGERIA LTD",
    "CITIBANK NIGERIA LTD", "RAND MERCHANT BANK NIGERIA LTD",
    "NOVA MERCHANT BANK LTD", "CORONATION MERCHANT BANK LTD",
    "FBNQUEST MERCHANT BANK LTD", "OPTIMUS BANK LTD"
]
KNOWN_NAMES_LOWER = [n.lower() for n in KNOWN_NAMES]


def clean_entity(name):
    name = re.sub(r'[^A-Z\s\-]', '', name.upper())
    name = re.sub(
        r'\b(LIMITED|LTD|PLC|ENTERPRISE|COMPANY|TRANSFER|ACCOUNT|AC|CIB|USD FOREX PURCHASE TRANSACTION|NIP|WILLOW)\b',
        '', name,
    )
    name = re.sub(r'\s+', ' ', name).strip()
    return name.title()


def extract_transaction_name(description):
    if not isinstance(description, str):
        return ''
    desc_lower = description.lower()
    for idx, name in enumerate(KNOWN_NAMES_LOWER):
        if name in desc_lower:
            return KNOWN_NAMES[idx]
    description = description.upper()
    description = re.sub(r'(\*+\d+|\d{8,}|\d{4,}\/\d+|\d{5,})', '', description)
    description = re.sub(r'WILLOW COMMERCIAL( LIMITED)?', '', description)
    description = re.sub(r'(PAYT TO|RIB:PP|PRB TRSF TO WILOW|TO WILLOW)', '', description)
    description = re.sub(
        r'A/C TO A/C TRANSFER THROUGH IBS INTERNET\s+TRANSFER FROM', '', description
    )
    description = re.sub(r'-\s*FT\s+GTL\s+WILLOW\s+COM.*$', '', description)
    description = re.sub(r'-\s*NIP$', '', description)
    description = re.sub(r'-\s*USD FOREX PURCHASE TRANSACTION$', '', description)
    description = re.sub(r'-\s*CIB$', '', description)
    description = re.sub(r'-\s*WILLOW$', '', description)
    description = re.sub(r'\bTO PAYMENT\b', '', description)
    match = re.search(r'FROM\s+[A-Z\s]+/\s*([A-Z\s\-\.]+)', description)
    if match:
        return clean_entity(match.group(1))
    match = re.search(r'TO\s+[A-Z\s]+/\s*([A-Z\s\-\.]+)', description)
    if match:
        return clean_entity(match.group(1))
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
    if (
        'charges' in desc_lower
        or 'charge' in desc_lower
        or 'offshore charges' in desc_lower
    ):
        return 'Bank Charges'
    if 'sms alert' in desc_lower:
        return 'Bank Charges'
    try:
        credit = float(str(row.get('Credit', 0)).replace(',', '').strip())
        debit = float(str(row.get('Debit', 0)).replace(',', '').strip())
    except (ValueError, TypeError):
        credit = 0.0
        debit = 0.0
    if (
        debit < 50
        or 'ELECTRONIC MONEY TRANSFER LEVY' in desc_lower
        or 'STAMP DUTY' in desc_lower
    ):
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



def process_file(uploaded_file):
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

    return df, uploaded_file.name

def export_summary(df):
    return df.groupby(['Extracted Name', 'Transaction Type']).size().unstack(fill_value=0)



if uploaded_files:
    dfs = []
    for uploaded_file in uploaded_files:
        try:
            df, filename = process_file(uploaded_file)
            dfs.append({'df': df, 'filename': filename})
            st.success(f"✅ Processed: {filename}")
            st.dataframe(df[['description', 'Extracted Name', 'Transaction Type']].head(20))

            summary = export_summary(df)
            st.subheader(f"📊 Summary: {filename}")
            st.dataframe(summary)

        except Exception as e:
            st.error(f"❌ Error processing {uploaded_file.name}: {e}")

    with st.expander("⬇ Download All Processed Files as ZIP"):
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for item in dfs:
                df = item['df']
                filename = item['filename']
                cleaned_name = f"cleaned_{filename.replace('.xlsx', '').replace('.csv', '')}.xlsx"
                # Save each DataFrame to an Excel file in memory
                excel_file = BytesIO()
                df.to_excel(excel_file, index=False)
                excel_file.seek(0)
                zf.writestr(cleaned_name, excel_file.getvalue())

        zip_buffer.seek(0)
        st.download_button(
            "Download All as ZIP",
            data=zip_buffer,
            file_name="cleaned_statements.zip",
            mime="application/zip",
        )
