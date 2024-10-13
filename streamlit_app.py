import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Streamlit App", layout="wide", initial_sidebar_state="expanded")

# Funksjon for å lese fakturanummer fra PDF
def get_invoice_number(file):
    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                match = re.search(r"Fakturanummer\s*[:\-]?\s*(\d+)", text, re.IGNORECASE)
                if match:
                    return match.group(1)
        return None
    except Exception as e:
        st.error(f"Kunne ikke lese fakturanummer fra PDF: {e}")
        return None

# Funksjon for å lese faktura og returnere data uten rabatt (for avvikstabellen)
def extract_data_for_avvik(file, doc_type, invoice_number=None):
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            start_reading = False

            for page in pdf.pages:
                text = page.extract_text()
                if text is None:
                    st.error(f"Ingen tekst funnet på side {page.page_number} i PDF-filen.")
                    continue
                
                lines = text.split('\n')
                for line in lines:
                    if doc_type == "Faktura" and "Artikkel" in line:
                        start_reading = True
                        continue

                    if start_reading:
                        columns = line.split()
                        if len(columns) >= 5:
                            item_number = columns[1]
                            if not item_number.isdigit():
                                continue

                            description = " ".join(columns[2:-3])
                            try:
                                quantity = columns[-3]
                                unit_price = float(columns[-2].replace('.', '').replace(',', '.')) if columns[-2].replace('.', '').replace(',', '').isdigit() else None
                                total_price = float(columns[-1].replace('.', '').replace(',', '.')) if columns[-1].replace('.', '').replace(',', '').isdigit() else None
                            except ValueError as e:
                                st.error(f"Kunne ikke konvertere til flyttall: {e}")
                                continue

                            unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number
                            data.append({
                                "UnikID": unique_id,
                                "Varenummer": item_number,
                                "Faktura_Beskrivelse": description,
                                "Faktura_Antall": quantity,
                                "Faktura_Enhetspris": unit_price,
                                "Faktura_Totalt_pris": total_price,
                                "Type": doc_type
                            })
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

# Funksjon for å lese faktura med rabatt (for manglende tilbudsartikler)
def extract_data_with_rabatt(file, doc_type, invoice_number=None):
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            start_reading = False

            for page in pdf.pages:
                text = page.extract_text()
                if text is None:
                    st.error(f"Ingen tekst funnet på side {page.page_number} i PDF-filen.")
                    continue
                
                lines = text.split('\n')
                for line in lines:
                    if doc_type == "Faktura" and "Artikkel" in line:
                        start_reading = True
                        continue

                    if start_reading:
                        columns = line.split()
                        if len(columns) >= 6:
                            item_number = columns[1]
                            if not item_number.isdigit():
                                continue

                            description = " ".join(columns[2:-4])
                            try:
                                quantity = columns[-4]
                                discount = columns[-3]
                                unit_price = float(columns[-2].replace('.', '').replace(',', '.')) if columns[-2].replace('.', '').replace(',', '').isdigit() else None
                                total_price = float(columns[-1].replace('.', '').replace(',', '.')) if columns[-1].replace('.', '').replace(',', '').isdigit() else None
                            except ValueError as e:
                                st.error(f"Kunne ikke konvertere til flyttall: {e}")
                                continue

                            unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number
                            data.append({
                                "UnikID": unique_id,
                                "Varenummer": item_number,
                                "Faktura_Beskrivelse": description,
                                "Faktura_Antall": quantity,
                                "Rabatt": discount,
                                "Faktura_Enhetspris": unit_price,
                                "Faktura_Totalt_pris": total_price,
                                "Type": doc_type
                            })
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

# Hovedfunksjon for Streamlit-appen
def main():
    st.title("Sammenlign Faktura mot Tilbud")
    
    # Opprett tre kolonner
    col1, col2, col3 = st.columns([1, 5, 1])

    with col1:
        st.header("Last opp filer")
        invoice_file = st.file_uploader("Last opp faktura fra Brødrene Dahl", type="pdf")
        offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl (Excel)", type="xlsx")

    if invoice_file and offer_file:
        # Hent fakturanummer
        invoice_number = get_invoice_number(invoice_file)

        if invoice_number:
            st.success(f"Fakturanummer funnet: {invoice_number}")
            
            # Ekstraher data for avvikstabellen uten rabatt
            invoice_data_avvik = extract_data_for_avvik(invoice_file, "Faktura", invoice_number)

            # Ekstraher data med rabatt for manglende artikler
            invoice_data_rabatt = extract_data_with_rabatt(invoice_file, "Faktura", invoice_number)

            # Les tilbudet fra Excel-filen
            offer_data = pd.read_excel(offer_file)

            # Riktige kolonnenavn fra Excel-filen for tilbud
            offer_data.rename(columns={
                'VARENR': 'Varenummer',
                'BESKRIVELSE': 'Tilbud_Beskrivelse',
                'ANTALL': 'Tilbud_Antall',
                'ENHET': 'Tilbud_Enhet',
                'ENHETSPRIS': 'Tilbud_Enhetspris',
                'TOTALPRIS': 'Tilbud_Totalt_pris'
            }, inplace=True)

            # Sammenligne faktura mot tilbud (avvikstabell)
            merged_data = pd.merge(offer_data, invoice_data_avvik, on="Varenummer", how='inner', suffixes=('_Tilbud', '_Faktura'))

            # Finne avvik
            merged_data["Avvik_Antall"] = merged_data["Faktura_Antall"] - merged_data["Tilbud_Antall"]
            merged_data["Avvik_Enhetspris"] = merged_data["Faktura_Enhetspris"] - merged_data["Tilbud_Enhetspris"]
            merged_data["Prosentvis_økning"] = ((merged_data["Faktura_Enhetspris"] - merged_data["Tilbud_Enhetspris"]) / merged_data["Tilbud_Enhetspris"]) * 100

            st.subheader("Avvik mellom Faktura og Tilbud")
            st.dataframe(merged_data)

            # Artikler som finnes i faktura, men ikke i tilbud (med rabatt)
            unmatched_items = pd.merge(offer_data, invoice_data_rabatt, on="Varenummer", how="outer", indicator=True)
            only_in_invoice = unmatched_items[unmatched_items['_merge'] == 'right_only'][["Varenummer", "Faktura_Beskrivelse", "Faktura_Antall", "Faktura_Enhetspris", "Faktura_Totalt_pris", "Rabatt"]]
            
            st.subheader("Varenummer som finnes i faktura, men ikke i tilbud")
            st.dataframe(only_in_invoice)

if __name__ == "__main__":
    main()
