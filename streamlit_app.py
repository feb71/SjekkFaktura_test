import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
import re

# Funksjon for å lese fakturanummer fra PDF-filer ved å søke gjennom alle sidene
def get_invoice_number_from_pdf(file):
    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                # Bruk regex for å finne fakturanummeret
                match = re.search(r"Fakturanummer\s*[:\-]?\s*(\d+)", text, re.IGNORECASE)
                if match:
                    return match.group(1)
        return None
    except Exception as e:
        st.error(f"Kunne ikke lese fakturanummer fra PDF: {e}")
        return None

# Funksjon for å sjekke om en streng er numerisk
def is_numeric(value):
    try:
        float(value.replace('.', '').replace(',', '.'))
        return True
    except ValueError:
        return False

# Funksjon for å lese oppsummeringstabellen fra PDF-filer
def extract_summary_data_from_pdf(file, doc_type, invoice_number=None):
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            for page in pdf.pages:
                text = page.extract_text()
                lines = text.split('\n')
                
                for line in lines:
                    # Splitter på store mellomrom (kolonneseparatorer)
                    columns = re.split(r'\s{2,}', line)  
                    
                    # Sjekk om linjen inneholder nok kolonner og om første element er et varenummer
                    if len(columns) >= 6 and columns[0].isdigit():  # Sjekk at første kolonne er et nummer (Artikkel)
                        item_number = columns[0]  # Artikkel = Nummer
                        description = columns[1]  # Beskrivelse
                        
                        # Kolonnene skal inneholde Antall, Enhet, Totalpris
                        try:
                            quantity = float(columns[3].replace(',', '.')) if is_numeric(columns[3]) else None  # Antall
                            total_price = float(columns[5].replace('.', '').replace(',', '.')) if is_numeric(columns[5]) else None  # Beløp
                        except ValueError as e:
                            st.error(f"Kunne ikke konvertere til flyttall: {e}")
                            continue
                        
                        # Lagre data i riktig format
                        unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number
                        data.append({
                            "UnikID": unique_id,
                            "Artikkel": item_number,  # Artikkel fra oppsummeringen
                            "Beskrivelse": description,
                            "Antall": quantity,
                            "Totalt pris": total_price,
                            "Type": doc_type
                        })

            # Hvis ingen data ble funnet
            if len(data) == 0:
                st.error("Ingen data ble funnet i PDF-filen.")
                
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

# Funksjon for å lese detaljtabellen og hente enhetspriser
def extract_detail_data_from_pdf(file, doc_type):
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            for page in pdf.pages:
                text = page.extract_text()
                lines = text.split('\n')
                
                for line in lines:
                    # Splitter på store mellomrom (kolonneseparatorer)
                    columns = re.split(r'\s{2,}', line)  
                    
                    # Sjekk om linjen inneholder nok kolonner og om første element er et varenummer
                    if len(columns) >= 6 and columns[0].isdigit():  # Sjekk at første kolonne er et nummer (Nummer)
                        item_number = columns[0]  # Nummer fra detaljtabellen
                        
                        # Kolonnene skal inneholde Enhetspris
                        try:
                            unit_price = float(columns[5].replace('.', '').replace(',', '.')) if is_numeric(columns[5]) else None  # Enhetspris
                        except ValueError as e:
                            st.error(f"Kunne ikke konvertere til flyttall: {e}")
                            continue
                        
                        # Lagre data i riktig format
                        data.append({
                            "Nummer": item_number,  # Nummer fra detaljtabellen
                            "Enhetspris": unit_price
                        })

            # Hvis ingen data ble funnet
            if len(data) == 0:
                st.error("Ingen enhetspriser ble funnet i PDF-filen.")
                
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

# Funksjon for å sammenligne faktura med tilbud
def compare_invoice_offer(summary_data, detail_data):
    # Feilsøking: Vis kolonnenavnene før sammenslåing
    st.write("Kolonnenavn i oppsummeringen:", summary_data.columns)
    st.write("Kolonnenavn i detaljtabellen:", detail_data.columns)
    
    # Merge oppsummering (Artikkel) med detaljer (Nummer) på varenummer
    try:
        merged_data = pd.merge(summary_data, detail_data, left_on="Artikkel", right_on="Nummer", how='left')
    except KeyError as e:
        st.error(f"Feil under sammenslåing: {e}")
        return pd.DataFrame()

    # Beregn enhetsprisen hvis den ikke finnes
    merged_data["Enhetspris_utregnet"] = merged_data["Totalt pris"] / merged_data["Antall"]

    return merged_data

# Hovedfunksjon for Streamlit-appen
def main():
    st.title("Sammenlign Faktura mot Tilbud")

    # Opplasting av filer
    invoice_file = st.file_uploader("Last opp faktura fra Brødrene Dahl", type="pdf")
    offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl (Excel)", type="xlsx")

    if invoice_file:
        # Hent fakturanummer fra PDF
        invoice_number = get_invoice_number_from_pdf(invoice_file)

        if invoice_number:
            # Ekstraher oppsummering og detaljtabeller fra PDF
            summary_data = extract_summary_data_from_pdf(invoice_file, "Oppsummering", invoice_number)
            detail_data = extract_detail_data_from_pdf(invoice_file, "Detaljer")

            # Sammenlign oppsummering og detaljer for å hente enhetspriser
            result = compare_invoice_offer(summary_data, detail_data)

            # Vis resultatet
            st.subheader("Sammenslått data fra fakturaen")
            st.dataframe(result)

            # Last ned rapport som Excel
            def convert_df_to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                return output.getvalue()

            # Last ned-knapp for Excel-filen
            st.download_button(
                label="Last ned sammenslått data som Excel",
                data=convert_df_to_excel(result),
                file_name="faktura_data.xlsx"
            )
        else:
            st.error("Fakturanummeret ble ikke funnet i PDF-filen.")

if __name__ == "__main__":
    main()
