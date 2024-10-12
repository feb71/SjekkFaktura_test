import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
import re

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

# Funksjon for å sjekke om en streng er numerisk
def is_numeric(value):
    try:
        float(value.replace('.', '').replace(',', '.'))
        return True
    except ValueError:
        return False

# Funksjon for å lese tabeller fra PDF-filen
def extract_data_from_pdf(file, doc_type, invoice_number=None):
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            start_reading = False

            for page in pdf.pages:
                text = page.extract_text()
                lines = text.split('\n')
                
                # Finne startpunkt for innlesning ved å identifisere kolonneoverskrifter
                for line in lines:
                    # Vi ser etter linjen med kolonneoverskriftene: "Nummer", "Beskrivelse", "Antall", "Enhet", "Pris", "Beløp"
                    if "Nummer" in line and "Beskrivelse" in line:
                        start_reading = True
                        continue
                    
                    if start_reading:
                        columns = re.split(r'\s{2,}', line)  # Splitter på store mellomrom (kolonneseparatorer)
                        
                        # Sjekk om linjen inneholder nok kolonner og om første element er et varenummer
                        if len(columns) >= 6 and columns[0].isdigit():
                            item_number = columns[0]  # Nummer = Varenummer
                            description = columns[1]  # Beskrivelse (kan være lang, så vi tar hele kolonnen)
                            
                            # Kolonnene skal inneholde Antall, Enhet, Pris, Beløp
                            try:
                                quantity = float(columns[3].replace(',', '.')) if is_numeric(columns[3]) else None  # Antall
                                unit = columns[4]  # Enhet er tekst, for eksempel "NAR"
                                unit_price = float(columns[5].replace('.', '').replace(',', '.')) if is_numeric(columns[5]) else None  # Enhetspris
                                total_price = float(columns[6].replace('.', '').replace(',', '.')) if is_numeric(columns[6]) else None  # Beløp
                            except ValueError as e:
                                st.error(f"Kunne ikke konvertere til flyttall: {e}")
                                continue
                            
                            # Lagre data i riktig format
                            unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number
                            data.append({
                                "UnikID": unique_id,
                                "Varenummer": item_number,  # Bruk "Varenummer" her for sammenslåing
                                "Beskrivelse_Faktura": description,
                                "Antall_Faktura": quantity,
                                "Enhet_Faktura": unit,  # Enhet er tekst
                                "Enhetspris_Faktura": unit_price,
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

# Funksjon for å sammenligne faktura med tilbud
def compare_invoice_offer(invoice_data, offer_data):
    # Debugging: Skriv ut kolonnenavnene fra faktura og tilbud
    st.write("Kolonner fra fakturaen:", invoice_data.columns)
    st.write("Kolonner fra tilbudet:", offer_data.columns)

    # Merge faktura og tilbud på varenummer
    try:
        merged_data = pd.merge(offer_data, invoice_data, on="Varenummer", how='outer', suffixes=('_Tilbud', '_Faktura'))
    except KeyError as e:
        st.error(f"Feil under sammenslåing: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # Konverter kolonner til numerisk for å sikre korrekt beregning
    merged_data["Antall_Faktura"] = pd.to_numeric(merged_data["Antall_Faktura"], errors='coerce')
    merged_data["Antall_Tilbud"] = pd.to_numeric(merged_data["Antall_Tilbud"], errors='coerce')
    merged_data["Enhetspris_Faktura"] = pd.to_numeric(merged_data["Enhetspris_Faktura"], errors='coerce')
    merged_data["Enhetspris_Tilbud"] = pd.to_numeric(merged_data["Enhetspris_Tilbud"], errors='coerce')

    # Beregn avvik
    merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
    merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
    
    # Prosentvis økning i enhetspris
    merged_data["Prosentvis_økning"] = ((merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]) / merged_data["Enhetspris_Tilbud"]) * 100

    # Filtrer for å vise kun avvik
    avvik = merged_data[(merged_data["Avvik_Antall"].notna() & (merged_data["Avvik_Antall"] != 0)) |
                        (merged_data["Avvik_Enhetspris"].notna() & (merged_data["Avvik_Enhetspris"] != 0))]
    
    # Varenummer som ikke finnes i tilbud
    only_in_invoice = merged_data[merged_data['Enhetspris_Tilbud'].isna()]
    
    return avvik, only_in_invoice, merged_data

# Hovedfunksjon for Streamlit-appen
def main():
    st.title("Sammenlign Faktura mot Tilbud")

    # Opplasting av filer
    invoice_file = st.file_uploader("Last opp faktura fra Brødrene Dahl", type="pdf")
    offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl (Excel)", type="xlsx")

    if invoice_file and offer_file:
        # Hent fakturanummer
        invoice_number = get_invoice_number(invoice_file)

        if invoice_number:
            # Ekstraher data fra PDF-filer
            invoice_data = extract_data_from_pdf(invoice_file, "Faktura", invoice_number)

            # Les tilbudsdata
            offer_data = pd.read_excel(offer_file)
            offer_data.rename(columns={
                'VARENR': 'Varenummer',  # Endrer VARENR til Varenummer for å matche fakturaen
                'BESKRIVELSE': 'Beskrivelse_Tilbud',
                'ANTALL': 'Antall_Tilbud',
                'ENHET': 'Enhet_Tilbud',
                'ENHETSPRIS': 'Enhetspris_Tilbud',
                'TOTALPRIS': 'Totalt pris_Tilbud'
            }, inplace=True)

            # Sammenlign faktura med tilbud
            avvik, only_in_invoice, merged_data = compare_invoice_offer(invoice_data, offer_data)

            # Vis avvik
            st.subheader("Avvik mellom Faktura og Tilbud")
            st.dataframe(avvik)

            # Vis varenummer som kun finnes i faktura
            st.subheader("Varenummer som kun finnes i faktura")
            st.dataframe(only_in_invoice)

            # Last ned avviksrapport som Excel
            def convert_df_to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                return output.getvalue()

            st.download_button(
                label="Last ned avviksrapport som Excel",
                data=convert_df_to_excel(avvik),
                file_name="avvik_rapport.xlsx"
            )

            st.download_button(
                label="Last ned varer kun i faktura",
                data=convert_df_to_excel(only_in_invoice),
                file_name="varer_kun_i_faktura.xlsx"
            )
        else:
            st.error("Fakturanummeret ble ikke funnet i PDF-filen.")
    else:
        st.info("Vennligst last opp både faktura og tilbud for sammenligning.")

if __name__ == "__main__":
    main()
